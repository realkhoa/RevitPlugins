using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using Autodesk.Revit.UI.Selection;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Reflection;

namespace AngularDIM
{
    public class App : IExternalApplication
    {
        public Result OnStartup(UIControlledApplication app)
        {
            string tabName = "DIM";

            try { app.CreateRibbonTab(tabName); } catch { }

            RibbonPanel panel = app.GetRibbonPanels(tabName)
                .FirstOrDefault(p => p.Name == "Dimension")
                ?? app.CreateRibbonPanel(tabName, "Dimension");

            string assemblyPath = Assembly.GetExecutingAssembly().Location;

            PushButtonData btnData = new PushButtonData(
                "AngularDIM",
                "Angular\nDIM",
                assemblyPath,
                "AngularDIM.Command"
            );

            if (!panel.GetItems().OfType<PushButton>().Any(b => b.Name == "AngularDIM"))
            {
                PushButton button = panel.AddItem(btnData) as PushButton;
                button.ToolTip = "DIM góc chuẩn giữa 2 edge 🔥 (Project + Family)";
            }

            return Result.Succeeded;
        }

        public Result OnShutdown(UIControlledApplication app)
        {
            return Result.Succeeded;
        }
    }

    [Autodesk.Revit.Attributes.Transaction(Autodesk.Revit.Attributes.TransactionMode.Manual)]
    public class Command : IExternalCommand
    {
        public Result Execute(ExternalCommandData commandData, ref string message, ElementSet elements)
        {
            UIApplication uiApp = commandData.Application;
            UIDocument uiDoc = uiApp.ActiveUIDocument;
            Document doc = uiDoc.Document;

            try
            {
                Reference r1 = uiDoc.Selection.PickObject(ObjectType.Edge, "Chọn cạnh thứ 1");
                Reference r2 = uiDoc.Selection.PickObject(ObjectType.Edge, "Chọn cạnh thứ 2");

                if (r1 == null || r2 == null)
                    return Result.Cancelled;

                Line line1 = GetLineFromReference(doc, r1);
                Line line2 = GetLineFromReference(doc, r2);

                if (line1 == null || line2 == null)
                {
                    message = "Không lấy được line từ edge.";
                    return Result.Failed;
                }

                using (Transaction tx = new Transaction(doc, "Angular DIM"))
                {
                    tx.Start();

                    View view = doc.ActiveView;

                    // 🔥 Lấy mặt phẳng view
                    Plane plane = GetViewPlane(view);

                    // 🔥 Project line về mặt phẳng view
                    Line l1 = ProjectLineToPlane(line1, plane);
                    Line l2 = ProjectLineToPlane(line2, plane);

                    XYZ mid1 = (l1.GetEndPoint(0) + l1.GetEndPoint(1)) * 0.5;
                    XYZ mid2 = (l2.GetEndPoint(0) + l2.GetEndPoint(1)) * 0.5;

                    l1.MakeUnbound();
                    l2.MakeUnbound();

                    // 🔥 Tìm giao điểm
                    IntersectionResultArray ira;
                    var result = l1.Intersect(l2, out ira);

                    if (result != SetComparisonResult.Overlap || ira == null || ira.IsEmpty)
                    {
                        message = "2 cạnh không giao nhau trong view (có thể song song hoặc lệch mặt phẳng).";
                        tx.RollBack();
                        return Result.Failed;
                    }

                    XYZ center = ira.get_Item(0).XYZPoint;

                    XYZ v1 = (mid1 - center).Normalize();
                    XYZ v2 = (mid2 - center).Normalize();

                    double angle = v1.AngleTo(v2);

                    if (angle < 1e-6)
                    {
                        message = "2 cạnh gần như song song.";
                        tx.RollBack();
                        return Result.Failed;
                    }

                    // 🔥 vector pháp tuyến theo view
                    XYZ normal = plane.Normal;

                    XYZ xAxis = v1;
                    XYZ yAxis = normal.CrossProduct(xAxis).Normalize();

                    double r1Len = mid1.DistanceTo(center);
                    double r2Len = mid2.DistanceTo(center);
                    double radius = Math.Min(r1Len, r2Len) * 0.5;

                    if (radius < 1)
                        radius = 2;

                    Arc arc = Arc.Create(center, radius, 0, angle, xAxis, yAxis);

                    if (doc.IsFamilyDocument)
                    {
                        // 🔥 FAMILY
                        doc.FamilyCreate.NewAngularDimension(view, arc, r1, r2);
                    }
                    else
                    {
                        // 🔥 PROJECT
                        DimensionType angDimType = new FilteredElementCollector(doc)
                            .OfClass(typeof(DimensionType))
                            .Cast<DimensionType>()
                            .FirstOrDefault(x => x.StyleType == DimensionStyleType.Angular);

                        if (angDimType == null)
                        {
                            message = "Không có Angular Dimension Type.";
                            tx.RollBack();
                            return Result.Failed;
                        }

                        IList<Reference> refs = new List<Reference> { r1, r2 };

                        AngularDimension.Create(
                            doc,
                            view,
                            arc,
                            refs,
                            angDimType
                        );
                    }

                    tx.Commit();
                }

                return Result.Succeeded;
            }
            catch (Autodesk.Revit.Exceptions.OperationCanceledException)
            {
                return Result.Cancelled;
            }
            catch (Exception ex)
            {
                message = ex.Message;
                return Result.Failed;
            }
        }

        // ================= CORE =================

        private Line GetLineFromReference(Document doc, Reference r)
        {
            Element el = doc.GetElement(r);

            Options opt = new Options();
            opt.ComputeReferences = true;
            opt.IncludeNonVisibleObjects = true;

            GeometryElement geo = el.get_Geometry(opt);

            foreach (GeometryObject obj in geo)
            {
                // 🔥 CASE 1: Solid trực tiếp
                if (obj is Solid solid)
                {
                    Line line = FindLineInSolid(solid, r);
                    if (line != null) return line;
                }

                // 🔥 CASE 2: Geometry Instance (family, link...)
                if (obj is GeometryInstance inst)
                {
                    GeometryElement instGeo = inst.GetInstanceGeometry();

                    foreach (GeometryObject iobj in instGeo)
                    {
                        if (iobj is Solid solid2)
                        {
                            Line line = FindLineInSolid(solid2, r);
                            if (line != null)
                            {
                                // 🔥 apply transform
                                return line.CreateTransformed(inst.Transform) as Line;
                            }
                        }
                    }
                }
            }

            return null;
        }

        private Line FindLineInSolid(Solid solid, Reference targetRef)
        {
            foreach (Edge edge in solid.Edges)
            {
                if (edge.Reference != null && edge.Reference.ElementId == targetRef.ElementId)
                {
                    Curve c = edge.AsCurve();
                    if (c is Line line)
                        return line;
                }
            }
            return null;
        }

        private Plane GetViewPlane(View view)
        {
            return Plane.CreateByNormalAndOrigin(view.ViewDirection, view.Origin);
        }

        private Line ProjectLineToPlane(Line line, Plane plane)
        {
            XYZ p1 = ProjectPoint(line.GetEndPoint(0), plane);
            XYZ p2 = ProjectPoint(line.GetEndPoint(1), plane);
            return Line.CreateBound(p1, p2);
        }

        private XYZ ProjectPoint(XYZ p, Plane plane)
        {
            XYZ v = p - plane.Origin;
            double dist = v.DotProduct(plane.Normal);
            return p - dist * plane.Normal;
        }
    }
}