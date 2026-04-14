using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using Autodesk.Revit.UI.Selection;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Reflection;
using View = Autodesk.Revit.DB.View;

namespace AngularDIMParam
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
                "AngularDIMParam",
                "Angular\nDIM\nParam",
                assemblyPath,
                "AngularDIMParam.Class1"
            );

            panel.AddItem(btnData);
            return Result.Succeeded;
        }

        public Result OnShutdown(UIControlledApplication app) => Result.Succeeded;
    }

    [Autodesk.Revit.Attributes.Transaction(Autodesk.Revit.Attributes.TransactionMode.Manual)]
    public class Class1 : IExternalCommand
    {
        public Result Execute(ExternalCommandData commandData, ref string message, ElementSet elements)
        {
            UIDocument uiDoc = commandData.Application.ActiveUIDocument;
            Document doc = uiDoc.Document;

            try
            {
                Reference r1 = uiDoc.Selection.PickObject(ObjectType.Edge, "Chọn cạnh 1");
                Reference r2 = uiDoc.Selection.PickObject(ObjectType.Edge, "Chọn cạnh 2");

                Line line1 = GetLineFromReference(doc, r1);
                Line line2 = GetLineFromReference(doc, r2);

                if (line1 == null || line2 == null)
                {
                    message = "Không lấy được line.";
                    return Result.Failed;
                }

                using (Transaction tx = new Transaction(doc, "Angular DIM 3D"))
                {
                    tx.Start();

                    XYZ mid1 = (line1.GetEndPoint(0) + line1.GetEndPoint(1)) * 0.5;
                    XYZ mid2 = (line2.GetEndPoint(0) + line2.GetEndPoint(1)) * 0.5;

                    XYZ dir1 = line1.Direction.Normalize();
                    XYZ dir2 = line2.Direction.Normalize();

                    XYZ normal = dir1.CrossProduct(dir2);

                    bool isParallel = normal.GetLength() < 1e-6;

                    if (isParallel)
                    {
                        // 🔥 Handle song song
                        XYZ connect = (mid2 - mid1).Normalize();
                        normal = dir1.CrossProduct(connect).Normalize();
                    }
                    else
                    {
                        normal = normal.Normalize();
                    }

                    // 🔥 Tạo plane riêng (không dùng view nữa)
                    Plane plane = Plane.CreateByNormalAndOrigin(normal, mid1);

                    Line l1 = ProjectLineToPlane(line1, plane);
                    Line l2 = ProjectLineToPlane(line2, plane);

                    // 🔥 ép thành unbound
                    l1 = Line.CreateUnbound(mid1, l1.Direction);
                    l2 = Line.CreateUnbound(mid2, l2.Direction);

                    XYZ center = mid1;

                    XYZ v1 = l1.Direction.Normalize();
                    XYZ v2 = l2.Direction.Normalize();

                    double angle = v1.AngleTo(v2);

                    // 🔥 nếu song song → fake góc nhỏ
                    if (angle < 1e-6)
                        angle = Math.PI / 180; // 1 độ cho DIM hiển thị

                    XYZ xAxis = v1;
                    XYZ yAxis = normal.CrossProduct(xAxis).Normalize();

                    double radius = mid1.DistanceTo(mid2) * 0.5;
                    if (radius < 1) radius = 2;

                    Arc arc = Arc.Create(center, radius, 0, angle, xAxis, yAxis);

                    View view = doc.ActiveView;

                    if (doc.IsFamilyDocument)
                    {
                        Dimension dim = doc.FamilyCreate.NewAngularDimension(view, arc, r1, r2);

                        var param = doc.FamilyManager.Parameters
                            .Cast<FamilyParameter>()
                            .FirstOrDefault(p => p.Definition.Name == "Angle_Param");

                        if (param != null)
                            dim.FamilyLabel = param;
                    }
                    else
                    {
                        DimensionType type = new FilteredElementCollector(doc)
                            .OfClass(typeof(DimensionType))
                            .Cast<DimensionType>()
                            .FirstOrDefault(x => x.StyleType == DimensionStyleType.Angular);

                        AngularDimension.Create(doc, view, arc, new List<Reference> { r1, r2 }, type);
                    }

                    tx.Commit();
                }

                return Result.Succeeded;
            }
            catch (Exception ex)
            {
                message = ex.Message;
                return Result.Failed;
            }
        }

        // ================= HELPER =================

        private Line GetLineFromReference(Document doc, Reference r)
        {
            if (r == null) return null;

            Element el = doc.GetElement(r);
            if (el == null) return null;

            GeometryObject geoObj = el.GetGeometryObjectFromReference(r);

            if (geoObj is Edge edge)
            {
                Curve c = edge.AsCurve();
                if (c is Line line)
                    return line;

                // nếu là curve khác (arc, spline) → convert tạm thành line
                return Line.CreateBound(c.GetEndPoint(0), c.GetEndPoint(1));
            }

            return null;
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