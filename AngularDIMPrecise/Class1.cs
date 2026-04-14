using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using Autodesk.Revit.UI.Selection;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Reflection;
using System.Reflection.Metadata;
using System.Transactions;
using System.Xml.Linq;
using static System.Net.Mime.MediaTypeNames;

namespace AngularDIMPrecise
{
    // ================= APP (RIBBON) =================
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
                "AngularDIMPrecise",
                "Angular\nDIM\nPrecise",
                assemblyPath,
                "AngularDIMPrecise.Command"
            );

            if (!panel.GetItems().OfType<PushButton>().Any(b => b.Name == "AngularDIMPrecise"))
            {
                PushButton button = panel.AddItem(btnData) as PushButton;
                button.ToolTip = "DIM góc chính xác giữa 2 wall hoặc line 🔥";
            }

            return Result.Succeeded;
        }

        public Result OnShutdown(UIControlledApplication app)
        {
            return Result.Succeeded;
        }
    }

    // ================= COMMAND =================
    [Autodesk.Revit.Attributes.Transaction(Autodesk.Revit.Attributes.TransactionMode.Manual)]
    public class Command : IExternalCommand
    {
        public Result Execute(ExternalCommandData commandData, ref string message, ElementSet elements)
        {
            UIApplication uiApp = commandData.Application;
            UIDocument uiDoc = uiApp.ActiveUIDocument;
            Autodesk.Revit.DB.Document doc = uiDoc.Document;

            try
            {
                Reference pick1 = uiDoc.Selection.PickObject(
                    ObjectType.Element,
                    "Chọn đối tượng thứ 1");

                Reference pick2 = uiDoc.Selection.PickObject(
                    ObjectType.Element,
                    "Chọn đối tượng thứ 2");

                if (pick1 == null || pick2 == null) return Result.Cancelled;

                IList<Reference> pickedRefs = new List<Reference> { pick1, pick2 };

                if (pickedRefs.Count != 2) return Result.Cancelled;

                IList<Reference> validRefs = new List<Reference>();
                List<Line> lines = new List<Line>();

                foreach (Reference r in pickedRefs)
                {
                    Element el = doc.GetElement(r);

                    if (ExtractData(el, doc, out Reference refObj, out Line geomLine))
                    {
                        validRefs.Add(refObj);
                        lines.Add(geomLine);
                    }
                }

                if (validRefs.Count < 2)
                {
                    message = "Không lấy được reference hợp lệ.";
                    return Result.Failed;
                }

                DimensionType angDimType = new FilteredElementCollector(doc)
                    .OfClass(typeof(DimensionType))
                    .Cast<DimensionType>()
                    .FirstOrDefault(x => x.StyleType == DimensionStyleType.Angular);

                if (angDimType == null)
                {
                    message = "Không có Angular Dimension Type.";
                    return Result.Failed;
                }

                using (Autodesk.Revit.DB.Transaction tx = new Autodesk.Revit.DB.Transaction(doc, "Angular DIM Precise"))
                {
                    tx.Start();

                    View view = doc.ActiveView;

                    Line l1 = FlattenLine(lines[0]);
                    Line l2 = FlattenLine(lines[1]);

                    l1.MakeUnbound();
                    l2.MakeUnbound();

                    IntersectionResultArray ira;
                    var result = l1.Intersect(l2, out ira);

                    if (result != SetComparisonResult.Overlap || ira == null || ira.IsEmpty)
                    {
                        message = "2 đối tượng song song, không tạo được góc.";
                        tx.RollBack();
                        return Result.Failed;
                    }

                    XYZ center = ira.get_Item(0).XYZPoint;

                    XYZ mid1 = FlattenPoint(lines[0].Evaluate(0.5, true));
                    XYZ mid2 = FlattenPoint(lines[1].Evaluate(0.5, true));

                    XYZ v1 = (mid1 - center).Normalize();
                    XYZ v2 = (mid2 - center).Normalize();

                    XYZ normal = v1.CrossProduct(v2).Normalize();
                    if (normal.Z < 0) normal = normal.Negate();

                    XYZ xAxis = v1;
                    XYZ yAxis = normal.CrossProduct(xAxis).Normalize();

                    double angle = v1.AngleTo(v2);

                    double r1 = mid1.DistanceTo(center);
                    double r2 = mid2.DistanceTo(center);
                    double radius = Math.Min(r1, r2) * 0.5;

                    if (radius < 2) radius = 3;

                    Arc arc = Arc.Create(center, radius, 0, angle, xAxis, yAxis);

                    AngularDimension.Create(doc, view, arc, validRefs, angDimType);

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

        private bool ExtractData(Autodesk.Revit.DB.Element el, Autodesk.Revit.DB.Document doc, out Autodesk.Revit.DB.Reference refObj, out Autodesk.Revit.DB.Line geomLine)
        {
            refObj = null;
            geomLine = null;

            if (el is Wall wall)
            {
                LocationCurve lc = wall.Location as LocationCurve;
                if (lc?.Curve is Line l)
                    geomLine = l;

                var faces = HostObjectUtils.GetSideFaces(wall, ShellLayerType.Exterior);
                if (faces.Count > 0)
                    refObj = faces[0];

                return refObj != null && geomLine != null;
            }

            if (el is CurveElement ce && ce.GeometryCurve is Line line)
            {
                geomLine = line;

                Options opt = new Options
                {
                    ComputeReferences = true,
                    View = doc.ActiveView
                };

                var geo = ce.get_Geometry(opt);

                foreach (GeometryObject obj in geo)
                {
                    if (obj is Line l && l.Reference != null)
                    {
                        refObj = l.Reference;
                        break;
                    }
                }

                return refObj != null;
            }

            return false;
        }

        private Line FlattenLine(Line line)
        {
            XYZ p1 = FlattenPoint(line.GetEndPoint(0));
            XYZ p2 = FlattenPoint(line.GetEndPoint(1));
            return Line.CreateBound(p1, p2);
        }

        private XYZ FlattenPoint(XYZ p)
        {
            return new XYZ(p.X, p.Y, 0);
        }
    }
}