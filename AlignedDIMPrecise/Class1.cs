using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using Autodesk.Revit.UI.Selection;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Reflection;

namespace AlignedDIMPrecise
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
                "AlignedDIMPrecise",
                "Aligned\nDIM\nPrecise",
                assemblyPath,
                "AlignedDIMPrecise.Command"
            );

            if (!panel.GetItems().OfType<PushButton>().Any(b => b.Name == "AlignedDIMPrecise"))
            {
                PushButton button = panel.AddItem(btnData) as PushButton;
                button.ToolTip = "1 click: DIM thông minh (wall length / distance)";
            }

            return Result.Succeeded;
        }

        public Result OnShutdown(UIControlledApplication app) => Result.Succeeded;
    }

    [Autodesk.Revit.Attributes.Transaction(Autodesk.Revit.Attributes.TransactionMode.Manual)]
    public class Command : IExternalCommand
    {
        public Result Execute(ExternalCommandData commandData, ref string message, ElementSet elements)
        {
            UIDocument uiDoc = commandData.Application.ActiveUIDocument;
            Document doc = uiDoc.Document;

            try
            {
                // Pick 1
                Reference pick1 = uiDoc.Selection.PickObject(
                    ObjectType.PointOnElement,
                    "Pick object 1");

                if (pick1 == null) return Result.Cancelled;

                // Pick 2 (ESC = bỏ qua)
                Reference pick2 = null;

                try
                {
                    pick2 = uiDoc.Selection.PickObject(
                        ObjectType.PointOnElement,
                        "Pick object 2 (ESC = đo chiều dài tường)");
                }
                catch
                {
                    // ESC → chỉ có 1 object
                }

                using (Transaction tx = new Transaction(doc, "Smart DIM"))
                {
                    tx.Start();

                    View view = doc.ActiveView;

                    // =========================
                    // 🎯 CASE 1: 1 object (Wall length)
                    // =========================
                    if (pick2 == null)
                    {
                        Element el = doc.GetElement(pick1);

                        if (el is Wall wall)
                        {
                            LocationCurve lc = wall.Location as LocationCurve;

                            if (lc?.Curve is Line line)
                            {
                                XYZ p1 = line.GetEndPoint(0);
                                XYZ p2 = line.GetEndPoint(1);

                                Reference ref1 = GetClosestEdgeReference(wall, p1, view);
                                Reference ref2 = GetClosestEdgeReference(wall, p2, view);

                                if (ref1 == null || ref2 == null)
                                {
                                    message = "Không lấy được reference.";
                                    tx.RollBack();
                                    return Result.Failed;
                                }

                                ReferenceArray refArray = new ReferenceArray();
                                refArray.Append(ref1);
                                refArray.Append(ref2);

                                // =========================
                                // OFFSET DIM LINE
                                // =========================

                                // direction của tường
                                XYZ lineDir = (p2 - p1).Normalize();

                                // hướng nhìn của view
                                XYZ viewDir = view.ViewDirection;

                                // vector vuông góc để offset
                                XYZ offsetDir = lineDir.CrossProduct(viewDir).Normalize();

                                // khoảng cách offset (300mm)
                                double offset = UnitUtils.ConvertToInternalUnits(300, UnitTypeId.Millimeters);

                                // đảm bảo offset ra phía "nhìn thấy"
                                if (offsetDir.Z < 0)
                                {
                                    offsetDir = -offsetDir;
                                }

                                // dịch 2 điểm ra ngoài
                                XYZ p1Offset = p1 + offsetDir * offset;
                                XYZ p2Offset = p2 + offsetDir * offset;

                                // tạo dim line mới
                                Line dimLine = Line.CreateBound(p1Offset, p2Offset);

                                doc.Create.NewDimension(view, dimLine, refArray);
                            }
                        }
                    }

                    // =========================
                    // 🎯 CASE 2: 2 object (Distance)
                    // =========================
                    else
                    {
                        List<Reference> picked = new List<Reference> { pick1, pick2 };

                        ReferenceArray refArray = new ReferenceArray();
                        List<XYZ> pts = new List<XYZ>();
                        List<Line> lines = new List<Line>();

                        foreach (Reference r in picked)
                        {
                            Element el = doc.GetElement(r);

                            Reference bestRef = r;

                            // Nếu reference KHÔNG usable thì mới fallback
                            if (!IsValidReferenceForDimension(r))
                            {
                                if (r.GlobalPoint != null)
                                {
                                    var closeRef = GetClosestEdgeReference(el, r.GlobalPoint, view);
                                    if (closeRef != null)
                                        bestRef = closeRef;
                                }
                            }

                            refArray.Append(bestRef);

                            if (TryGetLine(el, out Line l))
                            {
                                lines.Add(l);
                                pts.Add(l.Evaluate(0.5, true));
                            }
                            else if (r.GlobalPoint != null)
                            {
                                pts.Add(r.GlobalPoint);
                            }
                        }

                        if (pts.Count < 2)
                        {
                            message = "Không lấy được dữ liệu.";
                            tx.RollBack();
                            return Result.Failed;
                        }

                        XYZ lineDir = (lines[0].GetEndPoint(1) - lines[0].GetEndPoint(0)).Normalize();
                        XYZ viewDir = view.ViewDirection;
                        XYZ dimDir = lineDir.CrossProduct(viewDir).Normalize();

                        XYZ mid1 = pts[0];
                        XYZ mid2 = pts[1];

                        double dist = (mid2 - mid1).DotProduct(dimDir);

                        if (Math.Abs(dist) < 1e-6)
                        {
                            message = "Hai đối tượng trùng phương.";
                            tx.RollBack();
                            return Result.Failed;
                        }

                        XYZ basePt = mid1 + dimDir * 2.0;

                        Line dimLine = Line.CreateBound(basePt, basePt + dimDir * dist);

                        doc.Create.NewDimension(view, dimLine, refArray);
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

        private bool TryGetLine(Element el, out Line line)
        {
            line = null;

            if (el is Wall wall)
            {
                LocationCurve lc = wall.Location as LocationCurve;
                if (lc?.Curve is Line l)
                {
                    line = l;
                    return true;
                }
            }

            if (el is CurveElement ce && ce.GeometryCurve is Line l2)
            {
                line = l2;
                return true;
            }

            return false;
        }

        private Reference GetClosestEdgeReference(Element el, XYZ pickPoint, View view)
        {
            Options opt = new Options
            {
                ComputeReferences = true,
                View = view
            };

            GeometryElement geo = el.get_Geometry(opt);

            double minDist = double.MaxValue;
            Reference closestRef = null;

            foreach (GeometryObject obj in geo)
            {
                if (obj is Solid solid)
                {
                    foreach (Edge edge in solid.Edges)
                    {
                        Curve c = edge.AsCurve();

                        var projResult = c.Project(pickPoint);
                        if (projResult == null) continue;

                        XYZ proj = projResult.XYZPoint;
                        double dist = proj.DistanceTo(pickPoint);

                        if (dist < minDist)
                        {
                            minDist = dist;
                            closestRef = edge.Reference;
                        }
                    }
                }
            }

            return closestRef;
        }

        private bool IsValidReferenceForDimension(Reference r)
        {
            if (r == null) return false;

            // Nếu đã là reference chuẩn (face, edge, curve)
            return r.ElementReferenceType == ElementReferenceType.REFERENCE_TYPE_SURFACE
                || r.ElementReferenceType == ElementReferenceType.REFERENCE_TYPE_LINEAR
                || r.ElementReferenceType == ElementReferenceType.REFERENCE_TYPE_NONE;
        }
    }
}