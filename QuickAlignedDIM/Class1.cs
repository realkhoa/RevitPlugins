using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using Autodesk.Revit.UI.Selection;
using System;

namespace QuickAlignedDIM
{
    [Autodesk.Revit.Attributes.Transaction(Autodesk.Revit.Attributes.TransactionMode.Manual)]
    [Autodesk.Revit.Attributes.Regeneration(Autodesk.Revit.Attributes.RegenerationOption.Manual)]
    public class Command : IExternalCommand
    {
        private const double TOLERANCE = 1e-6;

        public Result Execute(ExternalCommandData commandData, ref string message, ElementSet elements)
        {
            UIApplication uiApp = commandData.Application;
            UIDocument uidoc = uiApp.ActiveUIDocument;
            Document doc = uidoc.Document;
            View view = doc.ActiveView;

            if (view is View3D || view is ViewSheet)
            {
                message = "Chỉ dùng trong view 2D!";
                return Result.Failed;
            }

            int count = 0;

            while (true)
            {
                try
                {
                    Reference r1 = uidoc.Selection.PickObject(ObjectType.PointOnElement, "Pick ref 1 (ESC để thoát)");
                    Reference r2 = uidoc.Selection.PickObject(ObjectType.PointOnElement, "Pick ref 2");

                    DimRefData d1 = GetData(doc, view, r1);
                    DimRefData d2 = GetData(doc, view, r2);

                    if (d1 == null || d2 == null)
                    {
                        TaskDialog.Show("Lỗi", "Pick không hợp lệ, thử lại!");
                        continue;
                    }

                    if (!IsParallel(d1.Normal, d2.Normal))
                    {
                        TaskDialog.Show("Lỗi", "2 đối tượng không song song!");
                        continue;
                    }

                    using (Transaction tx = new Transaction(doc, "Quick DIM"))
                    {
                        tx.Start();

                        Dimension dim = CreateDim(doc, view, d1, d2);

                        if (dim != null)
                        {
                            tx.Commit();
                            count++;
                        }
                        else
                        {
                            tx.RollBack();
                        }
                    }
                }
                catch (Autodesk.Revit.Exceptions.OperationCanceledException)
                {
                    break;
                }
                catch (Exception ex)
                {
                    TaskDialog.Show("ERROR", ex.Message);
                    break;
                }
            }

            TaskDialog.Show("DONE 😎", $"Tạo được {count} DIM");

            return Result.Succeeded;
        }

        private XYZ GetOffsetMidPoint(View view, DimRefData d1, DimRefData d2, double offset)
        {
            XYZ p1 = d1.GlobalPoint;
            XYZ p2 = d2.GlobalPoint;

            XYZ dimVec = (p2 - p1).Normalize();
            XYZ viewNormal = view.ViewDirection;

            // hướng vuông góc
            XYZ dimDir = viewNormal.CrossProduct(dimVec).Normalize();

            // midpoint
            XYZ mid = (p1 + p2) / 2.0;

            // 👉 vector từ center tới điểm user click (đại diện hướng "bên ngoài")
            XYZ userDir = (d1.GlobalPoint - mid);

            // 👉 nếu đang ngược hướng → đảo lại
            if (userDir.DotProduct(dimDir) < 0)
            {
                dimDir = dimDir.Negate();
            }

            // 👉 offset ra ngoài
            mid += dimDir * offset;

            // fix Z cho plan
            if (view is ViewPlan)
            {
                double z = view.GenLevel?.Elevation ?? 0;
                mid = new XYZ(mid.X, mid.Y, z);
            }

            return mid;
        }

        // ================= CREATE DIM =================
        private Dimension CreateDim(Document doc, View view, DimRefData d1, DimRefData d2)
        {
            ReferenceArray arr = new ReferenceArray();
            arr.Append(d1.Ref);
            arr.Append(d2.Ref);

            XYZ p1 = d1.GlobalPoint;
            XYZ p2 = d2.GlobalPoint;

            if (p1 == null || p2 == null) return null;

            // vector nối 2 điểm
            XYZ dimVec = (p2 - p1).Normalize();
            XYZ viewNormal = view.ViewDirection;

            // hướng DIM (vuông góc)
            XYZ dimDir = viewNormal.CrossProduct(dimVec).Normalize();

            // midpoint
            XYZ mid = (p1 + p2) / 2.0;

            // =========================
            // 🔥 FIX 1: AUTO HƯỚNG RA NGOÀI
            // =========================
            XYZ userDir = (d1.GlobalPoint - mid);

            if (userDir.DotProduct(dimDir) < 0)
            {
                dimDir = dimDir.Negate();
            }

            // =========================
            // 🔥 FIX 2: OFFSET RA NGOÀI
            // =========================
            double offset = UnitUtils.ConvertToInternalUnits(100, UnitTypeId.Millimeters);
            mid += dimDir * offset;

            // =========================
            // FIX Z (PLAN)
            // =========================
            if (view is ViewPlan)
            {
                double z = view.GenLevel?.Elevation ?? 0;
                mid = new XYZ(mid.X, mid.Y, z);
            }

            // tạo line
            Line line = Line.CreateBound(
                mid - dimDir * 10,
                mid + dimDir * 10
            );

            return doc.Create.NewDimension(view, line, arr);
        }

        // ================= GET DATA =================
        private DimRefData GetData(Document doc, View view, Reference reference)
        {
            Element elem = doc.GetElement(reference);
            GeometryObject geo = elem.GetGeometryObjectFromReference(reference);

            XYZ normal = null;
            XYZ pt = reference.GlobalPoint;

            // ---- FACE ----
            if (geo is PlanarFace pf)
            {
                normal = pf.FaceNormal;
            }

            // ---- EDGE ----
            else if (geo is Edge edge)
            {
                Curve c = edge.AsCurve();

                // 🔥 Lấy midpoint thật của edge (không dùng GlobalPoint nữa)
                pt = c.Evaluate(0.5, true);

                if (c is Line l)
                {
                    XYZ dir = (l.GetEndPoint(1) - l.GetEndPoint(0)).Normalize();
                    normal = Perp2D(dir);
                }
                else if (c is Arc a)
                {
                    XYZ dir = (a.GetEndPoint(1) - a.GetEndPoint(0)).Normalize();
                    normal = Perp2D(dir);
                }
            }

            // ---- LINE ----
            else if (geo is Line line)
            {
                XYZ dir = (line.GetEndPoint(1) - line.GetEndPoint(0)).Normalize();
                normal = Perp2D(dir);
            }

            // ---- LOCATION CURVE ----
            if (normal == null && elem.Location is LocationCurve lc)
            {
                if (lc.Curve is Line l)
                {
                    XYZ dir = (l.GetEndPoint(1) - l.GetEndPoint(0)).Normalize();
                    normal = Perp2D(dir);

                    if (pt == null)
                        pt = l.Evaluate(0.5, true);
                }
            }

            // ---- GRID ----
            if (normal == null && elem is Grid grid)
            {
                if (grid.Curve is Line l)
                {
                    XYZ dir = (l.GetEndPoint(1) - l.GetEndPoint(0)).Normalize();
                    normal = Perp2D(dir);

                    if (pt == null)
                        pt = l.Evaluate(0.5, true);
                }
            }

            // ---- REF PLANE ----
            if (normal == null && elem is ReferencePlane rp)
            {
                normal = rp.Normal;

                if (pt == null)
                    pt = rp.BubbleEnd;
            }

            // ---- FAMILY ----
            if (normal == null && elem is FamilyInstance fi)
            {
                if (fi.Location is LocationPoint lp)
                {
                    normal = XYZ.BasisX;
                    pt = lp.Point;
                }
            }

            if (normal == null) return null;

            if (view is ViewPlan)
            {
                normal = new XYZ(normal.X, normal.Y, 0).Normalize();
            }

            if (pt == null)
                pt = XYZ.Zero;

            return new DimRefData
            {
                Ref = reference,
                Normal = normal,
                GlobalPoint = pt,
                SourceElement = elem
            };
        }

        // ================= HELPER =================
        private XYZ Perp2D(XYZ v)
        {
            return new XYZ(-v.Y, v.X, 0).Normalize();
        }

        private bool IsParallel(XYZ a, XYZ b)
        {
            return a.IsAlmostEqualTo(b, TOLERANCE) ||
                   a.IsAlmostEqualTo(b.Negate(), TOLERANCE);
        }

        // ================= DATA CLASS =================
        private class DimRefData
        {
            public Reference Ref;
            public XYZ Normal;
            public XYZ GlobalPoint;
            public Element SourceElement;
        }
    }
}