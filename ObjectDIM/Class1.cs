using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using Autodesk.Revit.UI.Selection;
using System;
using System.Collections.Generic;
using System.Linq;

namespace ObjectDIM
{
    [Autodesk.Revit.Attributes.Transaction(Autodesk.Revit.Attributes.TransactionMode.Manual)]
    public class Command : IExternalCommand
    {
        public Result Execute(ExternalCommandData commandData, ref string message, ElementSet elements)
        {
            UIDocument uidoc = commandData.Application.ActiveUIDocument;
            Document doc = uidoc.Document;
            View view = doc.ActiveView;

            TaskDialog.Show("DEBUG", "🚀 Start command");

            // Pick element
            Reference pick = uidoc.Selection.PickObject(ObjectType.Element, "Pick a family");
            FamilyInstance fi = doc.GetElement(pick) as FamilyInstance;

            if (fi == null)
            {
                TaskDialog.Show("ERROR", "❌ Not a FamilyInstance");
                return Result.Failed;
            }

            TaskDialog.Show("INFO", $"✅ Picked: {fi.Name}");

            // Get edges (lọc đúng mặt phẳng view)
            List<EdgeData> edges = GetAllEdges(fi, view);

            TaskDialog.Show("INFO", $"📊 Edges valid: {edges.Count}");

            if (edges.Count < 2)
            {
                TaskDialog.Show("ERROR", "❌ Not enough valid edges");
                return Result.Failed;
            }

            // 👉 hướng ngang của view
            XYZ rightDir = view.RightDirection;

            // tìm 2 cạnh ngoài cùng
            EdgeData left = edges.OrderBy(e => e.MidPoint.DotProduct(rightDir)).First();
            EdgeData right = edges.OrderByDescending(e => e.MidPoint.DotProduct(rightDir)).First();

            // tạo line DIM
            Line dimLine = Line.CreateBound(left.MidPoint, right.MidPoint);

            // reference
            ReferenceArray refArr = new ReferenceArray();
            refArr.Append(left.Ref);
            refArr.Append(right.Ref);

            using (Transaction t = new Transaction(doc, "Outer DIM"))
            {
                try
                {
                    t.Start();

                    Dimension dim = doc.Create.NewDimension(view, dimLine, refArr);

                    if (dim == null)
                    {
                        TaskDialog.Show("ERROR", "❌ Create DIM failed");
                        t.RollBack();
                        return Result.Failed;
                    }

                    t.Commit();
                }
                catch (Exception ex)
                {
                    TaskDialog.Show("EXCEPTION", $"💥 {ex.Message}");
                    return Result.Failed;
                }
            }

            TaskDialog.Show("SUCCESS", "🎉 DIM created!");

            return Result.Succeeded;
        }

        // ====================================
        // 🔥 GET ALL EDGES
        // ====================================
        private List<EdgeData> GetAllEdges(FamilyInstance fi, View view)
        {
            List<EdgeData> result = new List<EdgeData>();

            Options opt = new Options();
            opt.ComputeReferences = true;
            opt.IncludeNonVisibleObjects = true;

            GeometryElement geo = fi.get_Geometry(opt);

            foreach (GeometryObject obj in geo)
            {
                ExtractEdges(obj, result, view);
            }

            return result;
        }

        private void ExtractEdges(GeometryObject obj, List<EdgeData> result, View view)
        {
            // Solid
            if (obj is Solid solid && solid.Volume > 0)
            {
                foreach (Edge edge in solid.Edges)
                {
                    if (!IsEdgeInViewPlane(view, edge))
                        continue;

                    IList<XYZ> pts = edge.Tessellate();
                    if (pts.Count == 0) continue;

                    XYZ mid = pts[pts.Count / 2];

                    result.Add(new EdgeData
                    {
                        Ref = edge.Reference,
                        MidPoint = mid
                    });
                }
            }
            // Nested
            else if (obj is GeometryInstance gi)
            {
                GeometryElement instGeo = gi.GetInstanceGeometry();
                foreach (GeometryObject instObj in instGeo)
                {
                    ExtractEdges(instObj, result, view);
                }
            }
        }

        // ====================================
        // 🎯 CHECK EDGE NẰM TRONG VIEW PLANE
        // ====================================
        private bool IsEdgeInViewPlane(View view, Edge edge)
        {
            Curve c = edge.AsCurve();

            if (!(c is Line line)) return false;

            XYZ dir = line.Direction.Normalize();
            XYZ viewDir = view.ViewDirection.Normalize();

            double dot = Math.Abs(dir.DotProduct(viewDir));

            double tol = 1e-4;

            // dot ≈ 0 => edge nằm trong plane
            return dot < tol;
        }

        // ====================================
        // 📦 DATA
        // ====================================
        private class EdgeData
        {
            public Reference Ref;
            public XYZ MidPoint;
        }
    }
}