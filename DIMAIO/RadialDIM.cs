using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using Autodesk.Revit.UI.Selection;
using System;
using System.Linq;
using View = Autodesk.Revit.DB.View;

namespace DIMAIO
{
    [Autodesk.Revit.Attributes.Transaction(Autodesk.Revit.Attributes.TransactionMode.Manual)]
    public class RadialDIMCommand : IExternalCommand
    {
        public Result Execute(ExternalCommandData commandData, ref string message, ElementSet elements)
        {
            UIDocument uiDoc = commandData.Application.ActiveUIDocument;
            Document doc = uiDoc.Document;

            try
            {
                // Pick Wall element
                Reference wallRef = uiDoc.Selection.PickObject(ObjectType.Element, new WallSelectionFilter(), "Chọn tường cong (arc/circle)");
                Element wallEl = doc.GetElement(wallRef);

                LocationCurve lc = (wallEl as Wall)?.Location as LocationCurve;
                if (!(lc?.Curve is Arc wallArc))
                {
                    message = "Tường được chọn không có curve dạng Arc.";
                    return Result.Failed;
                }

                // Lay arc reference tu wall geometry
                Reference arcEdgeRef = FindArcEdgeReferenceOnWall(wallEl, wallArc, doc.ActiveView);
                if (arcEdgeRef == null)
                {
                    message = "Không tìm được arc reference trên tường.";
                    return Result.Failed;
                }

                using (Transaction tx = new Transaction(doc, "Radial DIM"))
                {
                    tx.Start();
                    View view = doc.ActiveView;

                    // Thu 1: RadialDimension.Create(doc, view, ref, bool) - default placement
                    try
                    {
                        Dimension radDim = RadialDimension.Create(doc, view, arcEdgeRef, false);
                        if (radDim != null)
                        {
                            tx.Commit();
                            return Result.Succeeded;
                        }
                    }
                    catch { }

                    // Thu 2: doc.FamilyCreate.NewRadialDimension(view, ref, placementPoint)
                    try
                    {
                        XYZ placementPoint = uiDoc.Selection.PickPoint("Chọn vị trí đặt DIM");
                        doc.FamilyCreate.NewRadialDimension(view, arcEdgeRef, placementPoint);
                        tx.Commit();
                        return Result.Succeeded;
                    }
                    catch { }

                    tx.RollBack();
                    message = "Không tạo được Radial Dimension.";
                    return Result.Failed;
                }
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

        private Reference FindArcEdgeReferenceOnWall(Element wallEl, Arc wallArc, View view)
        {
            XYZ arcCenter = wallArc.Center;
            double arcRadius = wallArc.Radius;
            Reference bestRef = null;
            double bestDiff = double.MaxValue;

            Options opt = new Options { ComputeReferences = true, View = view, IncludeNonVisibleObjects = true };
            GeometryElement geo = wallEl.get_Geometry(opt);
            if (geo == null) return null;

            // Duyet geometry tra ve (la GeometryObject, khong phai Solid)
            foreach (GeometryObject obj in geo)
            {
                FindBestArcEdgeRecursive(obj, arcCenter, arcRadius, ref bestRef, ref bestDiff);
            }
            return bestRef;
        }

        private void FindBestArcEdgeRecursive(GeometryObject obj, XYZ arcCenter, double arcRadius,
            ref Reference bestRef, ref double bestDiff)
        {
            if (obj is Solid)
            {
                Solid solid = obj as Solid;
                foreach (Edge edge in solid.Edges)
                {
                    if (edge.AsCurve() is Arc edgeArc)
                    {
                        double cDist = edgeArc.Center.DistanceTo(arcCenter);
                        double diff = Math.Abs(edgeArc.Radius - arcRadius);
                        // Mo rong tolerance: vi du arc radius = 12.5, edge radii = 12.1667 hoac 12.8333 (chenh 0.33)
                        if (cDist < 0.1 && diff < bestDiff)
                        {
                            bestRef = edge.Reference;
                            bestDiff = diff;
                        }
                    }
                }
            }
            else if (obj is GeometryInstance inst)
            {
                foreach (GeometryObject iobj in inst.GetInstanceGeometry())
                    FindBestArcEdgeRecursive(iobj, arcCenter, arcRadius, ref bestRef, ref bestDiff);
            }
        }

        private class WallSelectionFilter : ISelectionFilter
        {
            public bool AllowElement(Element elem) => elem is Wall;
            public bool AllowReference(Reference reference, XYZ position) => false;
        }
    }
}
