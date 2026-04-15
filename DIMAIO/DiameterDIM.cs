using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using Autodesk.Revit.UI.Selection;
using System;
using System.Linq;
using System.Text;
using View = Autodesk.Revit.DB.View;
using TaskDialog = Autodesk.Revit.UI.TaskDialog;

namespace DIMAIO
{
    [Autodesk.Revit.Attributes.Transaction(Autodesk.Revit.Attributes.TransactionMode.Manual)]
    public class DiameterDIMCommand : IExternalCommand
    {
        public Result Execute(ExternalCommandData commandData, ref string message, ElementSet elements)
        {
            UIDocument uiDoc = commandData.Application.ActiveUIDocument;
            Document doc = uiDoc.Document;
            var sb = new StringBuilder();

            try
            {
                Reference pickedRef = uiDoc.Selection.PickObject(ObjectType.Element, "Chọn cung tròn");
                Element el = doc.GetElement(pickedRef);

                sb.AppendLine($"Element: {el.GetType().Name}, Id: {el.Id}");

                Arc arcCurve = null;

                // DetailArc
                if (el is CurveElement ce)
                {
                    arcCurve = ce.GeometryCurve as Arc;
                }
                // Wall
                else if (el is Wall wall)
                {
                    LocationCurve lc = wall.Location as LocationCurve;
                    arcCurve = lc?.Curve as Arc;
                }

                if (arcCurve == null)
                {
                    TaskDialog.Show("Debug", "Khong co Arc.");
                    message = "Khong co Arc geometry.";
                    return Result.Failed;
                }

                XYZ center = arcCurve.Center;
                XYZ startPt = arcCurve.GetEndPoint(0);
                XYZ dir = (startPt - center).Normalize();
                XYZ dimEnd = center + dir * (arcCurve.Radius * 1.5);
                Line dimLine = Line.CreateBound(center, dimEnd);

                Reference arcRef = null;

                // DetailArc -> dung Reference(el)
                if (el is CurveElement)
                {
                    arcRef = new Reference(el);
                }
                // Wall -> thu cach khac nhau
                else if (el is Wall wall)
                {
                    sb.AppendLine("Wall: thu cac cach lay reference...");

                    // Cach 1: Edge reference (nhu Radial)
                    arcRef = FindArcEdgeRef(wall, arcCurve, doc.ActiveView, sb);

                    // Cach 2: HostObjectUtils face reference
                    if (arcRef == null)
                    {
                        sb.AppendLine("  Cach 2: GetSideFaces...");
                        var faces = HostObjectUtils.GetSideFaces(wall, ShellLayerType.Exterior);
                        if (faces != null && faces.Count > 0)
                        {
                            sb.AppendLine($"  GetSideFaces: {faces.Count} faces");
                            arcRef = faces[0];
                            sb.AppendLine($"  -> Using face ref: {arcRef.ElementId}");
                        }
                    }
                }

                if (arcRef == null)
                {
                    TaskDialog.Show("Debug", sb.ToString());
                    message = "Khong lay duoc reference.";
                    return Result.Failed;
                }

                sb.AppendLine($"arcRef: {arcRef.ElementId}");

                using (Transaction tx = new Transaction(doc, "Diameter DIM"))
                {
                    tx.Start();

                    ReferenceArray refArray = new ReferenceArray();
                    refArray.Append(arcRef);
                    sb.AppendLine($"refArray.Size: {refArray.Size}");

                    try
                    {
                        Dimension dim = doc.Create.NewDimension(doc.ActiveView, dimLine, refArray);
                        if (dim != null)
                        {
                            DimensionType dt = new FilteredElementCollector(doc)
                                .OfClass(typeof(DimensionType))
                                .Cast<DimensionType>()
                                .FirstOrDefault(x => x.StyleType == DimensionStyleType.Diameter);
                            if (dt != null) dim.DimensionType = dt;

                            tx.Commit();
                            TaskDialog.Show("OK", sb.ToString());
                            return Result.Succeeded;
                        }
                    }
                    catch (Exception ex)
                    {
                        sb.AppendLine($"FAILED: {ex.Message}");
                        TaskDialog.Show("Debug", sb.ToString());
                    }

                    tx.RollBack();
                    message = "Tao Diameter that bai.";
                    return Result.Failed;
                }
            }
            catch (Autodesk.Revit.Exceptions.OperationCanceledException)
            {
                return Result.Cancelled;
            }
            catch (Exception ex)
            {
                sb.AppendLine($"EXCEPTION: {ex.Message}");
                TaskDialog.Show("Debug", sb.ToString());
                message = ex.Message;
                return Result.Failed;
            }
        }

        private Reference FindArcEdgeRef(Element wall, Arc arcCurve, View view, StringBuilder sb)
        {
            double targetRadius = arcCurve.Radius;
            Reference bestRef = null;
            double bestDiff = double.MaxValue;

            Options opt = new Options { ComputeReferences = true, View = view, IncludeNonVisibleObjects = true };
            GeometryElement geo = wall.get_Geometry(opt);
            if (geo == null) return null;

            foreach (GeometryObject obj in geo)
            {
                bestDiff = FindArcs(obj, targetRadius, ref bestRef, bestDiff);
            }
            sb.AppendLine($"  EdgeRef: {(bestRef != null ? $"{bestRef.ElementId}" : "null")}, diff={bestDiff:F4}");
            return bestRef;
        }

        private double FindArcs(GeometryObject obj, double targetRadius, ref Reference bestRef, double bestDiff)
        {
            if (obj is Solid solid)
            {
                foreach (Edge edge in solid.Edges)
                {
                    if (edge.AsCurve() is Arc edgeArc)
                    {
                        double rDiff = Math.Abs(edgeArc.Radius - targetRadius);
                        if (rDiff < bestDiff)
                        {
                            bestDiff = rDiff;
                            bestRef = edge.Reference;
                        }
                    }
                }
            }
            else if (obj is GeometryInstance inst)
            {
                foreach (GeometryObject iobj in inst.GetInstanceGeometry())
                    bestDiff = FindArcs(iobj, targetRadius, ref bestRef, bestDiff);
            }
            return bestDiff;
        }
    }
}
