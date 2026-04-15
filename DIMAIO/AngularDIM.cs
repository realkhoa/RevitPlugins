using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using Autodesk.Revit.UI.Selection;
using System;
using System.Collections.Generic;
using System.Linq;
using View = Autodesk.Revit.DB.View;

namespace DIMAIO
{
    [Autodesk.Revit.Attributes.Transaction(Autodesk.Revit.Attributes.TransactionMode.Manual)]
    public class AngularDIMCommand : IExternalCommand
    {
        public Result Execute(ExternalCommandData commandData, ref string message, ElementSet elements)
        {
            UIDocument uiDoc = commandData.Application.ActiveUIDocument;
            Document doc = uiDoc.Document;

            try
            {
                Reference r1 = uiDoc.Selection.PickObject(ObjectType.Edge, "Chọn cạnh thứ 1");
                Reference r2 = uiDoc.Selection.PickObject(ObjectType.Edge, "Chọn cạnh thứ 2");

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

                    XYZ mid1 = (line1.GetEndPoint(0) + line1.GetEndPoint(1)) * 0.5;
                    XYZ mid2 = (line2.GetEndPoint(0) + line2.GetEndPoint(1)) * 0.5;
                    XYZ dir1 = line1.Direction.Normalize();
                    XYZ dir2 = line2.Direction.Normalize();

                    XYZ normal = dir1.CrossProduct(dir2);
                    if (normal.GetLength() < 1e-6)
                    {
                        XYZ connect = (mid2 - mid1).Normalize();
                        normal = dir1.CrossProduct(connect).Normalize();
                    }
                    else
                    {
                        normal = normal.Normalize();
                    }

                    Plane plane = Plane.CreateByNormalAndOrigin(normal, mid1);
                    Line l1 = ProjectLineToPlane(line1, plane);
                    Line l2 = ProjectLineToPlane(line2, plane);

                    l1 = Line.CreateUnbound(mid1, l1.Direction);
                    l2 = Line.CreateUnbound(mid2, l2.Direction);

                    XYZ v1 = l1.Direction.Normalize();
                    XYZ v2 = l2.Direction.Normalize();
                    double angle = v1.AngleTo(v2);
                    if (angle < 1e-6)
                    {
                        message = "2 cạnh gần như song song.";
                        tx.RollBack();
                        return Result.Failed;
                    }

                    XYZ xAxis = v1;
                    XYZ yAxis = normal.CrossProduct(xAxis).Normalize();
                    double radius = mid1.DistanceTo(mid2) * 0.5;
                    if (radius < 0.1) radius = 1.0;

                    Arc arc = Arc.Create(mid1, radius, 0, angle, xAxis, yAxis);

                    if (doc.IsFamilyDocument)
                    {
                        doc.FamilyCreate.NewAngularDimension(view, arc, r1, r2);
                    }
                    else
                    {
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

                        AngularDimension.Create(doc, view, arc, new List<Reference> { r1, r2 }, angDimType);
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

        private Line GetLineFromReference(Document doc, Reference r)
        {
            Element el = doc.GetElement(r);

            // Thử trực tiếp (hoạt động tốt trong Family Editor)
            try
            {
                GeometryObject geoObj = el.GetGeometryObjectFromReference(r);
                if (geoObj is Edge edge)
                {
                    Curve curve = edge.AsCurve();
                    if (curve is Line l) return l;
                    return Line.CreateBound(curve.GetEndPoint(0), curve.GetEndPoint(1));
                }
            }
            catch { }

            // Fallback: duyệt qua geometry (cần cho Project document với family instance)
            Options opt = new Options { ComputeReferences = true, IncludeNonVisibleObjects = true };
            GeometryElement geo = el.get_Geometry(opt);
            if (geo == null) return null;

            foreach (GeometryObject obj in geo)
            {
                if (obj is Solid solid)
                {
                    Line line = FindLineInSolid(solid, r);
                    if (line != null) return line;
                }
                if (obj is GeometryInstance inst)
                {
                    foreach (GeometryObject iobj in inst.GetInstanceGeometry())
                    {
                        if (iobj is Solid solid2)
                        {
                            Line line = FindLineInSolid(solid2, r);
                            if (line != null) return line;
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
                    if (c is Line line) return line;
                }
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
