using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using Autodesk.Revit.UI.Selection;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Reflection;

namespace AlignedDIM
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
                "AlignedDIM",
                "Aligned\nDIM",
                assemblyPath,
                "AlignedDIM.Command"
            );

            if (!panel.GetItems().OfType<PushButton>().Any(b => b.Name == "AlignedDIM"))
            {
                PushButton button = panel.AddItem(btnData) as PushButton;
                button.ToolTip = "DIM thông minh: Chọn N đối tượng (bao gồm Ref Plane, Cạnh Solid) → Live render DIM liên tục.";
            }

            return Result.Succeeded;
        }

        public Result OnShutdown(UIControlledApplication app) => Result.Succeeded;
    }

    public class DimSelectionFilter : ISelectionFilter
    {
        private Document _doc;
        public DimSelectionFilter(Document doc) { _doc = doc; }
        public bool AllowElement(Element elem) => true;
        public bool AllowReference(Reference reference, XYZ position) => true;
    }

    [Autodesk.Revit.Attributes.Transaction(Autodesk.Revit.Attributes.TransactionMode.Manual)]
    public class Command : IExternalCommand
    {
        public Result Execute(ExternalCommandData commandData, ref string message, ElementSet elements)
        {
            UIDocument uiDoc = commandData.Application.ActiveUIDocument;
            Document doc = uiDoc.Document;
            bool isFamilyDoc = doc.IsFamilyDocument;

            if (doc.ActiveView.ViewType == ViewType.ThreeD)
            {
                TaskDialog.Show("Cảnh báo", "Lệnh DIM chỉ hoạt động trên mặt bằng, mặt cắt hoặc mặt đứng.");
                return Result.Failed;
            }

            try
            {
                while (true)
                {
                    List<Reference> pickedRefs = new List<Reference>();
                    Line dimLine = null;
                    ElementId currentDimId = ElementId.InvalidElementId;
                    View view = doc.ActiveView;

                    while (true)
                    {
                        string prompt = pickedRefs.Count < 2
                            ? $"Pick tham chiếu thứ {pickedRefs.Count + 1} (Click gần cạnh cần DIM - ESC để hoàn tất)"
                            : $"Pick thêm tham chiếu thứ {pickedRefs.Count + 1} (ESC để hoàn tất)";

                        Reference r = TryPickReference(uiDoc, prompt);
                        if (r == null) break;

                        pickedRefs.Add(r);

                        if (pickedRefs.Count >= 2)
                        {
                            try
                            {
                                if (dimLine == null) dimLine = CalculateDimLine(doc, view, pickedRefs, isFamilyDoc);

                                if (dimLine == null)
                                {
                                    TaskDialog.Show("Lỗi Toán Học", "Không thể nội suy được trục của DIM từ 2 điểm chọn. Vui lòng chọn lại.");
                                    pickedRefs.RemoveAt(pickedRefs.Count - 1);
                                    continue;
                                }
                            }
                            catch (Exception calcEx)
                            {
                                TaskDialog.Show("Lỗi Tính Toán Trục DIM", calcEx.Message);
                                pickedRefs.RemoveAt(pickedRefs.Count - 1);
                                continue;
                            }

                            using (Transaction tx = new Transaction(doc, "Live Render DIM"))
                            {
                                tx.Start();
                                try
                                {
                                    if (currentDimId != ElementId.InvalidElementId) doc.Delete(currentDimId);

                                    Dimension newDim = CreateMultiDim(doc, view, pickedRefs, dimLine, isFamilyDoc, out string err);

                                    if (newDim != null)
                                    {
                                        currentDimId = newDim.Id;
                                        tx.Commit();
                                    }
                                    else
                                    {
                                        tx.RollBack();
                                        pickedRefs.RemoveAt(pickedRefs.Count - 1);
                                        TaskDialog.Show("Từ chối tạo DIM", $"Revit API không chấp nhận mảng tham chiếu này.\n\nChi tiết lỗi: {err}");
                                    }
                                }
                                catch (Exception txEx)
                                {
                                    if (tx.HasStarted()) tx.RollBack();
                                    pickedRefs.RemoveAt(pickedRefs.Count - 1);
                                    TaskDialog.Show("Lỗi Kịch Bản Transaction", txEx.Message);
                                }
                            }
                        }
                    }

                    // Xử lý 1 đối tượng
                    if (pickedRefs.Count == 1 && currentDimId == ElementId.InvalidElementId)
                    {
                        using (Transaction tx = new Transaction(doc, "Smart DIM 1 Object"))
                        {
                            tx.Start();
                            try
                            {
                                bool success = isFamilyDoc
                                    ? Case1_Family(doc, view, pickedRefs[0], out string err)
                                    : Case1_Document(doc, view, pickedRefs[0], out err);

                                if (success) tx.Commit();
                                else
                                {
                                    tx.RollBack();
                                    TaskDialog.Show("Lỗi DIM 1 Đối tượng", err ?? "Không có lỗi chi tiết được trả về.");
                                }
                            }
                            catch (Exception case1Ex)
                            {
                                tx.RollBack();
                                TaskDialog.Show("Lỗi Hệ Thống DIM 1 Đối Tượng", case1Ex.Message);
                            }
                        }
                    }

                    if (pickedRefs.Count == 0) break;
                }

                return Result.Succeeded;
            }
            catch (Exception mainEx)
            {
                TaskDialog.Show("Lỗi Vòng Lặp Tổng", $"Có lỗi nghiêm trọng gây crash lệnh:\n\n{mainEx.Message}\n\n{mainEx.StackTrace}");
                message = mainEx.Message;
                return Result.Failed;
            }
        }

        private Reference TryPickReference(UIDocument uiDoc, string prompt)
        {
            try
            {
                DimSelectionFilter filter = new DimSelectionFilter(uiDoc.Document);
                return uiDoc.Selection.PickObject(ObjectType.PointOnElement, filter, prompt);
            }
            catch (Autodesk.Revit.Exceptions.OperationCanceledException) { return null; }
            catch (Exception ex)
            {
                TaskDialog.Show("Lỗi Pick Object", ex.Message);
                return null;
            }
        }

        // ═══════════════════════════════════════════════════════
        //  CORE LOGIC CHO MULTI-DIM LIVE RENDER
        // ═══════════════════════════════════════════════════════
        private Line CalculateDimLine(Document doc, View view, List<Reference> picks, bool isFamilyDoc)
        {
            List<XYZ> pts = new List<XYZ>();
            XYZ refDir = XYZ.Zero;

            for (int i = 0; i < 2; i++)
            {
                Reference r = picks[i];
                Element el = doc.GetElement(r);
                Reference resolvedRef = ResolveReference(el, r, view);

                if (refDir.IsZeroLength())
                {
                    if (el is ReferencePlane rp) refDir = rp.Direction;
                    else if (el is Grid grid && grid.Curve is Line gLine) refDir = gLine.Direction;
                    else
                    {
                        try
                        {
                            GeometryObject geoObj = el.GetGeometryObjectFromReference(resolvedRef);
                            if (geoObj is Edge edge && edge.AsCurve() is Line line) refDir = line.Direction;
                            else if (geoObj is PlanarFace pf) refDir = view.ViewDirection.CrossProduct(pf.FaceNormal);
                        }
                        catch { /* Bỏ qua lỗi trích xuất hướng, giữ nguyên fallback XYZ.Zero */ }
                    }
                }

                if (r.GlobalPoint != null) pts.Add(r.GlobalPoint);
                else
                {
                    if (el is ReferencePlane rp) pts.Add(rp.FreeEnd);
                    else if (el is Grid grid && grid.Curve is Line gl) pts.Add(gl.Origin);
                }
            }

            if (pts.Count < 2) return null;

            XYZ refDirOnPlane = refDir - refDir.DotProduct(view.ViewDirection) * view.ViewDirection;
            refDirOnPlane = refDirOnPlane.IsZeroLength() ? view.RightDirection : refDirOnPlane.Normalize();

            XYZ measureDir = view.ViewDirection.CrossProduct(refDirOnPlane).Normalize();
            if (measureDir.IsZeroLength()) measureDir = view.UpDirection;

            XYZ midPoint = (pts[0] + pts[1]) / 2.0;
            XYZ offsetDir = view.ViewDirection.CrossProduct(measureDir).Normalize();
            if (offsetDir.IsZeroLength()) offsetDir = view.UpDirection;
            if (offsetDir.DotProduct(view.UpDirection) < 0) offsetDir = -offsetDir;

            double offset = UnitUtils.ConvertToInternalUnits(300, UnitTypeId.Millimeters);
            XYZ placementPt = midPoint + offsetDir * offset;

            double distToPlane = (placementPt - view.Origin).DotProduct(view.ViewDirection);
            placementPt = placementPt - distToPlane * view.ViewDirection;

            return Line.CreateBound(placementPt, placementPt + measureDir * 10);
        }

        private Dimension CreateMultiDim(Document doc, View view, List<Reference> picks, Line dimLine, bool isFamilyDoc, out string error)
        {
            error = null;
            ReferenceArray refArray = new ReferenceArray();

            foreach (Reference r in picks)
            {
                Element el = doc.GetElement(r);
                refArray.Append(ResolveReference(el, r, view));
            }

            try
            {
                return isFamilyDoc
                    ? doc.FamilyCreate.NewDimension(view, dimLine, refArray)
                    : doc.Create.NewDimension(view, dimLine, refArray);
            }
            catch (Exception ex)
            {
                error = $"Lỗi từ Revit API: {ex.Message}";
                return null;
            }
        }

        // ═══════════════════════════════════════════════════════
        //  CASE 1: SỬA LỖI ĐIỂM MÚT TRONG FAMILY EDITOR
        // ═══════════════════════════════════════════════════════
        private bool Case1_Family(Document doc, View view, Reference pick, out string error)
        {
            error = null;
            try
            {
                Element el = doc.GetElement(pick);
                if (el is ReferencePlane) { error = "Không thể auto DIM chiều dài cho 1 Reference Plane (vì nó vô hạn)."; return false; }

                Reference resolvedRef = ResolveReference(el, pick, view);
                GeometryObject geoObj = el.GetGeometryObjectFromReference(resolvedRef);

                if (!(geoObj is Edge edge) || !(edge.AsCurve() is Line line))
                {
                    error = "Hãy đưa chuột trúng vào một cạnh thẳng để DIM.";
                    return false;
                }

                Reference ref1 = edge.GetEndPointReference(0);
                Reference ref2 = edge.GetEndPointReference(1);

                if (ref1 == null || ref2 == null)
                {
                    GeometryElement geoElem = el.get_Geometry(GetGeometryOptions(view));
                    if (TryFindFacesByNormal(geoElem, line.Direction, out PlanarFace f1, out PlanarFace f2))
                    {
                        ref1 = f1.Reference;
                        ref2 = f2.Reference;
                    }
                }

                if (ref1 == null || ref2 == null)
                {
                    error = "Hình học Solid này không hỗ trợ tham chiếu ở 2 mút. Vui lòng DIM từng cạnh thủ công bằng cách pick 2 lần.";
                    return false;
                }

                XYZ p1 = line.GetEndPoint(0);
                XYZ p2 = line.GetEndPoint(1);
                XYZ midPoint = (p1 + p2) / 2.0;

                XYZ offsetDir = view.ViewDirection.CrossProduct(line.Direction).Normalize();
                if (offsetDir.IsZeroLength()) offsetDir = view.UpDirection;
                if (offsetDir.DotProduct(view.UpDirection) < 0) offsetDir = -offsetDir;

                XYZ placementPt = midPoint + offsetDir * UnitUtils.ConvertToInternalUnits(200, UnitTypeId.Millimeters);

                double dist = (placementPt - view.Origin).DotProduct(view.ViewDirection);
                placementPt = placementPt - dist * view.ViewDirection;

                Line dimLine = Line.CreateBound(placementPt, placementPt + line.Direction * 10);

                doc.FamilyCreate.NewDimension(view, dimLine, BuildRefArray(ref1, ref2));
                return true;
            }
            catch (Exception ex)
            {
                error = $"Lỗi xử lý Case1_Family: {ex.Message}";
                return false;
            }
        }

        private bool Case1_Document(Document doc, View view, Reference pick, out string error)
        {
            error = null;
            try
            {
                Element el = doc.GetElement(pick);

                if (el is Wall wall) return DimWall(doc, view, wall, out error);
                if (el is ReferencePlane) { error = "Không thể auto DIM cho 1 Reference Plane."; return false; }
                if (pick.GlobalPoint != null) return DimNearestEdge(doc, view, el, pick.GlobalPoint, out error);

                return false;
            }
            catch (Exception ex)
            {
                error = $"Lỗi xử lý Case1_Document: {ex.Message}";
                return false;
            }
        }

        private bool DimWall(Document doc, View view, Wall wall, out string error)
        {
            error = null;
            try
            {
                LocationCurve lc = wall.Location as LocationCurve;
                if (!(lc?.Curve is Line wallLine)) return false;

                XYZ p1 = wallLine.GetEndPoint(0);
                XYZ p2 = wallLine.GetEndPoint(1);

                Reference refLen1 = GetClosestEdgeReference(wall, p1, view);
                Reference refLen2 = GetClosestEdgeReference(wall, p2, view);

                if (refLen1 != null && refLen2 != null)
                {
                    XYZ placementPt = (p1 + p2) / 2.0 + view.ViewDirection * 300;
                    double dist = (placementPt - view.Origin).DotProduct(view.ViewDirection);
                    placementPt = placementPt - dist * view.ViewDirection;

                    doc.Create.NewDimension(view, Line.CreateBound(placementPt, placementPt + wallLine.Direction * 10), BuildRefArray(refLen1, refLen2));
                }

                return true;
            }
            catch (Exception ex)
            {
                error = $"Lỗi khi DIM tường: {ex.Message}";
                return false;
            }
        }

        private bool DimNearestEdge(Document doc, View view, Element el, XYZ pickPoint, out string error)
        {
            error = null;
            try
            {
                Edge bestEdge = FindClosestEdge(el.get_Geometry(GetGeometryOptions(view)), pickPoint, lineOnly: true);
                if (bestEdge == null) { error = "Không tìm thấy cạnh hợp lệ tại vị trí click."; return false; }

                Curve curve = bestEdge.AsCurve();
                Reference ref1 = bestEdge.GetEndPointReference(0);
                Reference ref2 = bestEdge.GetEndPointReference(1);

                if (ref1 == null || ref2 == null) { error = "API không trả về tham chiếu điểm mút của cạnh này."; return false; }

                XYZ p1 = curve.GetEndPoint(0);
                XYZ p2 = curve.GetEndPoint(1);
                XYZ dir = (p2 - p1).Normalize();
                XYZ offsetDir = view.ViewDirection.CrossProduct(dir).Normalize();

                XYZ placementPt = (p1 + p2) / 2.0 + offsetDir * 300;
                double dist = (placementPt - view.Origin).DotProduct(view.ViewDirection);
                placementPt = placementPt - dist * view.ViewDirection;

                doc.Create.NewDimension(view, Line.CreateBound(placementPt, placementPt + dir * 10), BuildRefArray(ref1, ref2));
                return true;
            }
            catch (Exception ex)
            {
                error = $"Lỗi khi DIM cạnh đối tượng: {ex.Message}";
                return false;
            }
        }

        // ═══════════════════════════════════════════════════════
        //  HELPERS & GEOMETRY
        // ═══════════════════════════════════════════════════════
        private bool TryFindFacesByNormal(GeometryElement geoElem, XYZ normalDir, out PlanarFace face1, out PlanarFace face2)
        {
            face1 = face2 = null;
            var parallelFaces = new List<PlanarFace>();

            foreach (GeometryObject obj in geoElem)
            {
                if (obj is Solid solid && solid.Faces.Size > 0)
                {
                    foreach (Face face in solid.Faces)
                    {
                        if (face is PlanarFace pf && Math.Abs(pf.FaceNormal.DotProduct(normalDir)) > 0.99)
                        {
                            parallelFaces.Add(pf);
                        }
                    }
                }
            }

            if (parallelFaces.Count >= 2)
            {
                face1 = parallelFaces[0];
                face2 = parallelFaces[1];
                return true;
            }
            return false;
        }

        private Reference ResolveReference(Element el, Reference r, View view)
        {
            try
            {
                if (el is ReferencePlane rp) return rp.GetReference();
                if (el is Grid grid) return grid.Curve?.Reference ?? r;

                if (r.GlobalPoint != null)
                {
                    Reference closestEdge = GetClosestEdgeReference(el, r.GlobalPoint, view);
                    if (closestEdge != null) return closestEdge;
                }
            }
            catch { /* Lỗi thì trả về Reference gốc */ }
            return r;
        }

        private static ReferenceArray BuildRefArray(Reference a, Reference b)
        {
            var ra = new ReferenceArray();
            ra.Append(a); ra.Append(b);
            return ra;
        }

        private static Options GetGeometryOptions(View view, bool includeNonVisible = false) => new Options
        {
            ComputeReferences = true,
            View = view,
            IncludeNonVisibleObjects = includeNonVisible
        };

        private void TraverseEdges(GeometryObject obj, Action<Edge> visit)
        {
            if (obj is Solid solid && solid.Faces.Size > 0) foreach (Edge e in solid.Edges) visit(e);
            else if (obj is GeometryInstance gi) foreach (GeometryObject sub in gi.GetInstanceGeometry()) TraverseEdges(sub, visit);
        }

        private Edge FindClosestEdge(GeometryElement geo, XYZ point, bool lineOnly = false)
        {
            if (geo == null) return null;
            double minDist = double.MaxValue;
            Edge best = null;

            foreach (GeometryObject obj in geo)
            {
                TraverseEdges(obj, e =>
                {
                    try
                    {
                        if (lineOnly && !(e.AsCurve() is Line)) return;
                        var proj = e.AsCurve().Project(point);
                        if (proj == null) return;
                        double d = proj.XYZPoint.DistanceTo(point);
                        if (d < minDist) { minDist = d; best = e; }
                    }
                    catch { /* Bỏ qua các cạnh bị lỗi hình học */ }
                });
            }
            return best;
        }

        private Reference GetClosestEdgeReference(Element el, XYZ pickPoint, View view)
        {
            try
            {
                var geo = el.get_Geometry(GetGeometryOptions(view));
                return FindClosestEdge(geo, pickPoint)?.Reference;
            }
            catch { return null; }
        }
    }
}