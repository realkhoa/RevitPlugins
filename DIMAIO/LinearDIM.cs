using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using Autodesk.Revit.UI.Selection;
using View = Autodesk.Revit.DB.View;

namespace DIMAIO
{
    [Autodesk.Revit.Attributes.Transaction(Autodesk.Revit.Attributes.TransactionMode.Manual)]
    public class LinearDIMCommand : IExternalCommand
    {
        public Result Execute(ExternalCommandData commandData, ref string message, ElementSet elements)
        {
            UIApplication uiApp = commandData.Application;
            UIDocument uiDoc = uiApp.ActiveUIDocument;
            Document doc = uiDoc.Document;
            View view = doc.ActiveView;

            // ✋ BƯỚC 1: HỎI USER CHỌN CHẾ ĐỘ
            Autodesk.Revit.UI.TaskDialog modeDialog = new Autodesk.Revit.UI.TaskDialog("Chọn chế độ Linear DIM");
            modeDialog.MainInstruction = "Bạn muốn chọn bằng cách nào?";
            modeDialog.AddCommandLink(TaskDialogCommandLinkId.CommandLink1, "Chọn 2 THAM CHIẾU (Reference) trên đối tượng");
            modeDialog.AddCommandLink(TaskDialogCommandLinkId.CommandLink2, "Chọn 2 ĐIỂM (Point) tự do trong không gian");
            modeDialog.CommonButtons = TaskDialogCommonButtons.Cancel;
            modeDialog.DefaultButton = TaskDialogResult.CommandLink1;

            TaskDialogResult modeResult = modeDialog.Show();
            if (modeResult == TaskDialogResult.Cancel || modeResult == TaskDialogResult.Close)
                return Result.Cancelled;

            bool useReferences = (modeResult == TaskDialogResult.CommandLink1);

            try
            {
                if (useReferences)
                {
                    Reference ref1 = uiDoc.Selection.PickObject(ObjectType.Element,
                        "Chọn tham chiếu thứ 1 (click gần góc/điểm cần đo)");
                    Reference ref2 = uiDoc.Selection.PickObject(ObjectType.Element,
                        "Chọn tham chiếu thứ 2 (click gần góc/điểm cần đo)");

                    XYZ pt1 = ref1.GlobalPoint;
                    XYZ pt2 = ref2.GlobalPoint;

                    XYZ placePt = uiDoc.Selection.PickPoint(
                        "Click chọn vị trí đặt DIM line...");

                    return RunLinearDim(uiDoc, doc, view, ref1, ref2, pt1, pt2, placePt, ref message);
                }
                else
                {
                    XYZ pt1 = uiDoc.Selection.PickPoint("Chọn điểm thứ 1");
                    XYZ pt2 = uiDoc.Selection.PickPoint("Chọn điểm thứ 2");

                    XYZ placePt = uiDoc.Selection.PickPoint(
                        "Click chọn vị trí đặt DIM line...");

                    return RunLinearDimPoint(uiDoc, doc, view, pt1, pt2, placePt, ref message);
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

        // ========== CHẾ ĐỘ REFERENCE ==========
        private Result RunLinearDim(
            UIDocument uiDoc,
            Document doc,
            View view,
            Reference ref1,
            Reference ref2,
            XYZ pt1,
            XYZ pt2,
            XYZ placePt,
            ref string message)
        {
            try
            {
                using (Transaction tx = new Transaction(doc, "Linear DIM"))
                {
                    tx.Start();

                    // ========== TẠO REFERENCE PLANE TỪ VIEW ==========
                    Transform refPlane = BuildReferencePlane(view);

                    // ========== PROJECT LÊN REFERENCE PLANE ==========
                    XYZ proj1 = ProjectPointToRefPlane(pt1, refPlane);
                    XYZ proj2 = ProjectPointToRefPlane(pt2, refPlane);
                    XYZ projPlace = ProjectPointToRefPlane(placePt, refPlane);

                    // ========== TỰ ĐỘNG XÁC ĐỊNH HƯỚNG DIM ==========
                    bool lockToX = AutoDetectDirection(refPlane, proj1, proj2, projPlace);

                    // ========== TẠO DIM LINE ==========
                    XYZ direction = lockToX ? refPlane.BasisX : refPlane.BasisY;

                    XYZ dimP1 = placePt - direction * 1000.0;
                    XYZ dimP2 = placePt + direction * 1000.0;
                    Line dimLine = Line.CreateBound(dimP1, dimP2);

                    // ========== TẠO FACE REFERENCE ARRAY ==========
                    // Key fix: dùng Face Reference thay vì Edge Endpoint Reference.
                    // Face Reference có Face.Normal — Revit dùng normal này làm
                    // constraint direction, đảm bảo Linear DIM lock đúng trục khi grip.
                    ReferenceArray refArray = new ReferenceArray();

                    // Lấy face reference từ element đã pick
                    Reference faceRef1 = GetFaceReferenceForDirection(doc, ref1, pt1, refPlane, direction);
                    Reference faceRef2 = GetFaceReferenceForDirection(doc, ref2, pt2, refPlane, direction);

                    if (faceRef1 != null && faceRef2 != null)
                    {
                        refArray.Append(faceRef1);
                        refArray.Append(faceRef2);
                    }
                    else
                    {
                        // Fallback: dùng edge endpoint reference
                        Reference pointRef1 = GetNearestVertexReference(doc, ref1, pt1);
                        Reference pointRef2 = GetNearestVertexReference(doc, ref2, pt2);

                        if (pointRef1 == null || pointRef2 == null)
                        {
                            message = "Không lấy được Reference từ đối tượng đã chọn.";
                            tx.RollBack();
                            return Result.Failed;
                        }

                        refArray.Append(pointRef1);
                        refArray.Append(pointRef2);
                    }

                    // ========== TẠO LINEAR DIMENSION ==========
                    Dimension dim = CreateLinearDimension(doc, view, dimLine, refArray);
                    if (dim == null)
                    {
                        message = "Không tạo được Linear Dimension.";
                        tx.RollBack();
                        return Result.Failed;
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

        // ========== CHẾ ĐỘ POINT TỰ DO ==========
        private Result RunLinearDimPoint(
            UIDocument uiDoc,
            Document doc,
            View view,
            XYZ pt1,
            XYZ pt2,
            XYZ placePt,
            ref string message)
        {
            try
            {
                using (Transaction tx = new Transaction(doc, "Linear DIM"))
                {
                    tx.Start();

                    Transform refPlane = BuildReferencePlane(view);

                    XYZ proj1 = ProjectPointToRefPlane(pt1, refPlane);
                    XYZ proj2 = ProjectPointToRefPlane(pt2, refPlane);
                    XYZ projPlace = ProjectPointToRefPlane(placePt, refPlane);

                    bool lockToX = AutoDetectDirection(refPlane, proj1, proj2, projPlace);

                    XYZ direction = lockToX ? refPlane.BasisX : refPlane.BasisY;

                    XYZ dimP1 = placePt - direction * 1000.0;
                    XYZ dimP2 = placePt + direction * 1000.0;
                    Line dimLine = Line.CreateBound(dimP1, dimP2);

                    // Tìm face reference từ geometry gần điểm pick
                    Reference faceRef1 = GetFaceReferenceFromPosition(doc, view, pt1, refPlane, direction);
                    Reference faceRef2 = GetFaceReferenceFromPosition(doc, view, pt2, refPlane, direction);

                    ReferenceArray refArray = new ReferenceArray();

                    if (faceRef1 != null && faceRef2 != null)
                    {
                        refArray.Append(faceRef1);
                        refArray.Append(faceRef2);
                    }
                    else
                    {
                        Reference pointRef1 = GetPointReferenceFromPosition(doc, view, pt1);
                        Reference pointRef2 = GetPointReferenceFromPosition(doc, view, pt2);

                        if (pointRef1 == null || pointRef2 == null)
                        {
                            message = "Không lấy được Reference từ điểm đã chọn. Thử chế độ Reference.";
                            tx.RollBack();
                            return Result.Failed;
                        }

                        refArray.Append(pointRef1);
                        refArray.Append(pointRef2);
                    }

                    Dimension dim = CreateLinearDimension(doc, view, dimLine, refArray);
                    if (dim == null)
                    {
                        message = "Không tạo được Linear Dimension.";
                        tx.RollBack();
                        return Result.Failed;
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

        // ========== TẠO FACE REFERENCE PHÙ HỢP VỚI HƯỚNG DIM ==========
        // Face normal: vector vuông góc với mặt face.
        // Khi Revit tạo Linear DIM, nó dùng face normal làm constraint:
        //   - DIM line phải vuông góc với face normal → nằm trong mặt phẳng face
        //   - Nếu face nằm trong reference plane → DIM line nằm trong reference plane → Linear DIM
        // Để DIM đo ngang (song song BasisX): tìm face có normal song song BasisY (vuông góc với BasisX)
        // Để DIM đo dọc (song song BasisY): tìm face có normal song song BasisX (vuông góc với BasisY)
        private Reference GetFaceReferenceForDirection(
            Document doc,
            Reference pickedRef,
            XYZ clickPt,
            Transform refPlane,
            XYZ dimDirection)
        {
            Element el = doc.GetElement(pickedRef);
            if (el == null) return null;

            // dimDirection là hướng của DIM line (song song BasisX hoặc BasisY)
            // Face cần tìm: normal vuông góc với dimDirection
            // Nghĩa là face nằm trong mặt phẳng chứa dimDirection
            XYZ perpDir = dimDirection; // face normal cần vuông góc với dimDirection

            Options opt = new Options
            {
                ComputeReferences = true,
                IncludeNonVisibleObjects = false,
                DetailLevel = ViewDetailLevel.Fine
            };

            GeometryElement geoElem = el.get_Geometry(opt);
            if (geoElem == null) return null;

            Reference bestFaceRef = null;
            double bestScore = double.MaxValue;

            foreach (GeometryObject obj in geoElem)
            {
                Reference found = FindFaceInGeometry(doc, obj, pickedRef, clickPt, refPlane, perpDir, ref bestScore);
                if (found != null) bestFaceRef = found;
            }

            return bestFaceRef;
        }

        private Reference FindFaceInGeometry(
            Document doc,
            GeometryObject obj,
            Reference pickedRef,
            XYZ clickPt,
            Transform refPlane,
            XYZ perpDir,
            ref double bestScore)
        {
            Reference bestFaceRef = null;

            if (obj is Solid solid)
            {
                foreach (Face face in solid.Faces)
                {
                    Reference faceRef = face.Reference;
                    if (faceRef == null) continue;

                    // Kiểm tra face có gần điểm click không
                    UV center = GetFaceCenter(face);
                    XYZ facePt = face.Evaluate(center);
                    double dist = facePt.DistanceTo(clickPt);

                    // Lấy normal của face tại tâm bằng ComputeNormal
                    XYZ normal = face.ComputeNormal(center);
                    if (normal == null || normal.GetLength() < 1e-6) continue;
                    normal = normal.Normalize();

                    // Chuyển sang world coordinate system nếu cần
                    // (face normal đã là world coordinate)

                    // Tính độ song song giữa face normal và perpDir
                    double dot = Math.Abs(normal.DotProduct(perpDir));

                    // Score: ưu tiên face gần click + có normal vuông góc với perpDir
                    // dot gần 1 = normal vuông góc với perpDir (tốt)
                    // dot gần 0 = normal // perpDir (không tốt)
                    double score = dist + (1.0 - dot) * 10.0;

                    if (score < bestScore && dot > 0.3)
                    {
                        bestScore = score;
                        bestFaceRef = faceRef;
                    }
                }
            }
            else if (obj is GeometryInstance inst)
            {
                Transform instTransform = inst.Transform;
                GeometryElement instGeo = inst.GetInstanceGeometry(instTransform);

                foreach (GeometryObject iobj in instGeo)
                {
                    Reference found = FindFaceInGeometry(doc, iobj, pickedRef, clickPt, refPlane, perpDir, ref bestScore);
                    if (found != null) bestFaceRef = found;
                }
            }

            return bestFaceRef;
        }

        private UV GetFaceCenter(Face face)
        {
            BoundingBoxUV bbox = face.GetBoundingBox();
            return new UV(
                (bbox.Min.U + bbox.Max.U) * 0.5,
                (bbox.Min.V + bbox.Max.V) * 0.5);
        }

        // ========== TÌM FACE REFERENCE TỪ VỊ TRÍ XYZ (cho Point mode) ==========
        private Reference GetFaceReferenceFromPosition(
            Document doc,
            View view,
            XYZ point,
            Transform refPlane,
            XYZ dimDirection)
        {
            Outline outline = new Outline(
                point - new XYZ(0.5, 0.5, 0.5),
                point + new XYZ(0.5, 0.5, 0.5));

            BoundingBoxIntersectsFilter bbFilter = new BoundingBoxIntersectsFilter(outline);

            FilteredElementCollector collector = new FilteredElementCollector(doc, view.Id)
                .WherePasses(bbFilter);

            Reference bestFaceRef = null;
            double bestScore = double.MaxValue;

            foreach (Element el in collector)
            {
                if (el.Category == null) continue;

                Options opt = new Options
                {
                    ComputeReferences = true,
                    DetailLevel = ViewDetailLevel.Fine
                };

                GeometryElement geoElem = el.get_Geometry(opt);
                if (geoElem == null) continue;

                foreach (GeometryObject obj in geoElem)
                {
                    Reference found = FindFaceInGeometry(doc, obj, null, point, refPlane, dimDirection, ref bestScore);
                    if (found != null) bestFaceRef = found;
                }
            }

            return bestFaceRef;
        }

        // ========== TỰ ĐỘNG XÁC ĐỊNH HƯỚNG DIM ==========
        private bool AutoDetectDirection(Transform refPlane, XYZ p1, XYZ p2, XYZ placePt)
        {
            XYZ midpoint = (p1 + p2) * 0.5;
            XYZ offset = placePt - midpoint;
            XYZ offsetInPlane = refPlane.Inverse.OfVector(offset);

            double offsetX = Math.Abs(offsetInPlane.X);
            double offsetY = Math.Abs(offsetInPlane.Y);

            return offsetY >= offsetX;
        }

        // ========== TẠO LINEAR DIMENSION ==========
        private Dimension CreateLinearDimension(Document doc, View view, Line dimLine, ReferenceArray refArray)
        {
            DimensionType dimType = new FilteredElementCollector(doc)
                .OfClass(typeof(DimensionType))
                .Cast<DimensionType>()
                .FirstOrDefault(dt => dt.StyleType == DimensionStyleType.Linear);

            if (dimType != null)
                return doc.Create.NewDimension(view, dimLine, refArray, dimType);

            return doc.Create.NewDimension(view, dimLine, refArray);
        }

        // ========== TẠO REFERENCE PLANE TỪ VIEW ==========
        private Transform BuildReferencePlane(View view)
        {
            Transform t = Transform.Identity;
            t.Origin = view.Origin;
            t.BasisX = view.RightDirection;
            t.BasisY = view.UpDirection;
            t.BasisZ = view.ViewDirection;
            return t;
        }

        // ========== PROJECT ĐIỂM LÊN REFERENCE PLANE ==========
        private XYZ ProjectPointToRefPlane(XYZ point, Transform refPlane)
        {
            XYZ planeNormal = refPlane.BasisZ;
            XYZ planeOrigin = refPlane.Origin;
            double dist = (point - planeOrigin).DotProduct(planeNormal);
            return point - dist * planeNormal;
        }

        // ========== HELPER: Tìm vertex reference gần nhất (fallback) ==========
        private Reference GetNearestVertexReference(Document doc, Reference pickedRef, XYZ clickPt)
        {
            Reference nearestRef = null;
            double minDistance = double.MaxValue;
            XYZ _ = null;

            Element el = doc.GetElement(pickedRef);
            if (el == null) return null;

            Options opt = new Options
            {
                ComputeReferences = true,
                IncludeNonVisibleObjects = false,
                DetailLevel = ViewDetailLevel.Fine
            };

            GeometryElement geoElem = el.get_Geometry(opt);
            if (geoElem == null) return null;

            ScanForVertices(geoElem, clickPt, ref nearestRef, ref _, ref minDistance);

            return nearestRef;
        }

        private void ScanForVertices(
            GeometryElement geoElem,
            XYZ clickPt,
            ref Reference nearestRef,
            ref XYZ nearestVertexPt,
            ref double minDistance)
        {
            foreach (GeometryObject obj in geoElem)
            {
                if (obj is Solid solid && solid.Edges.Size > 0)
                {
                    foreach (Edge edge in solid.Edges)
                    {
                        Curve curve = edge.AsCurve();
                        for (int i = 0; i < 2; i++)
                        {
                            XYZ pt = curve.GetEndPoint(i);
                            double dist = pt.DistanceTo(clickPt);
                            if (dist < minDistance)
                            {
                                Reference r = edge.GetEndPointReference(i);
                                if (r != null)
                                {
                                    minDistance = dist;
                                    nearestRef = r;
                                    nearestVertexPt = pt;
                                }
                            }
                        }
                    }
                }
                else if (obj is GeometryInstance inst)
                {
                    ScanForVertices(inst.GetInstanceGeometry(), clickPt,
                        ref nearestRef, ref nearestVertexPt, ref minDistance);
                }
                else if (obj is Line line)
                {
                    for (int i = 0; i < 2; i++)
                    {
                        XYZ pt = line.GetEndPoint(i);
                        double dist = pt.DistanceTo(clickPt);
                        if (dist < minDistance)
                        {
                            Reference r = line.GetEndPointReference(i);
                            if (r != null)
                            {
                                minDistance = dist;
                                nearestRef = r;
                                nearestVertexPt = pt;
                            }
                        }
                    }
                }
            }
        }

        // ========== HELPER: Tìm point reference từ vị trí XYZ (fallback cho Point mode) ==========
        private Reference GetPointReferenceFromPosition(Document doc, View view, XYZ point)
        {
            Outline outline = new Outline(
                point - new XYZ(0.5, 0.5, 0.5),
                point + new XYZ(0.5, 0.5, 0.5));

            BoundingBoxIntersectsFilter bbFilter = new BoundingBoxIntersectsFilter(outline);

            FilteredElementCollector collector = new FilteredElementCollector(doc, view.Id)
                .WherePasses(bbFilter);

            Reference nearestRef = null;
            double minDist = double.MaxValue;

            foreach (Element el in collector)
            {
                if (el.Category == null) continue;

                Options opt = new Options
                {
                    ComputeReferences = true,
                    DetailLevel = ViewDetailLevel.Fine
                };

                GeometryElement geoElem = el.get_Geometry(opt);
                if (geoElem == null) continue;

                foreach (GeometryObject obj in geoElem)
                {
                    ScanSolidEdgesForPointRef(obj, point, ref nearestRef, ref minDist);

                    if (obj is GeometryInstance inst)
                    {
                        foreach (GeometryObject iobj in inst.GetInstanceGeometry())
                        {
                            ScanSolidEdgesForPointRef(iobj, point, ref nearestRef, ref minDist);
                        }
                    }
                }
            }

            return nearestRef;
        }

        private void ScanSolidEdgesForPointRef(
            GeometryObject obj,
            XYZ point,
            ref Reference nearestRef,
            ref double minDist)
        {
            if (!(obj is Solid solid) || solid.Edges.Size <= 0) return;

            foreach (Edge edge in solid.Edges)
            {
                Curve curve = edge.AsCurve();
                for (int i = 0; i < 2; i++)
                {
                    XYZ pt = curve.GetEndPoint(i);
                    double dist = pt.DistanceTo(point);
                    if (dist < minDist && dist < 0.3)
                    {
                        Reference r = edge.GetEndPointReference(i);
                        if (r != null)
                        {
                            minDist = dist;
                            nearestRef = r;
                        }
                    }
                }
            }
        }
    }
}
