using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using Autodesk.Revit.UI.Selection;
using System;
using System.Linq;

namespace LinearDIM
{
    [Autodesk.Revit.Attributes.Transaction(Autodesk.Revit.Attributes.TransactionMode.Manual)]
    public class Command : IExternalCommand
    {
        public Result Execute(ExternalCommandData commandData, ref string message, ElementSet elements)
        {
            UIDocument uiDoc = commandData.Application.ActiveUIDocument;
            Document doc = uiDoc.Document;
            View activeView = doc.ActiveView;

            try
            {
                ReferenceArray refArray = new ReferenceArray();
                XYZ pt1 = null;
                XYZ pt2 = null;

                // --- BƯỚC 1: CHỌN VÀ TỰ ĐỘNG BẮT ĐIỂM ĐẦU MÚT ---
                for (int i = 1; i <= 2; i++)
                {
                    // Cho phép chọn Element bình thường, nhưng yêu cầu click gần góc
                    Reference pickedRef = uiDoc.Selection.PickObject(ObjectType.Element,
                        $"Chọn đối tượng thứ {i} (Click chuột GẦN GÓC / ĐẦU MÚT cần đo)...");

                    Element el = doc.GetElement(pickedRef);

                    // Lấy tọa độ click chuột thực tế của người dùng
                    XYZ clickPt = pickedRef.GlobalPoint;

                    // Hàm quét hình học: Tìm Reference của Điểm (Vertex) gần vị trí click chuột nhất
                    Reference pointRef = GetNearestVertexReference(el, clickPt, out XYZ vertexPt);

                    if (pointRef == null)
                    {
                        TaskDialog.Show("Lỗi", $"Không tìm thấy điểm đầu mút nào gần vị trí click trên đối tượng {i}.");
                        return Result.Failed;
                    }

                    refArray.Append(pointRef);

                    if (i == 1) pt1 = vertexPt;
                    else pt2 = vertexPt;
                }

                // --- BƯỚC 2: HIỂN THỊ PROMPT HỎI HƯỚNG ĐO ---
                TaskDialog dialog = new TaskDialog("Khóa hướng DIM Linear");
                dialog.MainInstruction = "Bạn muốn đo khoảng cách theo trục nào?";
                dialog.AddCommandLink(TaskDialogCommandLinkId.CommandLink1, "Đo phương NGANG (Trục X)");
                dialog.AddCommandLink(TaskDialogCommandLinkId.CommandLink2, "Đo phương DỌC (Trục Y)");
                dialog.CommonButtons = TaskDialogCommonButtons.Cancel;
                dialog.DefaultButton = TaskDialogResult.CommandLink1;

                TaskDialogResult result = dialog.Show();
                if (result == TaskDialogResult.Cancel || result == TaskDialogResult.Close)
                {
                    return Result.Cancelled;
                }

                bool isXAxis = (result == TaskDialogResult.CommandLink1);

                // --- BƯỚC 3: CHỌN ĐIỂM ĐẶT (CHỈ ĐỂ DI CHUYỂN, KHÔNG ĐỔI HƯỚNG) ---
                string placementPrompt = isXAxis
                    ? "Click để chọn CAO ĐỘ đặt đường DIM ngang..."
                    : "Click để chọn vị trí TRÁI/PHẢI đặt đường DIM dọc...";

                XYZ pt3 = uiDoc.Selection.PickPoint(placementPrompt);

                using (Transaction tx = new Transaction(doc, "Auto Linear DIM Locked to Endpoints"))
                {
                    tx.Start();

                    // --- BƯỚC 4: TÍNH TOÁN ĐƯỜNG DIM KHÓA TRỤC ---
                    XYZ viewRight = activeView.RightDirection; // Trục X
                    XYZ viewUp = activeView.UpDirection;       // Trục Y
                    XYZ p1, p2;

                    if (isXAxis)
                    {
                        p1 = pt3 - viewRight * 100.0;
                        p2 = pt3 + viewRight * 100.0;
                    }
                    else
                    {
                        p1 = pt3 - viewUp * 100.0;
                        p2 = pt3 + viewUp * 100.0;
                    }

                    Line dimLine = Line.CreateBound(p1, p2);

                    DimensionType linearDimType = new FilteredElementCollector(doc)
                        .OfClass(typeof(DimensionType))
                        .Cast<DimensionType>()
                        .FirstOrDefault(dt => dt.StyleType == DimensionStyleType.Linear);

                    if (linearDimType != null)
                    {
                        doc.Create.NewDimension(activeView, dimLine, refArray, linearDimType);
                    }
                    else
                    {
                        doc.Create.NewDimension(activeView, dimLine, refArray);
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

        // --- HÀM CỐT LÕI MỚI: TÌM THAM CHIẾU CỦA ĐIỂM ĐẦU MÚT GẦN NHẤT ---
        private Reference GetNearestVertexReference(Element el, XYZ clickPt, out XYZ nearestVertexPt)
        {
            Reference nearestRef = null;
            double minDistance = double.MaxValue;
            nearestVertexPt = null;

            Options opt = new Options
            {
                ComputeReferences = true, // Bắt buộc phải = true để lấy được Reference
                IncludeNonVisibleObjects = false,
                DetailLevel = ViewDetailLevel.Fine
            };

            GeometryElement geoElem = el.get_Geometry(opt);
            if (geoElem != null)
            {
                ScanForVertices(geoElem, clickPt, ref nearestRef, ref nearestVertexPt, ref minDistance);
            }

            return nearestRef;
        }

        private void ScanForVertices(GeometryElement geoElem, XYZ clickPt, ref Reference nearestRef, ref XYZ nearestVertexPt, ref double minDistance)
        {
            foreach (GeometryObject obj in geoElem)
            {
                // Nếu đối tượng có khối Solid (Tường, Dầm, Cột...)
                if (obj is Solid solid && solid.Edges.Size > 0)
                {
                    foreach (Edge edge in solid.Edges)
                    {
                        Curve curve = edge.AsCurve();
                        // Chỉ kiểm tra 2 điểm đầu mút (Endpoint 0 và 1) của cạnh
                        for (int i = 0; i < 2; i++)
                        {
                            XYZ pt = curve.GetEndPoint(i);
                            double dist = pt.DistanceTo(clickPt);
                            if (dist < minDistance)
                            {
                                Reference r = edge.GetEndPointReference(i); // ÉP LẤY POINT REFERENCE
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
                // Nếu là khối Family Instance lồng nhau
                else if (obj is GeometryInstance inst)
                {
                    ScanForVertices(inst.GetInstanceGeometry(), clickPt, ref nearestRef, ref nearestVertexPt, ref minDistance);
                }
                // Nếu đối tượng là đường nét (Grid, Line...)
                else if (obj is Line line)
                {
                    for (int i = 0; i < 2; i++)
                    {
                        XYZ pt = line.GetEndPoint(i);
                        double dist = pt.DistanceTo(clickPt);
                        if (dist < minDistance)
                        {
                            Reference r = line.GetEndPointReference(i); // ÉP LẤY POINT REFERENCE
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
    }
}