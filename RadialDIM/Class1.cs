using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using Autodesk.Revit.UI.Selection;
using System;
using System.Linq;
// ĐÃ XÓA: using System.Reflection.Metadata; (Để tránh lỗi Ambiguous Document)

namespace RadialDIM
{
    [Autodesk.Revit.Attributes.Transaction(Autodesk.Revit.Attributes.TransactionMode.Manual)]
    public class Command : IExternalCommand
    {
        public Result Execute(ExternalCommandData commandData, ref string message, ElementSet elements)
        {
            UIApplication uiApp = commandData.Application;
            UIDocument uiDoc = uiApp.ActiveUIDocument;
            Document doc = uiDoc.Document; // Đã fix lỗi phân biệt Document

            try
            {
                // 1. CHỌN ĐỐI TƯỢNG (Tường cong hoặc Arc Line)
                Reference pickedRef = uiDoc.Selection.PickObject(
                    ObjectType.Element,
                    new ArcSelectionFilter(),
                    "Bước 1: Click chọn một đối tượng đường cong (Tường cong hoặc Arc Line)");

                if (pickedRef == null) return Result.Cancelled;

                Element el = doc.GetElement(pickedRef);

                // 2. TRÍCH XUẤT REFERENCE TỪ ĐỐI TƯỢNG
                if (!ExtractArcData(el, doc, out Reference arcRef, out Arc geomArc))
                {
                    message = "Không trích xuất được Reference của cung tròn. Vui lòng thử lại.";
                    return Result.Failed;
                }

                // 3. CHỌN VỊ TRÍ ĐẶT DIMENSION
                XYZ placementPoint = uiDoc.Selection.PickPoint("Bước 2: Click chọn vị trí đặt Text của Dimension");

                // 4. TÌM DIMENSION TYPE CHO RADIAL DIM
                DimensionType radDimType = new FilteredElementCollector(doc)
                    .OfClass(typeof(DimensionType))
                    .Cast<DimensionType>()
                    .FirstOrDefault(x => x.StyleType == DimensionStyleType.Radial);

                if (radDimType == null)
                {
                    message = "File Revit không có loại Radial Dimension nào.";
                    return Result.Failed;
                }

                // 5. TẠO RADIAL DIMENSION (TÙY CHỌN MÔI TRƯỜNG/PHIÊN BẢN REVIT)
                using (Transaction tx = new Transaction(doc, "Create Radial Dimension"))
                {
                    tx.Start();

                    View activeView = doc.ActiveView;
                    Dimension radDim = null;

                    /* ⚠️ CHỌN 1 TRONG 2 CÁCH BÊN DƯỚI DỰA THEO NHU CẦU CỦA BẠN ⚠️ */

                    // CÁCH 1: NẾU BẠN DÙNG REVIT 2025 TRỞ LÊN (Cho môi trường Project & Family)
                    // (Nếu bạn dùng Revit <= 2024, dòng này sẽ bị lỗi đỏ, hãy xóa nó đi)
                    // radDim = RadialDimension.Create(doc, activeView, arcRef, placementPoint, false);


                    // CÁCH 2: NẾU BẠN CHẠY PLUGIN TRONG MÔI TRƯỜNG TẠO FAMILY (.rfa) CHUNG CHO CÁC ĐỜI REVIT
                    radDim = doc.FamilyCreate.NewRadialDimension(activeView, arcRef, placementPoint);


                    // Gán loại (Type) cho Dimension vừa tạo
                    if (radDim != null && radDimType != null)
                    {
                        radDim.DimensionType = radDimType;
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

        // --- HÀM HELPER: TRÍCH XUẤT DỮ LIỆU CUNG TRÒN ---
        private bool ExtractArcData(Element el, Document doc, out Reference refObj, out Arc geomArc)
        {
            refObj = null;
            geomArc = null;

            if (el is Wall wall)
            {
                LocationCurve locCurve = wall.Location as LocationCurve;
                if (locCurve != null && locCurve.Curve is Arc a) geomArc = a;

                System.Collections.Generic.IList<Reference> exteriorFaces = HostObjectUtils.GetSideFaces(wall, ShellLayerType.Exterior);
                if (exteriorFaces != null && exteriorFaces.Count > 0)
                {
                    refObj = exteriorFaces[0];
                }

                return refObj != null && geomArc != null;
            }
            else if (el is CurveElement ce && ce.GeometryCurve is Arc a)
            {
                geomArc = a;
                Options opt = new Options { ComputeReferences = true, View = doc.ActiveView };
                GeometryElement geo = ce.get_Geometry(opt);
                foreach (GeometryObject obj in geo)
                {
                    if (obj is Arc geoArc && geoArc.Reference != null)
                    {
                        refObj = geoArc.Reference;
                        break;
                    }
                }
                return refObj != null;
            }

            return false;
        }
    }

    public class ArcSelectionFilter : ISelectionFilter
    {
        public bool AllowElement(Element elem)
        {
            if (elem is CurveElement ce && ce.GeometryCurve is Arc) return true;
            if (elem is Wall wall && wall.Location is LocationCurve lc && lc.Curve is Arc) return true;
            return false;
        }

        public bool AllowReference(Reference reference, XYZ position)
        {
            return false;
        }
    }
}