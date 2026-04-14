using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using Autodesk.Revit.UI.Selection;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Reflection;
using TaskDialog = Autodesk.Revit.UI.TaskDialog;

namespace QParam
{
    // ═══════════════════════════════════════════════════════
    //  TẠO RIBBON
    // ═══════════════════════════════════════════════════════
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
                "QParam",
                "Q\nParam",
                assemblyPath,
                "QParam.Class1"
            );

            if (!panel.GetItems().OfType<PushButton>().Any(b => b.Name == "QParam"))
            {
                panel.AddItem(btnData);
            }

            return Result.Succeeded;
        }

        public Result OnShutdown(UIControlledApplication app) => Result.Succeeded;
    }

    // ═══════════════════════════════════════════════════════
    //  CHỌN DIM
    // ═══════════════════════════════════════════════════════
    public class DimSelectionFilter : ISelectionFilter
    {
        public bool AllowElement(Element elem) => elem is Dimension;
        public bool AllowReference(Reference reference, XYZ position) => true;
    }

    // ═══════════════════════════════════════════════════════
    //  HỘP CHỌN PARAM
    // ═══════════════════════════════════════════════════════
    public static class PromptHelper
    {
        public static string ShowDialog(string caption, string label, List<string> existingParams)
        {
            System.Windows.Forms.Form form = new System.Windows.Forms.Form()
            {
                Width = 400,
                Height = 180,
                Text = caption,
                StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen,
                FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedDialog,
                MaximizeBox = false,
                MinimizeBox = false
            };

            System.Windows.Forms.Label textLabel = new System.Windows.Forms.Label()
            {
                Left = 20,
                Top = 20,
                Text = label,
                AutoSize = true
            };

            System.Windows.Forms.ComboBox comboBox = new System.Windows.Forms.ComboBox()
            {
                Left = 20,
                Top = 50,
                Width = 340,
                DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDown
            };

            if (existingParams != null && existingParams.Count > 0)
            {
                comboBox.Items.AddRange(existingParams.ToArray());
                comboBox.SelectedIndex = 0;
            }

            System.Windows.Forms.Button btnOk = new System.Windows.Forms.Button()
            {
                Text = "OK",
                Left = 260,
                Top = 95,
                Width = 100,
                Height = 28,
                DialogResult = System.Windows.Forms.DialogResult.OK
            };

            form.Controls.Add(textLabel);
            form.Controls.Add(comboBox);
            form.Controls.Add(btnOk);
            form.AcceptButton = btnOk;

            if (form.ShowDialog() == System.Windows.Forms.DialogResult.OK && !string.IsNullOrWhiteSpace(comboBox.Text))
            {
                return comboBox.Text.Trim();
            }
            return null;
        }
    }

    // ═══════════════════════════════════════════════════════
    //  LỆNH CHÍNH
    // ═══════════════════════════════════════════════════════
    [Autodesk.Revit.Attributes.Transaction(Autodesk.Revit.Attributes.TransactionMode.Manual)]
    public class Class1 : IExternalCommand
    {
        public Result Execute(ExternalCommandData commandData, ref string message, ElementSet elements)
        {
            UIDocument uiDoc = commandData.Application.ActiveUIDocument;
            Document doc = uiDoc.Document;

            try
            {
                // 1. Chọn DIM
                Reference dimRef;
                try
                {
                    dimRef = uiDoc.Selection.PickObject(ObjectType.Element, new DimSelectionFilter(), "Chọn DIM để gắn Parameter");
                }
                catch (Autodesk.Revit.Exceptions.OperationCanceledException)
                {
                    return Result.Cancelled;
                }

                Dimension dim = doc.GetElement(dimRef) as Dimension;
                if (dim == null)
                {
                    TaskDialog.Show("Lỗi", "Không lấy được DIM đã chọn.");
                    return Result.Failed;
                }

                // 2. Kiểm tra loại DIM
                bool isAngular = IsAngularDimension(doc, dim);

                if (!doc.IsFamilyDocument)
                {
                    TaskDialog.Show("Cảnh báo", "Lệnh này chỉ hoạt động trong Family Editor.");
                    return Result.Failed;
                }

                // 3. Thu thập params phù hợp (dùng var thay vì SpecTypeId/GroupTypeId trực tiếp)
                ForgeTypeId targetSpecType = isAngular ? SpecTypeId.Angle : SpecTypeId.Length;

                List<string> existingParams = new List<string>();
                foreach (FamilyParameter fp in doc.FamilyManager.Parameters)
                {
                    if (fp.Definition.GetDataType() == targetSpecType)
                    {
                        existingParams.Add(fp.Definition.Name);
                    }
                }
                existingParams.Sort();

                // 4. Hỏi user chọn/tạo param
                string paramLabel = isAngular ? "Chọn/tạo Parameter Góc:" : "Chọn/tạo Parameter Chiều Dài:";
                string paramName = PromptHelper.ShowDialog("Gắn Parameter vào DIM", paramLabel, existingParams);

                if (string.IsNullOrWhiteSpace(paramName))
                {
                    return Result.Cancelled;
                }

                // 5. Tìm hoặc tạo param
                using (Transaction tx = new Transaction(doc, "Gắn Parameter vào DIM"))
                {
                    tx.Start();

                    FamilyParameter targetParam = null;
                    foreach (FamilyParameter fp in doc.FamilyManager.Parameters)
                    {
                        if (fp.Definition.Name.Equals(paramName, StringComparison.OrdinalIgnoreCase))
                        {
                            targetParam = fp;
                            break;
                        }
                    }

                    if (targetParam == null)
                    {
                        ForgeTypeId groupType = isAngular ? GroupTypeId.Text : GroupTypeId.Geometry;
                        targetParam = doc.FamilyManager.AddParameter(paramName, groupType, targetSpecType, false);
                    }

                    if (targetParam == null)
                    {
                        tx.RollBack();
                        TaskDialog.Show("Lỗi", "Không thể tạo/gắn Parameter.");
                        return Result.Failed;
                    }

                    dim.FamilyLabel = targetParam;
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

        private bool IsAngularDimension(Document doc, Dimension dim)
        {
            try
            {
                if (dim.DimensionType is DimensionType dimType && dimType.StyleType == DimensionStyleType.Angular)
                    return true;

                if (dim.Curve != null)
                    return dim.Curve is Arc;
            }
            catch { }

            return false;
        }
    }
}