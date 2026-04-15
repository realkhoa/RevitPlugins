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

                if (doc.IsFamilyDocument)
                {
                    return ExecuteInFamily(doc, dim, isAngular);
                }
                else
                {
                    return ExecuteInProject(doc, dim, isAngular);
                }
            }
            catch (Exception ex)
            {
                message = ex.Message;
                return Result.Failed;
            }
        }

        // ─────────────────────────────────────────────────────
        //  FAMILY DOCUMENT
        // ─────────────────────────────────────────────────────
        private Result ExecuteInFamily(Document doc, Dimension dim, bool isAngular)
        {
            ForgeTypeId targetSpecType = isAngular ? SpecTypeId.Angle : SpecTypeId.Length;

            // Thu thập params phù hợp
            List<string> existingParams = new List<string>();
            foreach (FamilyParameter fp in doc.FamilyManager.Parameters)
            {
                if (fp.Definition.GetDataType() == targetSpecType)
                {
                    existingParams.Add(fp.Definition.Name);
                }
            }
            existingParams.Sort();

            // Hỏi user chọn/tạo param
            string paramLabel = isAngular ? "Chọn/tạo Parameter Góc:" : "Chọn/tạo Parameter Chiều Dài:";
            string paramName = PromptHelper.ShowDialog("Gắn Parameter vào DIM", paramLabel, existingParams);

            if (string.IsNullOrWhiteSpace(paramName))
                return Result.Cancelled;

            // Tìm hoặc tạo param
            using (Transaction tx = new Transaction(doc, "Gắn Family Parameter vào DIM"))
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

        // ─────────────────────────────────────────────────────
        //  PROJECT DOCUMENT
        //  Uses dynamic to probe Label property at runtime
        // ─────────────────────────────────────────────────────
        private Result ExecuteInProject(Document doc, Dimension dim, bool isAngular)
        {
            ForgeTypeId targetSpecType = isAngular ? SpecTypeId.Angle : SpecTypeId.Length;

            // Probe dim.Label at runtime — may not exist in all API versions
            dynamic dimLabelProp = null;
            string dimLabelValue = "n/a";
            try
            {
                dimLabelProp = ((dynamic)dim).Label;
                dimLabelValue = dimLabelProp?.ToString() ?? "null";
            }
            catch (Exception ex)
            {
                dimLabelValue = $"ERROR: {ex.Message}";
            }

            // Collect ALL shared params of matching spec type via BindingMap
            // BindingMap is the RIGHT way — its Definition.Id is ElementId
            List<Definition> existingDefs = new List<Definition>();
            BindingMap bindingMap = doc.ParameterBindings;
            DefinitionBindingMapIterator it = bindingMap.ForwardIterator();
            it.Reset();
            while (it.MoveNext())
            {
                Definition def = it.Key;
                if (def.GetDataType() == targetSpecType)
                {
                    existingDefs.Add(def);
                }
            }
            existingDefs.Sort((a, b) => string.Compare(a.Name, b.Name, StringComparison.OrdinalIgnoreCase));

            // Also enumerate directly from shared param file (for reference)
            List<string> spFileParamNames = new List<string>();
            try
            {
                dynamic spFile = doc.Application.OpenSharedParameterFile();
                if (spFile != null)
                {
                    foreach (dynamic grp in spFile.Groups)
                    {
                        foreach (dynamic d in grp.Definitions)
                        {
                            if (d.GetDataType() == targetSpecType)
                            {
                                string n = d.Name;
                                if (!spFileParamNames.Contains(n)) spFileParamNames.Add(n);
                            }
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                spFileParamNames.Add($"Error: {ex.Message}");
            }
            spFileParamNames.Sort();

            // Get ElementId of each def from BindingMap using dynamic
            string debugDefs = "";
            foreach (var def in existingDefs)
            {
                try
                {
                    dynamic d = def;
                    object idObj = d.Id;
                    debugDefs += $"  {def.Name}: Id type={idObj?.GetType().Name}, value={idObj}\n";
                }
                catch (Exception ex)
                {
                    debugDefs += $"  {def.Name}: Id ERROR={ex.Message}\n";
                }
            }

            string debug1 =
                $"[DEBUG] Kiểu DIM: {(isAngular ? "Angular" : "Linear")}\n" +
                $"[DEBUG] dim.Label property: {(dimLabelProp != null ? "CÓ" : "KHÔNG (bỏ qua)")}\n" +
                $"[DEBUG] DimensionType: {dim.DimensionType?.Name ?? "null"}\n" +
                $"───────────BindingMap Definitions───────────\n" +
                $"Số params: {existingDefs.Count}\n" +
                $"{debugDefs}" +
                $"───────────SP File (tham khảo)───────────\n" +
                $"{string.Join("\n", spFileParamNames)}";
            TaskDialog.Show("DEBUG 1", debug1);

            // Ask user to select or type a new param name
            List<string> allNames = existingDefs.Select(d => d.Name)
                .Concat(spFileParamNames)
                .Distinct()
                .OrderBy(n => n)
                .ToList();

            string paramLabel = isAngular ? "Chọn Parameter Góc:" : "Chọn Parameter Chiều Dài:";
            string paramName = PromptHelper.ShowDialog("Gắn Parameter vào DIM", paramLabel, allNames);

            if (string.IsNullOrWhiteSpace(paramName))
                return Result.Cancelled;

            // Resolve definition — always from BindingMap (guarantees Definition with ElementId)
            Definition targetDef = existingDefs.FirstOrDefault(
                d => d.Name.Equals(paramName, StringComparison.OrdinalIgnoreCase));

            string targetDefInfo = "CHƯA TÌM THẤY trong BindingMap";
            ElementId targetId = null;

            if (targetDef != null)
            {
                // Get ElementId from definition (BindingMap defs have accessible .Id)
                try
                {
                    dynamic tdDyn = targetDef;
                    object idObj = tdDyn.Id;
                    if (idObj is ElementId eid)
                    {
                        targetId = eid;
                        targetDefInfo = $"TÌM THẤY - Name={targetDef.Name}, Id={targetId}";
                    }
                    else
                    {
                        targetDefInfo = $"TÌM THẤY nhưng Id KHÔNG PHẢI ElementId: type={idObj?.GetType().Name}";
                    }
                }
                catch (Exception ex)
                {
                    targetDefInfo = $"TÌM THẤY nhưng lỗi Id: {ex.Message}";
                }
            }
            else
            {
                // Not in BindingMap yet — create shared param, then look it up in BindingMap
                targetDefInfo = "ĐANG TẠO mới trong SP file...";
                TaskDialog.Show("DEBUG 2", targetDefInfo);

                bool created = CreateSharedParameter(doc, paramName, targetSpecType);
                if (created)
                {
                    // Look up the newly-created param in BindingMap
                    targetDef = FindDefinitionInBindingMap(doc, paramName, targetSpecType, out targetId);
                    if (targetDef != null)
                    {
                        targetDefInfo = $"ĐÃ TẠO - tìm thấy trong BindingMap: Name={targetDef.Name}, Id={targetId}";
                    }
                    else
                    {
                        targetDefInfo = "ĐÃ TẠO trong SP file nhưng KHÔNG tìm thấy trong BindingMap!";
                    }
                }
                else
                {
                    targetDefInfo = "TẠO THẤT BẠI";
                }
            }

            TaskDialog.Show("DEBUG 3 - Definition", $"Definition: {targetDefInfo}\nTargetId: {targetId?.ToString() ?? "null"}");

            if (targetDef == null || targetId == null)
            {
                TaskDialog.Show("Lỗi", $"Không thể lấy Definition hoặc Id.\n{targetDefInfo}");
                return Result.Failed;
            }

            // Try to assign label
            string assignResult = "CHƯA CHẠY";
            using (Transaction tx = new Transaction(doc, "Gắn Parameter vào DIM"))
            {
                tx.Start();
                try
                {
                    if (dimLabelProp != null)
                    {
                        try
                        {
                            ((dynamic)dim).Label = targetId;
                            assignResult = $"OK - gán via dim.Label = {targetId}";
                        }
                        catch (Exception ex2)
                        {
                            assignResult = $"dim.Label set THẤT BẠI: {ex2.Message}";
                            TaskDialog.Show("DEBUG 4 - dim.Label failed", assignResult);
                            AssignLabelViaDimensionType(doc, dim, targetDef, targetId);
                            assignResult += "\n→ Đã thử DimensionType";
                        }
                    }
                    else
                    {
                        assignResult = "dim.Label KHÔNG TỒN TẠI";
                        TaskDialog.Show("DEBUG 4 - Không có dim.Label", assignResult);
                        AssignLabelViaDimensionType(doc, dim, targetDef, targetId);
                        assignResult += "\n→ Đã thử DimensionType";
                    }
                }
                catch (Exception ex)
                {
                    tx.RollBack();
                    TaskDialog.Show("Lỗi", $"Không thể gắn Parameter: {ex.Message}");
                    return Result.Failed;
                }
                tx.Commit();
            }

            TaskDialog.Show("KẾT QUẢ", $"Gán: {assignResult}\n\nNếu không thấy thay đổi trên DIM,\nAPI không hỗ trợ gán label kiểu này.");
            return Result.Succeeded;
        }

        // Assign label via DimensionType — try to set its built-in Label parameter
        private void AssignLabelViaDimensionType(Document doc, Dimension dim, Definition targetDef, ElementId targetId)
        {
            DimensionType dimType = dim.DimensionType as DimensionType;
            string info = $"DimensionType: {dimType?.Name ?? "null"}\n";
            info += $"Target param: {targetDef.Name}\n";
            info += $"TargetId: {targetId}\n\n";

            if (dimType == null)
            {
                TaskDialog.Show("DEBUG - DimensionType", info + "DimensionType là null!");
                return;
            }

            // List all params on DimensionType
            info += "Params trên DimensionType:\n";
            foreach (Parameter pp in dimType.Parameters)
            {
                info += $"  - '{pp.Definition.Name}' type={pp.Definition.GetDataType()} storage={pp.StorageType}\n";
            }

            // Try to find and set the built-in "Label" parameter on DimensionType
            Parameter labelParam = dimType.LookupParameter("Label");
            if (labelParam != null)
            {
                info += $"\nBuilt-in 'Label' param TÌM THẤY!\n";
                info += $"  Storage: {labelParam.StorageType}\n";
                info += $"  Definition type: {labelParam.Definition.GetType().Name}\n";
                info += $"  Definition spec: {labelParam.Definition.GetDataType()}\n";

                // Try Set(ElementId)
                try
                {
                    bool ok = labelParam.Set(targetId);
                    info += $"\nSet(ElementId) = {ok}";
                }
                catch (Exception ex)
                {
                    info += $"\nSet(ElementId) LỖI: {ex.Message}";
                }
            }
            else
            {
                info += "\nBuilt-in 'Label' param KHÔNG tìm thấy trên DimensionType";
            }

            // Try finding any param that matches targetDef.Name on dimType
            Parameter namedParam = dimType.LookupParameter(targetDef.Name);
            if (namedParam != null)
            {
                info += $"\nParam '{targetDef.Name}' trên dimType:\n";
                info += $"  Storage: {namedParam.StorageType}\n";
                try
                {
                    bool ok = namedParam.Set(targetId);
                    info += $"  Set(ElementId) = {ok}";
                }
                catch (Exception ex)
                {
                    info += $"  Set(ElementId) LỖI: {ex.Message}";
                }
            }

            TaskDialog.Show("DEBUG - DimensionType", info);
        }

        // Create a shared parameter in the SP file
        private bool CreateSharedParameter(Document doc, string name, ForgeTypeId specType)
        {
            dynamic spFile = null;
            try
            {
                string spPath = doc.Application.SharedParametersFilename;
                if (!string.IsNullOrEmpty(spPath) && System.IO.File.Exists(spPath))
                {
                    spFile = doc.Application.OpenSharedParameterFile();
                }
                else
                {
                    string projDir = string.IsNullOrEmpty(doc.PathName)
                        ? System.IO.Path.GetDirectoryName(System.Reflection.Assembly.GetExecutingAssembly().Location)
                        : System.IO.Path.GetDirectoryName(doc.PathName);
                    string spFilePath = System.IO.Path.Combine(projDir, "QParam_Labels.txt");

                    using (System.IO.StreamWriter sw = new System.IO.StreamWriter(spFilePath))
                    {
                        sw.WriteLine("# Revit Shared Parameter File");
                        sw.WriteLine("GROUP DIM Labels");
                        sw.WriteLine($"PARAM {name} LABEL TEXT 0 0 1");
                    }
                    doc.Application.SharedParametersFilename = spFilePath;
                    spFile = doc.Application.OpenSharedParameterFile();
                }
            }
            catch (Exception ex)
            {
                TaskDialog.Show("DEBUG - CreateSharedParameter", $"Mở/tạo SP file lỗi: {ex.Message}");
                return false;
            }

            if (spFile == null) return false;

            try
            {
                dynamic group = null;
                foreach (dynamic g in spFile.Groups)
                {
                    if (g.Name == "DIM Labels") { group = g; break; }
                }
                if (group == null)
                    group = spFile.Groups.Create("DIM Labels");

                // Check if already exists
                foreach (dynamic d in group.Definitions)
                {
                    if (d.Name == name) return true;
                }

                ExternalDefinitionCreationOptions opts = new ExternalDefinitionCreationOptions(name, specType);
                group.Definitions.Create(opts);
                return true;
            }
            catch (Exception ex)
            {
                TaskDialog.Show("DEBUG - CreateSharedParameter", $"Tạo param lỗi: {ex.Message}");
                return false;
            }
        }

        // Find a Definition by name in BindingMap, ParameterMap, or SP file
        // Also probe ExternalDefinition's actual available properties
        private Definition FindDefinitionInBindingMap(Document doc, string name, ForgeTypeId specType, out ElementId outId)
        {
            outId = null;

            // 1. Try BindingMap (Definition.Id → ElementId)
            BindingMap bindingMap = doc.ParameterBindings;
            DefinitionBindingMapIterator it = bindingMap.ForwardIterator();
            it.Reset();
            while (it.MoveNext())
            {
                Definition def = it.Key;
                if (def.Name.Equals(name, StringComparison.OrdinalIgnoreCase)
                    && def.GetDataType() == specType)
                {
                    try
                    {
                        dynamic d = def;
                        object idObj = d.Id;
                        if (idObj is ElementId eid)
                        {
                            outId = eid;
                            return def;
                        }
                    }
                    catch { }
                }
            }

            // 2. Try via ParameterMap on any dimension element
            FilteredElementCollector dimCol = new FilteredElementCollector(doc)
                .OfCategory(BuiltInCategory.OST_Dimensions)
                .WhereElementIsNotElementType();
            foreach (Element dimElem in dimCol)
            {
                foreach (Parameter p in dimElem.Parameters)
                {
                    Definition dp = p.Definition;
                    if (dp.Name.Equals(name, StringComparison.OrdinalIgnoreCase)
                        && dp.GetDataType() == specType)
                    {
                        outId = p.Id;
                        return dp;
                    }
                }
            }

            // 3. Try via DimensionType elements (type objects, not instances)
            FilteredElementCollector dtCol = new FilteredElementCollector(doc)
                .OfCategory(BuiltInCategory.OST_Dimensions)
                .WhereElementIsElementType();
            foreach (Element dtElem in dtCol)
            {
                foreach (Parameter p in dtElem.Parameters)
                {
                    Definition dp = p.Definition;
                    if (dp.Name.Equals(name, StringComparison.OrdinalIgnoreCase)
                        && dp.GetDataType() == specType)
                    {
                        outId = p.Id;
                        return dp;
                    }
                }
            }

            // 4. Try via SP file + probe ExternalDefinition properties
            try
            {
                dynamic spFile = doc.Application.OpenSharedParameterFile();
                if (spFile != null)
                {
                    foreach (dynamic grp in spFile.Groups)
                    {
                        foreach (dynamic d in grp.Definitions)
                        {
                            if (d.Name == name && d.GetDataType() == specType)
                            {
                                // Probe all available properties on this ExternalDefinition
                                string props = $"ExternalDefinition props:\n";
                                foreach (var prop in ((object)d).GetType().GetProperties())
                                {
                                    try { props += $"  {prop.Name}: {prop.GetValue(d)}\n"; }
                                    catch { props += $"  {prop.Name}: (err)\n"; }
                                }
                                foreach (var field in ((object)d).GetType().GetFields())
                                {
                                    try { props += $"  [field] {field.Name}: {field.GetValue(d)}\n"; }
                                    catch { props += $"  [field] {field.Name}: (err)\n"; }
                                }
                                TaskDialog.Show("DEBUG - ExternalDefinition props", props);
                                break;
                            }
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                TaskDialog.Show("DEBUG - SP file probe", $"Lỗi: {ex.Message}");
            }

            return null;
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
