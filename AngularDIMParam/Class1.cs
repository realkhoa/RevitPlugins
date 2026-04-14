using System;
using System.Collections.Generic;
using System.Linq;
using System.Reflection;

// Revit API
using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using Autodesk.Revit.UI.Selection;

namespace AngularDIMParam
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
                "AngularDIMParam",
                "Angular\nDIM\nParam",
                assemblyPath,
                "AngularDIMParam.Class1"
            );

            panel.AddItem(btnData);
            return Result.Succeeded;
        }

        public Result OnShutdown(UIControlledApplication app) => Result.Succeeded;
    }

    [Autodesk.Revit.Attributes.Transaction(Autodesk.Revit.Attributes.TransactionMode.Manual)]
    public class Class1 : IExternalCommand
    {
        public Result Execute(ExternalCommandData commandData, ref string message, ElementSet elements)
        {
            UIDocument uiDoc = commandData.Application.ActiveUIDocument;
            Document doc = uiDoc.Document;

            try
            {
                Reference r1 = uiDoc.Selection.PickObject(ObjectType.Edge, "Chọn cạnh 1");
                Reference r2 = uiDoc.Selection.PickObject(ObjectType.Edge, "Chọn cạnh 2");

                Line line1 = GetLineFromReference(doc, r1);
                Line line2 = GetLineFromReference(doc, r2);

                if (line1 == null || line2 == null)
                {
                    message = "Không lấy được line.";
                    return Result.Failed;
                }

                string selectedParamName = null;
                if (doc.IsFamilyDocument)
                {
                    var existingParams = doc.FamilyManager.Parameters
                        .Cast<FamilyParameter>()
                        .Select(p => p.Definition.Name)
                        .OrderBy(n => n)
                        .ToList();

                    selectedParamName = ShowParameterDialog(existingParams);

                    if (selectedParamName == null) return Result.Cancelled;
                }

                using (Transaction tx = new Transaction(doc, "Angular DIM 3D"))
                {
                    tx.Start();

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
                    if (angle < 1e-6) angle = Math.PI / 180;

                    XYZ xAxis = v1;
                    XYZ yAxis = normal.CrossProduct(xAxis).Normalize();
                    double radius = mid1.DistanceTo(mid2) * 0.5;
                    if (radius < 0.1) radius = 1.0;

                    Arc arc = Arc.Create(mid1, radius, 0, angle, xAxis, yAxis);
                    Autodesk.Revit.DB.View view = doc.ActiveView;

                    if (doc.IsFamilyDocument)
                    {
                        Dimension dim = doc.FamilyCreate.NewAngularDimension(view, arc, r1, r2);

                        if (!string.IsNullOrEmpty(selectedParamName))
                        {
                            var param = doc.FamilyManager.Parameters
                                .Cast<FamilyParameter>()
                                .FirstOrDefault(p => p.Definition.Name == selectedParamName);

                            if (param == null)
                            {
                                param = doc.FamilyManager.AddParameter(
                                    selectedParamName,
                                    GroupTypeId.Text,
                                    SpecTypeId.Angle,
                                    false);
                            }

                            if (param != null) dim.FamilyLabel = param;
                        }
                    }
                    else
                    {
                        DimensionType type = new FilteredElementCollector(doc)
                            .OfClass(typeof(DimensionType))
                            .Cast<DimensionType>()
                            .FirstOrDefault(x => x.StyleType == DimensionStyleType.Angular);

                        AngularDimension.Create(doc, view, arc, new List<Reference> { r1, r2 }, type);
                    }

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

        private string ShowParameterDialog(List<string> existingParams)
        {
            System.Windows.Forms.Form form = new System.Windows.Forms.Form
            {
                Text = "Thiết lập Parameter Góc",
                Width = 380,
                Height = 200,
                StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen,
                FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedDialog,
                MaximizeBox = false,
                MinimizeBox = false,
                TopMost = true
            };

            System.Windows.Forms.Label label = new System.Windows.Forms.Label
            {
                Text = "Chọn hoặc nhập tên Parameter:",
                Left = 15,
                Top = 20,
                Width = 330
            };

            System.Windows.Forms.ComboBox cmb = new System.Windows.Forms.ComboBox
            {
                Left = 15,
                Top = 50,
                Width = 330,
                DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDown
            };
            foreach (var p in existingParams) cmb.Items.Add(p);
            if (existingParams.Count > 0) cmb.SelectedIndex = 0;

            System.Windows.Forms.Button btnOk = new System.Windows.Forms.Button
            {
                Text = "OK",
                Left = 270,
                Top = 100,
                Width = 80,
                Height = 28,
                DialogResult = System.Windows.Forms.DialogResult.OK
            };
            btnOk.Click += (s, e) => { form.DialogResult = System.Windows.Forms.DialogResult.OK; };

            form.Controls.Add(label);
            form.Controls.Add(cmb);
            form.Controls.Add(btnOk);

            if (form.ShowDialog() == System.Windows.Forms.DialogResult.OK && !string.IsNullOrWhiteSpace(cmb.Text))
            {
                return cmb.Text.Trim();
            }
            return null;
        }

        private Line GetLineFromReference(Document doc, Reference r)
        {
            Element el = doc.GetElement(r);
            GeometryObject geoObj = el.GetGeometryObjectFromReference(r);
            if (geoObj is Edge edge)
            {
                var curve = edge.AsCurve();
                if (curve is Line l) return l;
                return Line.CreateBound(curve.GetEndPoint(0), curve.GetEndPoint(1));
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