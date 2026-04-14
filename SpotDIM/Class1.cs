using Autodesk.Revit.UI;
using Autodesk.Revit.DB;
using Autodesk.Revit.UI.Selection;
using System;

namespace SpotDIM
{
    [Autodesk.Revit.Attributes.Transaction(Autodesk.Revit.Attributes.TransactionMode.Manual)]
    public class Command : IExternalCommand
    {
        public Result Execute(ExternalCommandData commandData, ref string message, ElementSet elements)
        {
            UIDocument uidoc = commandData.Application.ActiveUIDocument;
            Document doc = uidoc.Document;
            View view = doc.ActiveView;

            try
            {
                // 👉 Pick face
                Reference pickedRef = uidoc.Selection.PickObject(ObjectType.Face, "Pick a face");

                if (pickedRef == null) return Result.Cancelled;

                XYZ pickPoint = pickedRef.GlobalPoint;

                using (Transaction t = new Transaction(doc, "Create Spot Elevation"))
                {
                    t.Start();

                    // 👉 Offset vị trí text (để nó không đè lên điểm)
                    XYZ bend = pickPoint + new XYZ(1, 1, 0);
                    XYZ end = pickPoint + new XYZ(2, 2, 0);

                    SpotDimension spot = doc.Create.NewSpotElevation(
                        view,
                        pickedRef,
                        pickPoint, // origin
                        bend,      // bend point
                        end,       // end point
                        pickPoint, // ref point
                        true       // has leader
                    );

                    t.Commit();
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
    }
}