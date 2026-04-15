using Autodesk.Revit.UI;
using System.Reflection;

namespace DIMAIO
{
    public class App : IExternalApplication
    {
        public Result OnShutdown(UIControlledApplication app) => Result.Succeeded;

        public Result OnStartup(UIControlledApplication application)
        {
            string assemblyPath = Assembly.GetExecutingAssembly().Location;
            string tabName = "Prima";
            string sectionName = "Dimension";

            try { application.CreateRibbonTab(tabName); } catch { }

            RibbonPanel panel = application.GetRibbonPanels(tabName)
                .FirstOrDefault(p => p.Name == sectionName)
                ?? application.CreateRibbonPanel(tabName, sectionName);

            // Linear DIM
            PushButtonData LinearDimBtnData = new PushButtonData(
                "Linear DIM",
                "Linear\nDIM",
                assemblyPath,
                "DIMAIO.LinearDIMCommand"
            );

            if (!panel.GetItems().OfType<PushButton>().Any(b => b.Name == "Linear DIM"))
            {
                PushButton button = panel.AddItem(LinearDimBtnData) as PushButton;
                button.ToolTip = "Places horizontal or vertical dimensions that measure the\r\ndistance between reference points.\r\n\r\nThe dimensions are aligned with the horizontal or vertical axis\r\nof the view.";
            }

            // Angular DIM
            PushButtonData AngularDimBtnData = new PushButtonData(
                "Angular DIM",
                "Angular\nDIM",
                assemblyPath,
                "DIMAIO.AngularDIMCommand"
            );
            if (!panel.GetItems().OfType<PushButton>().Any(b => b.Name == "Angular DIM"))
            {
                PushButton button = panel.AddItem(AngularDimBtnData) as PushButton;
                button.ToolTip = "Places a dimension that measures the angle between reference points sharing a common intersection.";
            }

            // Aligned DIM
            PushButtonData AlignedDimBtnData = new PushButtonData(
                "Aligned DIM",
                "Aligned\nDIM",
                assemblyPath,
                "DIMAIO.AlignedDIMCommand"
            );
            if (!panel.GetItems().OfType<PushButton>().Any(b => b.Name == "Aligned DIM"))
            {
                PushButton button = panel.AddItem(AlignedDimBtnData) as PushButton;
                button.ToolTip = "Places dimensions between parallel references, or between\r\nmultiple points.";
            }

            // Spot Coordinate
            PushButtonData SpotCoordinateBtnData = new PushButtonData(
                "Spot Coordinate",
                "Spot\nCoordinate",
                assemblyPath,
                "DIMAIO.SpotCoordinateDIMCommand"
            );
            if (!panel.GetItems().OfType<PushButton>().Any(b => b.Name == "Spot Coordinate"))
            {
                PushButton button = panel.AddItem(SpotCoordinateBtnData) as PushButton;
                button.ToolTip = "Displays the North/South and East/West coordinates of points\r\nin a project.\r\n\r\nYou can place spot coordinates on floors, walls, toposurfaces,\r\nand boundary lines. You can also place spot coordinates on\r\nnon-horizontal surfaces and non-planar edges.";
            }

            // Spot DIM
            PushButtonData SpotDimBtnData = new PushButtonData(
                "Spot DIM",
                "Spot\nDIM",
                assemblyPath,
                "DIMAIO.SpotDIMCommand"
            );
            if (!panel.GetItems().OfType<PushButton>().Any(b => b.Name == "Spot DIM"))
            {
                PushButton button = panel.AddItem(SpotDimBtnData) as PushButton;
                button.ToolTip = "Places a spot elevation that measures the vertical distance\r\nfrom a reference plane to a point.\r\n\r\nYou can place spot elevations on floors, walls, toposurfaces,\r\nand boundary lines. You can also place spot elevations on\r\nnon-horizontal surfaces and non-planar edges.";
            }

            // Radial DIM
            PushButtonData RadialDimBtnData = new PushButtonData(
                "Radial DIM",
                "Radial\nDIM",
                assemblyPath,
                "DIMAIO.RadialDIMCommand"
            );
            if (!panel.GetItems().OfType<PushButton>().Any(b => b.Name == "Radial DIM"))
            {
                PushButton button = panel.AddItem(RadialDimBtnData) as PushButton;
                button.ToolTip = "Places a dimension that measures the radius of an arc or circle.";
            }

            // Diameter DIM
            PushButtonData DiameterDimBtnData = new PushButtonData(
                "Diameter DIM",
                "Diameter\nDIM",
                assemblyPath,
                "DIMAIO.DiameterDIMCommand"
            );
            if (!panel.GetItems().OfType<PushButton>().Any(b => b.Name == "Diameter DIM"))
            {
                PushButton button = panel.AddItem(DiameterDimBtnData) as PushButton;
                button.ToolTip = "Places a dimension that measures the diameter of a circle or arc.";
            }

            return Result.Succeeded;
        }
    }
}
