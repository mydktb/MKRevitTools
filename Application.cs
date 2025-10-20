using Autodesk.Revit.UI;

namespace MKRevitTools
{
    public class Application : IExternalApplication
    {
        public Result OnStartup(UIControlledApplication application)
        {
            // Create ribbon tab
            application.CreateRibbonTab("MK Tools");

            // Create ribbon panel
            var panel = application.CreateRibbonPanel("MK Tools", "MK Revit Tools");

            // NEW: Main Dashboard button - SIMPLIFIED VERSION
            var mainButtonData = new PushButtonData(
                "MainTools",
                "MK Dashboard",
                System.Reflection.Assembly.GetExecutingAssembly().Location,  // FIX: Use this method
                "MKRevitTools.Commands.MainToolsCommand"  // FIX: Use string namespace
            );
            panel.AddItem(mainButtonData);

            // EXISTING: Your direct sheet creation buttons - ALSO SIMPLIFIED
            var sheetsButtonData = new PushButtonData(
                "CreateSheets",
                "Create Sheets",
                System.Reflection.Assembly.GetExecutingAssembly().Location,
                "MKRevitTools.Commands.CreateSheetsFromExcelCommand"
            );
            panel.AddItem(sheetsButtonData);

            var createSheetsButtonData = new PushButtonData(
                "CreateSheetsBasic",
                "Create Sheets Basic",
                System.Reflection.Assembly.GetExecutingAssembly().Location,
                "MKRevitTools.Commands.CreateSheetsCommand"
            );
            panel.AddItem(createSheetsButtonData);

            return Result.Succeeded;
        }

        public Result OnShutdown(UIControlledApplication application)
        {
            return Result.Succeeded;
        }
    }
}