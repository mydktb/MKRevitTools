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

            // Main Dash Board Button
            var mainButtonData = new PushButtonData(
                "MainTools",
                "MK Dashboard",
                System.Reflection.Assembly.GetExecutingAssembly().Location,
                "MKRevitTools.Commands.MainToolsCommand"
            );
            panel.AddItem(mainButtonData);

            // #1 Delete Filters Button
            var deleteFiltersButtonData = new PushButtonData(
                "DeleteFilters",
                "Delete Filters",
                System.Reflection.Assembly.GetExecutingAssembly().Location,
                "MKRevitTools.Commands.DeleteFiltersCommand"
            );
            panel.AddItem(deleteFiltersButtonData);

            // #2 Add Filters to View Template Button
            var addFiltersToTemplateButtonData = new PushButtonData(
                "AddFiltersToTemplate",
                "Add Filters to Templates",
                System.Reflection.Assembly.GetExecutingAssembly().Location,
                "MKRevitTools.Commands.AddFiltersToViewTemplateCommand"
            );
            panel.AddItem(addFiltersToTemplateButtonData);

            // #3 Create Sheets from Excel Button
            var sheetsButtonData = new PushButtonData(
                "CreateSheetsFromExcel",
                "Create Sheets From Excel",
                System.Reflection.Assembly.GetExecutingAssembly().Location,
                "MKRevitTools.Commands.CreateSheetsFromExcelCommand"
            );
            panel.AddItem(sheetsButtonData);

            // #4 Create Multiple Sheets Button
            var createSheetsButtonData = new PushButtonData(
                "CreateMultipleSheets",
                "Create Multiple Sheets",
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