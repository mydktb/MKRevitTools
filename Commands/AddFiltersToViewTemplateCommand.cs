using Autodesk.Revit.Attributes;
using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using MKRevitTools.UI;

namespace MKRevitTools.Commands
{
    [Transaction(TransactionMode.Manual)]
    [Regeneration(RegenerationOption.Manual)]
    public class AddFiltersToViewTemplateCommand : IExternalCommand
    {
        public Result Execute(ExternalCommandData commandData, ref string message, ElementSet elements)
        {
            try
            {
                UIApplication uiApp = commandData.Application;
                UIDocument uiDoc = uiApp.ActiveUIDocument;
                Document doc = uiDoc.Document;

                // Open the add filters to view template form
                var form = new AddFiltersToViewTemplateForm(doc);
                form.ShowDialog();

                return Result.Succeeded;
            }
            catch (System.Exception ex)
            {
                message = ex.Message;
                TaskDialog.Show("Error", "Failed to open Add Filters to View Template tool: " + ex.Message);
                return Result.Failed;
            }
        }
    }
}