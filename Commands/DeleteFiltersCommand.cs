using Autodesk.Revit.Attributes;
using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using MKRevitTools.UI;
using System;

namespace MKRevitTools.Commands
{
    [Transaction(TransactionMode.Manual)]
    [Regeneration(RegenerationOption.Manual)]
    public class DeleteFiltersCommand : IExternalCommand
    {
        public Result Execute(ExternalCommandData commandData, ref string message, ElementSet elements)
        {
            try
            {
                UIApplication uiApp = commandData.Application;
                UIDocument uiDoc = uiApp.ActiveUIDocument;
                Document doc = uiDoc.Document;

                // Open the delete filters form
                var form = new DeleteFiltersForm(doc);
                form.ShowDialog();

                return Result.Succeeded;
            }
            catch (System.Exception ex)
            {
                message = ex.Message;
                TaskDialog.Show("Error", "Failed to open Delete Filters tool: " + ex.Message);
                return Result.Failed;
            }
        }
    }
}