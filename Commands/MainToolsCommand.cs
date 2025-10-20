using Autodesk.Revit.Attributes;
using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using MKRevitTools.UI;
using System.Windows.Forms;

namespace MKRevitTools.Commands
{
    [Transaction(TransactionMode.Manual)]
    [Regeneration(RegenerationOption.Manual)]
    public class MainToolsCommand : IExternalCommand
    {
        public Result Execute(ExternalCommandData commandData, ref string message, ElementSet elements)
        {
            try
            {
                // This opens the NEW main menu form
                using (var form = new ModernMainForm())
                {
                    form.ShowDialog();
                }
                return Result.Succeeded;
            }
            catch (System.Exception ex)
            {
                message = ex.Message;
                TaskDialog.Show("Error", "Failed to open MK Tools Dashboard: " + ex.Message);
                return Result.Failed;
            }
        }
    }
}