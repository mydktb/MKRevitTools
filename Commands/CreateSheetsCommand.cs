using Autodesk.Revit.Attributes;
using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using MKRevitTools.UI;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Windows.Forms;

namespace MKRevitTools.Commands
{
    [Transaction(TransactionMode.Manual)]
    [Regeneration(RegenerationOption.Manual)]
    public class CreateSheetsCommand : IExternalCommand
    {
        public Result Execute(ExternalCommandData commandData, ref string message, ElementSet elements)
        {
            UIApplication uiApp = commandData.Application;
            UIDocument uiDoc = uiApp.ActiveUIDocument;
            Document doc = uiDoc.Document;

            try
            {
                // Show the Windows Form
                CreateSheetsForm dialog = new CreateSheetsForm();

                DialogResult result = dialog.ShowDialog();

                // Check if user clicked Create
                if (result != DialogResult.OK || !dialog.CreateSheets)
                {
                    return Result.Cancelled;
                }

                int sheetCount = dialog.SheetCount;
                int startNumber = dialog.StartNumber;
                string prefix = dialog.NumberPrefix;
                string suffix = dialog.NumberSuffix;

                // Get the title block family symbol to use
                FamilySymbol titleBlock = GetTitleBlock(doc);
                if (titleBlock == null)
                {
                    TaskDialog.Show("Error", "No title block family found in the project.");
                    return Result.Failed;
                }

                // Start transaction
                using (Transaction trans = new Transaction(doc, $"Create {sheetCount} Sheets"))
                {
                    trans.Start();

                    // Create sheets based on user input
                    List<ViewSheet> createdSheets = new List<ViewSheet>();

                    for (int i = 0; i < sheetCount; i++)
                    {
                        int currentSheetNumber = startNumber + i;
                        ViewSheet sheet = CreateSheet(doc, titleBlock, currentSheetNumber, prefix, suffix);
                        if (sheet != null)
                        {
                            createdSheets.Add(sheet);
                        }
                    }

                    trans.Commit();

                    // Show success message
                    string sheetsList = string.Join("\n", createdSheets.Select(s => $"Sheet: {s.SheetNumber} - {s.Name}"));
                    TaskDialog.Show("Success",
                        $"Successfully created {createdSheets.Count} sheets:\n\n{sheetsList}");
                }

                return Result.Succeeded;
            }
            catch (Exception ex)
            {
                TaskDialog.Show("Error", "Failed to create sheets: " + ex.Message);
                return Result.Failed;
            }
        }

        private FamilySymbol GetTitleBlock(Document doc)
        {
            // Get the first available title block family symbol
            FilteredElementCollector collector = new FilteredElementCollector(doc);
            collector.OfClass(typeof(FamilySymbol));
            collector.OfCategory(BuiltInCategory.OST_TitleBlocks);

            return collector.FirstElement() as FamilySymbol;
        }

        private ViewSheet CreateSheet(Document doc, FamilySymbol titleBlock, int sheetNumber, string prefix, string suffix)
        {
            try
            {
                // Create sheet with the title block
                ViewSheet sheet = ViewSheet.Create(doc, titleBlock.Id);

                // Set sheet number and name with prefix and suffix
                string sheetNum = $"{prefix}{sheetNumber:00}{suffix}";
                string sheetName = $"Sheet {sheetNumber}";

                // Set parameters
                sheet.SheetNumber = sheetNum;
                sheet.Name = sheetName;

                return sheet;
            }
            catch (Exception ex)
            {
                TaskDialog.Show("Sheet Creation Error", $"Failed to create sheet {sheetNumber}: {ex.Message}");
                return null;
            }
        }
    }
}