using Autodesk.Revit.Attributes;
using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using MKRevitTools.UI;
using System;
using System.Windows.Forms;

namespace MKRevitTools.Commands
{
    [Transaction(TransactionMode.Manual)]
    [Regeneration(RegenerationOption.Manual)]
    public class CreateSheetsFromExcelCommand : IExternalCommand
    {
        public Result Execute(ExternalCommandData commandData, ref string message, ElementSet elements)
        {
            UIApplication uiApp = commandData.Application;
            UIDocument uiDoc = uiApp.ActiveUIDocument;
            Document doc = uiDoc.Document;

            try
            {
                // Show the Excel Sheets Form
                var form = new CreateSheetsFromExcelForm();
                DialogResult result = form.ShowDialog();

                // Check if user clicked Create
                if (result != DialogResult.OK || !form.CreateSheets)
                {
                    return Result.Cancelled;
                }

                string excelPath = form.ExcelFilePath;
                string numberCol = form.SheetNumberColumn;
                string nameCol = form.SheetNameColumn;

                // Validate Excel file
                if (string.IsNullOrEmpty(excelPath) || !System.IO.File.Exists(excelPath))
                {
                    TaskDialog.Show("Error", "Please select a valid Excel file.");
                    return Result.Failed;
                }

                // Get the title block family symbol to use
                FamilySymbol titleBlock = GetTitleBlock(doc);
                if (titleBlock == null)
                {
                    TaskDialog.Show("Error", "No title block family found in the project.");
                    return Result.Failed;
                }

                // Read Excel data using your existing ExcelReader
                var excelReader = new Utils.ExcelReader();
                var sheetData = excelReader.ReadSheetData(excelPath, numberCol, nameCol);

                if (sheetData == null || sheetData.Count == 0)
                {
                    TaskDialog.Show("Error", "No sheet data found in the Excel file or invalid column names.");
                    return Result.Failed;
                }

                // Start transaction
                using (Transaction trans = new Transaction(doc, $"Create {sheetData.Count} Sheets from Excel"))
                {
                    trans.Start();

                    // Create sheets from Excel data
                    int createdCount = 0;
                    string sheetsList = "";

                    foreach (var data in sheetData)
                    {
                        try
                        {
                            ViewSheet sheet = ViewSheet.Create(doc, titleBlock.Id);
                            sheet.SheetNumber = data.SheetNumber;
                            sheet.Name = data.SheetName;
                            createdCount++;
                            sheetsList += $"Sheet: {sheet.SheetNumber} - {sheet.Name}\n";
                        }
                        catch (Exception ex)
                        {
                            // Log but continue with other sheets
                            TaskDialog.Show("Sheet Creation Warning",
                                $"Failed to create sheet {data.SheetNumber}: {ex.Message}");
                        }
                    }

                    trans.Commit();

                    // Show success message
                    TaskDialog.Show("Success",
                        $"Successfully created {createdCount} sheets from Excel:\n\n{sheetsList}");
                }

                return Result.Succeeded;
            }
            catch (Exception ex)
            {
                TaskDialog.Show("Error", "Failed to create sheets from Excel: " + ex.Message);
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
    }
}