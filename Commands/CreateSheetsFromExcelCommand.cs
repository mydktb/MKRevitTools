using System;
using System.Collections.Generic;
using System.Linq;
using Autodesk.Revit.UI;
using Autodesk.Revit.DB;
using Autodesk.Revit.Attributes;
using MKRevitTools.Utils;

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
                // Read sheet data from Excel
                List<ExcelReader.SheetData> sheetDataList = ExcelReader.ReadSheetDataFromExcel();

                if (sheetDataList.Count == 0)
                {
                    TaskDialog.Show("Information", "No sheet data found in Excel file or user cancelled.");
                    return Result.Cancelled;
                }

                // Show confirmation dialog
                string confirmationMessage = $"Found {sheetDataList.Count} sheets in Excel file:\n\n";
                confirmationMessage += string.Join("\n", sheetDataList.Take(5).Select(s => $"{s.SheetNumber} - {s.SheetName}"));

                if (sheetDataList.Count > 5)
                {
                    confirmationMessage += $"\n... and {sheetDataList.Count - 5} more";
                }

                confirmationMessage += "\n\nDo you want to create these sheets?";

                TaskDialogResult confirmation = TaskDialog.Show("Confirm Sheet Creation", confirmationMessage,
                    TaskDialogCommonButtons.Yes | TaskDialogCommonButtons.No);

                if (confirmation != TaskDialogResult.Yes)
                {
                    return Result.Cancelled;
                }

                // Get available title blocks
                List<FamilySymbol> titleBlocks = GetTitleBlocks(doc);
                if (titleBlocks.Count == 0)
                {
                    TaskDialog.Show("Error", "No title block families found in the project.");
                    return Result.Failed;
                }

                // Start transaction
                using (Transaction trans = new Transaction(doc, "Create Sheets from Excel"))
                {
                    trans.Start();

                    List<ViewSheet> createdSheets = new List<ViewSheet>();
                    List<string> errors = new List<string>();

                    foreach (var sheetData in sheetDataList)
                    {
                        try
                        {
                            // Select title block (use default if not specified in Excel)
                            FamilySymbol titleBlock = GetTitleBlockForSheet(sheetData, titleBlocks);

                            if (titleBlock == null)
                            {
                                errors.Add($"No suitable title block found for sheet: {sheetData.SheetNumber}");
                                continue;
                            }

                            ViewSheet sheet = CreateSheetFromData(doc, titleBlock, sheetData);
                            if (sheet != null)
                            {
                                createdSheets.Add(sheet);
                            }
                        }
                        catch (Exception ex)
                        {
                            errors.Add($"Failed to create sheet {sheetData.SheetNumber}: {ex.Message}");
                        }
                    }

                    trans.Commit();

                    // Show results
                    string resultMessage = $"Successfully created {createdSheets.Count} out of {sheetDataList.Count} sheets.";

                    if (errors.Count > 0)
                    {
                        resultMessage += $"\n\nErrors:\n{string.Join("\n", errors.Take(10))}";
                        if (errors.Count > 10)
                        {
                            resultMessage += $"\n... and {errors.Count - 10} more errors";
                        }
                    }

                    TaskDialog.Show("Results", resultMessage);
                }

                return Result.Succeeded;
            }
            catch (Exception ex)
            {
                TaskDialog.Show("Error", $"Failed to create sheets from Excel: {ex.Message}");
                return Result.Failed;
            }
        }

        private List<FamilySymbol> GetTitleBlocks(Document doc)
        {
            FilteredElementCollector collector = new FilteredElementCollector(doc);
            collector.OfClass(typeof(FamilySymbol));
            collector.OfCategory(BuiltInCategory.OST_TitleBlocks);

            return collector.Cast<FamilySymbol>().ToList();
        }

        private FamilySymbol GetTitleBlockForSheet(ExcelReader.SheetData sheetData, List<FamilySymbol> titleBlocks)
        {
            // If title block is specified in Excel, try to find it
            if (!string.IsNullOrEmpty(sheetData.TitleBlock))
            {
                var specifiedTitleBlock = titleBlocks.FirstOrDefault(tb =>
                    tb.Name.Equals(sheetData.TitleBlock, StringComparison.OrdinalIgnoreCase));

                if (specifiedTitleBlock != null)
                {
                    return specifiedTitleBlock;
                }
            }

            // Otherwise, return the first available title block
            return titleBlocks.FirstOrDefault();
        }

        private ViewSheet CreateSheetFromData(Document doc, FamilySymbol titleBlock, ExcelReader.SheetData sheetData)
        {
            try
            {
                // Create sheet with the title block
                ViewSheet sheet = ViewSheet.Create(doc, titleBlock.Id);

                // Set sheet number and name from Excel data
                sheet.SheetNumber = sheetData.SheetNumber;
                sheet.Name = string.IsNullOrEmpty(sheetData.SheetName) ? $"Sheet {sheetData.SheetNumber}" : sheetData.SheetName;

                return sheet;
            }
            catch (Exception ex)
            {
                throw new Exception($"Failed to create sheet {sheetData.SheetNumber}: {ex.Message}");
            }
        }
    }
}