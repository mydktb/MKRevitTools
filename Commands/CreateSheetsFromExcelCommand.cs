using Autodesk.Revit.Attributes;
using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using MKRevitTools.Excel;
using System;
using System.Collections.Generic;
using System.Linq;
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
                // Show form to get Excel file path
                using (var form = new CreateSheetsFromExcelForm())
                {
                    if (form.ShowDialog() == DialogResult.OK)
                    {
                        string filePath = form.FilePath;

                        if (string.IsNullOrEmpty(filePath))
                        {
                            TaskDialog.Show("Error", "Please select an Excel file.");
                            return Result.Failed;
                        }

                        // Read data from Excel
                        var sheetDataList = ExcelReader.ReadSheetData(filePath);

                        if (sheetDataList == null || sheetDataList.Count == 0)
                        {
                            TaskDialog.Show("Error", "No data found in the Excel file.");
                            return Result.Failed;
                        }

                        int sheetsCreated = 0;

                        // Start transaction for creating sheets
                        using (Transaction trans = new Transaction(doc, "Create Sheets from Excel"))
                        {
                            trans.Start();

                            // Process each sheet from Excel
                            foreach (var sheetData in sheetDataList)
                            {
                                // Skip empty sheets
                                if (sheetData.Rows == null || sheetData.Rows.Count < 2) // Assuming row 1 is header
                                    continue;

                                // Process rows (skip header row if needed)
                                for (int i = 1; i < sheetData.Rows.Count; i++) // Start from 1 to skip header
                                {
                                    var row = sheetData.Rows[i];

                                    // Ensure we have at least the basic data for a sheet
                                    if (row.Count >= 2) // At least sheet number and name
                                    {
                                        string sheetNumber = row[0]?.Trim();
                                        string sheetName = row[1]?.Trim();

                                        if (!string.IsNullOrEmpty(sheetNumber) && !string.IsNullOrEmpty(sheetName))
                                        {
                                            // Create the sheet
                                            ViewSheet newSheet = CreateSheet(doc, sheetNumber, sheetName);

                                            if (newSheet != null)
                                            {
                                                sheetsCreated++;
                                                // Optional: Apply additional parameters from Excel columns
                                                ApplySheetParameters(newSheet, row);
                                            }
                                        }
                                    }
                                }
                            }

                            trans.Commit();
                        }

                        TaskDialog.Show("Success", $"Successfully created {sheetsCreated} Revit sheets from Excel data.");
                        return Result.Succeeded;
                    }
                    else
                    {
                        return Result.Cancelled;
                    }
                }
            }
            catch (Exception ex)
            {
                TaskDialog.Show("Error", $"Failed to create sheets from Excel: {ex.Message}");
                return Result.Failed;
            }
        }

        private ViewSheet CreateSheet(Document doc, string sheetNumber, string sheetName)
        {
            try
            {
                // Get a title block family symbol to use
                FilteredElementCollector collector = new FilteredElementCollector(doc);
                collector.OfClass(typeof(FamilySymbol));
                collector.OfCategory(BuiltInCategory.OST_TitleBlocks);

                FamilySymbol titleBlock = collector.FirstElement() as FamilySymbol;

                if (titleBlock == null)
                {
                    throw new Exception("No title block family found in the project.");
                }

                // Ensure the title block is active
                if (!titleBlock.IsActive)
                {
                    titleBlock.Activate();
                }

                // Create the sheet
                ViewSheet sheet = ViewSheet.Create(doc, titleBlock.Id);

                // Set sheet number and name
                sheet.SheetNumber = sheetNumber;
                sheet.Name = sheetName;

                return sheet;
            }
            catch (Exception ex)
            {
                TaskDialog.Show("Sheet Creation Error", $"Failed to create sheet '{sheetNumber} - {sheetName}': {ex.Message}");
                return null;
            }
        }

        private void ApplySheetParameters(ViewSheet sheet, List<string> rowData)
        {
            try
            {
                // Example: Apply additional parameters from Excel columns
                // Column 2 (index 2) and beyond can be used for custom parameters

                if (rowData.Count > 2 && !string.IsNullOrEmpty(rowData[2]))
                {
                    // Example: Set sheet description or other parameters
                    Parameter descParam = sheet.LookupParameter("Sheet Description");
                    if (descParam != null && !descParam.IsReadOnly)
                    {
                        descParam.Set(rowData[2]);
                    }
                }

                // Add more parameter mappings as needed based on your Excel structure
            }
            catch (Exception ex)
            {
                // Log parameter application errors but don't fail the entire process
                System.Diagnostics.Debug.WriteLine($"Error applying parameters to sheet: {ex.Message}");
            }
        }
    }
}