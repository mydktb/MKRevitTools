using System;
using System.Collections.Generic;
using System.IO;
using OfficeOpenXml;

namespace MKRevitTools.Excel
{
    public class ExcelReader
    {
        public static List<SheetData> ReadSheetData(string filePath)
        {
            var sheetDataList = new List<SheetData>();

            // Set EPPlus license context
            ExcelPackage.LicenseContext = LicenseContext.NonCommercial;

            try
            {
                using (var package = new ExcelPackage(new FileInfo(filePath)))
                {
                    foreach (var worksheet in package.Workbook.Worksheets)
                    {
                        // Skip hidden worksheets
                        if (worksheet.Hidden != eWorkSheetHidden.Visible)
                            continue;

                        var sheetData = new SheetData
                        {
                            SheetName = worksheet.Name,
                            Rows = new List<List<string>>()
                        };

                        // Check if worksheet has data
                        if (worksheet.Dimension != null)
                        {
                            int rowCount = worksheet.Dimension.Rows;
                            int colCount = worksheet.Dimension.Columns;

                            for (int row = 1; row <= rowCount; row++)
                            {
                                var currentRow = new List<string>();
                                for (int col = 1; col <= colCount; col++)
                                {
                                    var cell = worksheet.Cells[row, col];
                                    string value = cell.Text?.Trim() ?? string.Empty;
                                    currentRow.Add(value);
                                }
                                sheetData.Rows.Add(currentRow);
                            }
                        }

                        sheetDataList.Add(sheetData);
                    }
                }
            }
            catch (Exception ex)
            {
                throw new Exception($"Error reading Excel file: {ex.Message}", ex);
            }

            return sheetDataList;
        }
    }

    public class SheetData
    {
        public string SheetName { get; set; }
        public List<List<string>> Rows { get; set; }
    }
}