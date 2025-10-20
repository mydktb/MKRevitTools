using System;
using System.Collections.Generic;
using System.IO;
using System.Windows.Forms;
using Autodesk.Revit.UI;

namespace MKRevitTools.Utils
{
    public class ExcelReader
    {
        public class SheetData
        {
            public string SheetNumber { get; set; }
            public string SheetName { get; set; }
            public string TitleBlock { get; set; }
        }

        public static List<SheetData> ReadSheetDataFromExcel()
        {
            List<SheetData> sheetDataList = new List<SheetData>();

            try
            {
                // Let user choose between Excel and CSV
                OpenFileDialog openFileDialog = new OpenFileDialog
                {
                    Filter = "Excel Files|*.xlsx;*.xls|CSV Files|*.csv|All Files|*.*",
                    Title = "Select Excel or CSV File with Sheet Data"
                };

                if (openFileDialog.ShowDialog() != DialogResult.OK)
                {
                    return sheetDataList; // User cancelled
                }

                string filePath = openFileDialog.FileName;
                string extension = Path.GetExtension(filePath).ToLower();

                if (extension == ".csv")
                {
                    sheetDataList = ReadFromCsv(filePath);
                }
                else
                {
                    sheetDataList = ReadFromExcel(filePath);
                }
            }
            catch (Exception ex)
            {
                TaskDialog.Show("File Read Error", $"Error reading file: {ex.Message}");
            }

            return sheetDataList;
        }

        private static List<SheetData> ReadFromCsv(string filePath)
        {
            var sheetDataList = new List<SheetData>();

            try
            {
                string[] lines = File.ReadAllLines(filePath);

                // Skip header row (if exists) and process data
                for (int i = 1; i < lines.Length; i++)
                {
                    string[] columns = lines[i].Split(',');

                    if (columns.Length >= 1 && !string.IsNullOrWhiteSpace(columns[0]))
                    {
                        sheetDataList.Add(new SheetData
                        {
                            SheetNumber = columns[0].Trim(),
                            SheetName = columns.Length > 1 ? columns[1].Trim() : "",
                            TitleBlock = columns.Length > 2 ? columns[2].Trim() : ""
                        });
                    }
                }
            }
            catch (Exception ex)
            {
                TaskDialog.Show("CSV Read Error", $"Error reading CSV file: {ex.Message}");
            }

            return sheetDataList;
        }

        private static List<SheetData> ReadFromExcel(string filePath)
        {
            // For now, just show a message to use CSV
            TaskDialog.Show("Excel Support",
                "For full Excel support, please install Microsoft Access Database Engine.\n\n" +
                "Alternatively, save your data as CSV file and try again.");

            return new List<SheetData>();
        }
    }
}