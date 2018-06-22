using System;
using System.Collections.Generic;
using System.IO;
using Excel = Microsoft.Office.Interop.Excel;

namespace ExcelInterop
{
    public class ExcelUtility
    {
        // Source sample - https://docs.microsoft.com/en-us/dotnet/csharp/programming-guide/interop/how-to-access-office-onterop-objects
        /// <summary>
        /// Saves the data into an excel sheet. Deletes an existing file.
        /// </summary>
        /// <param name="data">Data to be storedin the form of a 2 diimensional string array.</param>
        /// <param name="path">Path to be used to save the excel work sheet. Please use `.xlsx` extension to allow opening the file.</param>
        public static void SaveAsExcel(IList<IList<string>> data, string path)
        {
            if (File.Exists(path))
            {
                File.Delete(path);
            }

            var excelApp = new Excel.Application
            {
                // Dont need to see the app.
                Visible = false
            };

            // Create a new, empty workbook and add it to the collection returned 
            // by property Workbooks. The new workbook becomes the active workbook.
            // Add has an optional parameter for specifying a praticular template. 
            // Because no argument is sent in this example, Add creates a new workbook. 
            excelApp.Workbooks.Add();

            // This example uses a single workSheet. 
            Excel._Worksheet workSheet = (Excel.Worksheet)excelApp.ActiveSheet;

            var row = 0;
            var column = string.Empty;
            foreach (var rowData in data)
            {
                row++;
                column = string.Empty;
                foreach (var columnData in rowData)
                {
                    column = GetNextColumn(column);
                    workSheet.Cells[row, column] = columnData;
                }
            }

            workSheet.Range["A1", $"{column}1"].AutoFormat(
                Excel.XlRangeAutoFormat.xlRangeAutoFormatClassic2);

            workSheet.Range["A1", $"{column}1"].WrapText = true;
            workSheet.SaveAs(path);
            excelApp.ActiveWorkbook.Close();
            excelApp.Quit();

            Console.WriteLine($"Saving results file: {path}");
        }

        private static string GetNextColumn(string currentColumn)
        {
            if (string.IsNullOrEmpty(currentColumn))
            {
                return "A";
            }
            else
            {
                var lastChar = currentColumn[currentColumn.Length - 1];
                var previousChars = currentColumn.Length > 1 ? currentColumn.Substring(0, currentColumn.Length - 2) : string.Empty;

                return lastChar == 'Z' ?
                    $"{previousChars}AA" :
                    $"{previousChars}{char.ConvertFromUtf32(lastChar + 1)}";
            }
        }
    }
}
