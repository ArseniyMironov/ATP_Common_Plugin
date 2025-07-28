using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Runtime.InteropServices;
using Excel = Microsoft.Office.Interop.Excel;

namespace ATP_Common_Plugin
{
    public static class ExcelUtils
    {
        public static Dictionary<string, string> ReadSystemDict(
            string filePath, 
            string worksheetName = "Title", 
            char keyValueSeparator = ':')
        {
            var result = new Dictionary<string, string>();

            Excel.Application excelApp = null;

            try
            {
                excelApp = new Excel.Application
                {
                    Visible = false,
                    DisplayAlerts = false,
                    AskToUpdateLinks = false
                };

                Excel.Workbook workbook = excelApp.Workbooks.Open(
                    filePath, 
                    UpdateLinks: Excel.XlUpdateLinks.xlUpdateLinksNever, 
                    ReadOnly: true);

                Excel.Worksheet worksheet = FindWorksheet(workbook, worksheetName);
                if (worksheet == null)
                    throw new System.Exception($"Лист '{worksheetName}' не найден");

                Excel.Range usedRange = worksheet.UsedRange;
                object[,] valueArray = usedRange.Value as object[,];

                if (valueArray != null)
                {
                    ProcessValueArray(valueArray, result, keyValueSeparator);
                }

                return result;
            }
            finally
            {
                CleanupResources(excelApp);
            }
        }

        private static Excel.Worksheet FindWorksheet(Excel.Workbook workbook, string name)
        {
            foreach (Excel.Worksheet sheet in workbook.Sheets)
            {
                if (sheet.Name.Equals(name))
                {
                    return sheet;
                }
            }
            return null;
        }


        private static void ProcessValueArray(object[,] valueArray, Dictionary<string, string> result, char separator)
        {
            int rows = valueArray.GetLength(0);
            int cols = valueArray.GetLength(1);

            for (int row = 1; row <= rows; row++)
            {
                for (int col = 1; col <= cols; col++)
                {
                    try
                    {
                        string cellValue = valueArray[row, col]?.ToString();
                        if (string.IsNullOrEmpty(cellValue)) continue;

                        var parts = cellValue.Split(separator);
                        if (parts.Length == 2)
                        {
                            string key = parts[0];
                            string value = parts[1];

                            if (!string.IsNullOrEmpty(key) && !result.ContainsKey(key))
                                result.Add(key, value);
                        }
                    }
                    catch (Exception ex)
                    {
                        Debug.WriteLine($"Ошибка в ячейке [{row},{col}]: {ex.Message}");
                    }
                }
            }
        }

        private static void CleanupResources(Excel.Application excelApp)
        {
            try
            {

            }
            catch (Exception ex)
            {
                Debug.WriteLine($"Ошибка при очистке ресурсов: {ex.Message}");
            }
            finally
            {
                Marshal.FinalReleaseComObject(excelApp);
                System.GC.Collect();
                System.GC.WaitForPendingFinalizers();
            }
        }
    }
}
