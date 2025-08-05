using ATP_Common_Plugin.Utils;
using Autodesk.Revit.Attributes;
using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using Microsoft.Office.Interop.Excel;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Runtime.InteropServices;
using System.Windows.Forms;

namespace ATP_Common_Plugin.Commands
{
    [Transaction(TransactionMode.Manual)]
    [Regeneration(RegenerationOption.Manual)]
    public class MarkDictPart : IExternalCommand
    {
        public Result Execute(ExternalCommandData commandData, ref string message, ElementSet elements)
        {
            var uiApp = commandData.Application;
            var doc = uiApp.ActiveUIDocument.Document;

            // Показать форму выбора папки
            using (var form = new FolderSelectForm())
            {
                if (form.ShowDialog() != DialogResult.OK) return Result.Cancelled;

                string folderPath = form.SelectedFolder;

                foreach (string filePath in Directory.GetFiles(folderPath, "*.xlsx"))
                {
                    var dict = ReadExcelDictionary(filePath);
                    if (dict.Count == 0) continue;

                    using (Transaction t = new Transaction(doc, "Update elements from Excel"))
                    {
                        t.Start();

                        var allElements = new FilteredElementCollector(doc)
                            .WhereElementIsNotElementType()
                            .ToElements();

                        foreach (var element in allElements)
                        {
                            string index = GetAutoIndex(element);
                            if (index == null) continue;

                            if (dict.TryGetValue(index, out string comment))
                            {
                                Autodesk.Revit.DB.Parameter param = element.LookupParameter("АДСК_Маркировка_Комментарий");
                                if (param != null && !param.IsReadOnly)
                                    param.Set(comment);
                            }
                        }

                        t.Commit();
                    }
                }

                TaskDialog.Show("Готово", "Обработка файлов завершена.");
                return Result.Succeeded;
            }
        }

        private string GetAutoIndex(Element el)
        {
            string format = el.LookupParameter("MC Running index Format")?.AsString() ?? "";
            string group = el.LookupParameter("MC Running index Group")?.AsString() ?? "";
            string one = el.LookupParameter("MC Running index 1")?.AsString() ?? "";

            List<string> parts = new List<string>();
            if (!string.IsNullOrWhiteSpace(format)) parts.Add(format);
            if (!string.IsNullOrWhiteSpace(group)) parts.Add(group);
            if (!string.IsNullOrWhiteSpace(one)) parts.Add(one);

            return parts.Count > 0 ? string.Join("-", parts) : null;
        }

        private Dictionary<string, string> ReadExcelDictionary(string filePath)
        {
            var dict = new Dictionary<string, string>();
            Microsoft.Office.Interop.Excel.Application excelApp = null;
            Workbook workbook = null;

            try
            {
                excelApp = new Microsoft.Office.Interop.Excel.Application { Visible = false };
                workbook = excelApp.Workbooks.Open(filePath);
                Worksheet sheet = workbook.Worksheets["Title"] as Worksheet;
                if (sheet == null)
                    throw new Exception("Лист 'Title' не найден в файле: " + filePath);
                Range usedRange = sheet.UsedRange;

                int rows = usedRange.Rows.Count;
                int cols = usedRange.Columns.Count;

                int indexCol = -1, commentCol = -1;

                // Поиск нужных столбцов по заголовкам
                for (int col = 1; col <= cols; col++)
                {
                    string header = (usedRange.Cells[1, col] as Range)?.Text?.ToString();
                    if (header == "Автоиндекс") indexCol = col;
                    if (header == "АДСК_Маркировка_Комментарий") commentCol = col;
                }

                if (indexCol == -1 || commentCol == -1)
                    return dict;

                for (int row = 2; row <= rows; row++)
                {
                    string index = (usedRange.Cells[row, indexCol] as Range)?.Text?.ToString();
                    string comment = (usedRange.Cells[row, commentCol] as Range)?.Text?.ToString();

                    if (!string.IsNullOrWhiteSpace(index) && !dict.ContainsKey(index))
                        dict.Add(index, comment);
                }

                return dict;
            }
            catch (COMException ex)
            {
                TaskDialog.Show("Ошибка Excel", $"COM ошибка: {ex.Message}");
                return dict;
            }
            catch (Exception ex)
            {
                TaskDialog.Show("Ошибка", $"Ошибка чтения Excel: {ex.Message}");
                return dict;
            }
            finally
            {
                if (workbook != null)
                {
                    workbook.Close(false);
                    Marshal.ReleaseComObject(workbook);
                }
                if (excelApp != null)
                {
                    excelApp.Quit();
                    Marshal.ReleaseComObject(excelApp);
                }
            }
        }
    }
    public class FolderSelectForm : System.Windows.Forms.Form
    {
        public string SelectedFolder { get; private set; }

        public FolderSelectForm()
        {
            var folderDialog = new FolderBrowserDialog
            {
                Description = "Выберите папку с Excel файлами",
                ShowNewFolderButton = false
            };

            var result = folderDialog.ShowDialog();

            if (result == DialogResult.OK)
                SelectedFolder = folderDialog.SelectedPath;

            DialogResult = result;
        }
    }
}
