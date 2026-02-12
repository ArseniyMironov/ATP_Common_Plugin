using System;
using System.Collections.Generic;
using System.Linq;
using System.Runtime.InteropServices;
using System.Windows.Forms;
using Autodesk.Revit.Attributes;
using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using Application = Autodesk.Revit.ApplicationServices.Application;
using Excel = Microsoft.Office.Interop.Excel;
using Form = System.Windows.Forms.Form;
using Point = System.Drawing.Point;
using Size = System.Drawing.Size;
using TextBox = System.Windows.Forms.TextBox;
using ComboBox = System.Windows.Forms.ComboBox;
using Button = System.Windows.Forms.Button;
using Label = System.Windows.Forms.Label;
using CheckBox = System.Windows.Forms.CheckBox;
using OpenFileDialog = System.Windows.Forms.OpenFileDialog;

namespace ATP_Common_Plugin.Commands
{
    [Transaction(TransactionMode.Manual)]
    public class ImportSpaceParamsCommand : IExternalCommand
    {
        public Result Execute(ExternalCommandData commandData, ref string message, ElementSet elements)
        {
            UIDocument uidoc = commandData.Application.ActiveUIDocument;
            Document doc = uidoc.Document;

            // 1. Сбор параметров Space
            List<string> availableParams = GetWritableSpaceParameters(doc);

            // 2. Запуск формы
            using (var form = new ExcelImportForm(availableParams))
            {
                var result = form.ShowDialog();
                if (result != DialogResult.OK || !form.IsRun)
                {
                    return Result.Cancelled;
                }

                // 3. Сбор всех Space
                var spacesDict = new Dictionary<string, SpatialElement>();
                var collector = new FilteredElementCollector(doc)
                    .OfCategory(BuiltInCategory.OST_MEPSpaces)
                    .WhereElementIsNotElementType();

                foreach (SpatialElement space in collector)
                {
                    Parameter pNum = space.get_Parameter(BuiltInParameter.ROOM_NUMBER);
                    if (pNum != null && pNum.HasValue)
                    {
                        string cleanNum = pNum.AsString()?.Trim();
                        if (!string.IsNullOrEmpty(cleanNum) && !spacesDict.ContainsKey(cleanNum))
                        {
                            spacesDict.Add(cleanNum, space);
                        }
                    }
                }

                // 4. Чтение Excel
                List<ExcelDataRow> excelData;
                try
                {
                    excelData = ReadExcelData(form.ExcelPath, form.SheetName, form.IdxRoom, form.IdxLoad, form.IdxTemp);
                }
                catch (Exception ex)
                {
                    TaskDialog.Show("Ошибка Excel", ex.Message);
                    return Result.Failed;
                }

                // 5. Транзакция
                int updatedCount = 0;
                List<string> errors = new List<string>();

                using (Transaction t = new Transaction(doc, "Импорт параметров Space"))
                {
                    t.Start();

                    foreach (var item in excelData)
                    {
                        if (spacesDict.ContainsKey(item.RoomNumber))
                        {
                            SpatialElement space = spacesDict[item.RoomNumber];

                            // --- Нагрузка ---
                            if (form.DoLoad && item.Load != null)
                            {
                                // Передаем doc для конвертации единиц
                                bool ok = SetParameterValue(space, form.ParamLoadName, item.Load.Value, false, doc, out string msg);
                                if (!ok && errors.Count < 10) errors.Add($"Load Err ({item.RoomNumber}): {msg}");
                            }

                            // --- Температура ---
                            if (form.DoTemp && item.Temp != null)
                            {
                                bool ok = SetParameterValue(space, form.ParamTempName, item.Temp.Value, true, doc, out string msg);
                                if (ok)
                                {
                                    updatedCount++;
                                }
                                else if (errors.Count < 10)
                                {
                                    errors.Add($"Temp Err ({item.RoomNumber}): {msg}");
                                }
                            }
                        }
                    }

                    t.Commit();
                }

                // 6. Отчет
                string report = $"Готово!\nОбновлено (Temp): {updatedCount}\nНайдено строк в Excel: {excelData.Count}";
                if (errors.Count > 0)
                {
                    report += "\n\nОшибки (первые 10):\n" + string.Join("\n", errors);
                }

                TaskDialog.Show("Результат", report);

                return Result.Succeeded;
            }
        }

        // ================= ВСПОМОГАТЕЛЬНЫЕ МЕТОДЫ =================

        private List<string> GetWritableSpaceParameters(Document doc)
        {
            var paramsList = new List<string>();
            var collector = new FilteredElementCollector(doc)
                .OfCategory(BuiltInCategory.OST_MEPSpaces)
                .WhereElementIsNotElementType();

            Element firstSpace = collector.FirstElement();
            if (firstSpace != null)
            {
                foreach (Parameter p in firstSpace.Parameters)
                {
                    if (!p.IsReadOnly)
                    {
                        paramsList.Add(p.Definition.Name);
                    }
                }
                paramsList.Sort();
            }
            return paramsList;
        }

        // --- ГЛАВНОЕ ИЗМЕНЕНИЕ ЗДЕСЬ ---
        private bool SetParameterValue(Element elem, string paramName, double value, bool isTemperature, Document doc, out string message)
        {
            message = "OK";
            Parameter param = elem.LookupParameter(paramName);

            if (param == null && paramName == "Design Heating Load")
                param = elem.get_Parameter(BuiltInParameter.ROOM_DESIGN_HEATING_LOAD_PARAM);

            if (param == null)
            {
                message = $"Параметр '{paramName}' не найден";
                return false;
            }
            if (param.IsReadOnly)
            {
                message = $"Параметр '{paramName}' только для чтения";
                return false;
            }

            try
            {
                if (param.StorageType == StorageType.Double)
                {
                    double valToSet = value;

                    // 1. Попытка использовать UnitUtils (Умная конвертация)
                    // Получаем тип данных параметра (SpecTypeId)
                    ForgeTypeId specTypeId = param.Definition.GetDataType();

                    // Проверяем, является ли параметр измеряемой величиной (например, Мощность, Температура)
                    if (UnitUtils.IsMeasurableSpec(specTypeId))
                    {
                        // Получаем текущие настройки проекта для этого типа данных
                        FormatOptions fo = doc.GetUnits().GetFormatOptions(specTypeId);
                        ForgeTypeId currentDisplayUnit = fo.GetUnitTypeId();

                        // Конвертируем значение "как видит пользователь" во "внутренние единицы"
                        // Если проект в Ваттах: 1820 -> 6209.8 (BTU/h)
                        // Если проект в Цельсиях: 20 -> 293.15 (K)
                        valToSet = UnitUtils.ConvertToInternalUnits(value, currentDisplayUnit);
                    }
                    else
                    {
                        // 2. Если параметр просто "Число" (без единиц), но мы знаем, что это Температура (по флагу)
                        // Это Fallback для пользовательских параметров с типом "Number"
                        if (isTemperature)
                        {
                            valToSet = value + 273.15;
                        }
                    }

                    return param.Set(valToSet);
                }
                else if (param.StorageType == StorageType.String)
                {
                    return param.Set(value.ToString());
                }
                else if (param.StorageType == StorageType.Integer)
                {
                    return param.Set((int)value);
                }

                message = "Неверный тип данных параметра";
                return false;
            }
            catch (Exception ex)
            {
                message = ex.Message;
                return false;
            }
        }

        private List<ExcelDataRow> ReadExcelData(string path, string sheetName, int colRoom, int colLoad, int colTemp)
        {
            var result = new List<ExcelDataRow>();

            Excel.Application xlApp = null;
            Excel.Workbook xlWorkbook = null;
            Excel.Worksheet xlWorksheet = null;
            Excel.Range xlRange = null;

            try
            {
                xlApp = new Excel.Application();
                xlApp.Visible = false;

                xlWorkbook = xlApp.Workbooks.Open(path);
                try
                {
                    xlWorksheet = (Excel.Worksheet)xlWorkbook.Sheets[sheetName];
                }
                catch
                {
                    throw new Exception($"Лист '{sheetName}' не найден.");
                }

                xlRange = xlWorksheet.UsedRange;
                object[,] valueArray = (object[,])xlRange.Value2;

                int rowsCount = valueArray.GetLength(0);

                for (int r = 1; r <= rowsCount; r++)
                {
                    try
                    {
                        object rawRoom = valueArray[r, colRoom];
                        string roomVal = CleanString(rawRoom);

                        if (string.IsNullOrEmpty(roomVal)) continue;

                        double? valLoad = null;
                        double? valTemp = null;

                        object rawLoad = valueArray[r, colLoad];
                        if (rawLoad != null && double.TryParse(rawLoad.ToString(), out double parsedLoad))
                            valLoad = parsedLoad;

                        object rawTemp = valueArray[r, colTemp];
                        if (rawTemp != null && double.TryParse(rawTemp.ToString(), out double parsedTemp))
                            valTemp = parsedTemp;

                        result.Add(new ExcelDataRow
                        {
                            RoomNumber = roomVal,
                            Load = valLoad,
                            Temp = valTemp
                        });
                    }
                    catch { continue; }
                }
            }
            finally
            {
                if (xlRange != null) Marshal.ReleaseComObject(xlRange);
                if (xlWorksheet != null) Marshal.ReleaseComObject(xlWorksheet);
                if (xlWorkbook != null)
                {
                    xlWorkbook.Close(false);
                    Marshal.ReleaseComObject(xlWorkbook);
                }
                if (xlApp != null)
                {
                    xlApp.Quit();
                    Marshal.ReleaseComObject(xlApp);
                }
            }
            return result;
        }

        private string CleanString(object val)
        {
            if (val == null) return null;
            string s = val.ToString().Trim();
            if (s.EndsWith(".0")) s = s.Substring(0, s.Length - 2);
            return s.Trim();
        }
    }

    public class ExcelDataRow
    {
        public string RoomNumber { get; set; }
        public double? Load { get; set; }
        public double? Temp { get; set; }
    }

    // ================= КЛАСС ФОРМЫ (БЕЗ ИЗМЕНЕНИЙ, но включен для полноты) =================
    public class ExcelImportForm : Form
    {
        public string ExcelPath => txtPath.Text;
        public string SheetName => txtSheet.Text;
        public bool IsRun { get; private set; } = false;
        public bool DoLoad => chkLoad.Checked;
        public bool DoTemp => chkTemp.Checked;
        public string ParamLoadName => cmbLoad.Text;
        public string ParamTempName => cmbTemp.Text;
        public int IdxRoom { get; private set; }
        public int IdxLoad { get; private set; }
        public int IdxTemp { get; private set; }

        private TextBox txtPath;
        private TextBox txtSheet;
        private TextBox txtRoomCol;
        private TextBox txtLoadCol;
        private TextBox txtTempCol;
        private ComboBox cmbLoad;
        private ComboBox cmbTemp;
        private CheckBox chkLoad;
        private CheckBox chkTemp;
        private List<string> _paramList;

        public ExcelImportForm(List<string> paramList)
        {
            _paramList = paramList;
            InitializeComponent();
        }

        private void InitializeComponent()
        {
            this.Text = "Импорт Excel -> Revit (Smart Units)";
            this.Size = new Size(520, 520);
            this.FormBorderStyle = FormBorderStyle.FixedDialog;
            this.StartPosition = FormStartPosition.CenterScreen;
            this.TopMost = true;

            int startY = 20;
            int gap = 70;

            // 1. Файл
            var lblFile = new Label { Text = "Путь к файлу Excel:", Location = new Point(20, startY), Size = new Size(400, 20) };
            this.Controls.Add(lblFile);
            txtPath = new TextBox { Location = new Point(20, startY + 20), Size = new Size(350, 20) };
            this.Controls.Add(txtPath);
            var btnBrowse = new Button { Text = "...", Location = new Point(380, startY + 19), Size = new Size(80, 22) };
            btnBrowse.Click += BtnBrowse_Click;
            this.Controls.Add(btnBrowse);

            // 2. Лист
            int currentY = startY + 50;
            var lblSheet = new Label { Text = "Имя листа Excel:", Location = new Point(20, currentY), AutoSize = true };
            this.Controls.Add(lblSheet);
            txtSheet = new TextBox { Text = "Теплопотери К1_L1", Location = new Point(150, currentY - 3), Size = new Size(310, 20) };
            this.Controls.Add(txtSheet);

            currentY += 40;

            // 3. Параметры
            var lblRoom = new Label { Text = "1. Номер помещения (Колонка Excel):", Location = new Point(20, currentY), AutoSize = true };
            this.Controls.Add(lblRoom);
            txtRoomCol = new TextBox { Text = "A", Location = new Point(360, currentY), Size = new Size(100, 20) };
            this.Controls.Add(txtRoomCol);

            currentY += 40;

            chkLoad = new CheckBox { Text = "2. Обновлять Нагрузку (Вт)", Location = new Point(20, currentY), Size = new Size(300, 20), Checked = true };
            chkLoad.CheckedChanged += (s, e) => { cmbLoad.Enabled = txtLoadCol.Enabled = chkLoad.Checked; };
            this.Controls.Add(chkLoad);

            currentY += 25;
            cmbLoad = CreateComboBox("Design Heating Load", 20, currentY, 320);
            var lblLoadCol = new Label { Text = "Колонка:", Location = new Point(360, currentY - 15), AutoSize = true };
            this.Controls.Add(lblLoadCol);
            txtLoadCol = new TextBox { Text = "X", Location = new Point(360, currentY), Size = new Size(100, 20) };
            this.Controls.Add(txtLoadCol);

            currentY += gap;

            chkTemp = new CheckBox { Text = "3. Обновлять Температуру (C)", Location = new Point(20, currentY), Size = new Size(300, 20), Checked = true };
            chkTemp.CheckedChanged += (s, e) => { cmbTemp.Enabled = txtTempCol.Enabled = chkTemp.Checked; };
            this.Controls.Add(chkTemp);

            currentY += 25;
            cmbTemp = CreateComboBox("ADSK_Температура в помещении", 20, currentY, 320);
            var lblTempCol = new Label { Text = "Колонка:", Location = new Point(360, currentY - 15), AutoSize = true };
            this.Controls.Add(lblTempCol);
            txtTempCol = new TextBox { Text = "I", Location = new Point(360, currentY), Size = new Size(100, 20) };
            this.Controls.Add(txtTempCol);

            var btnRun = new Button { Text = "Запустить импорт", Location = new Point(150, currentY + 60), Size = new Size(200, 40) };
            btnRun.Click += BtnRun_Click;
            this.Controls.Add(btnRun);
        }

        private ComboBox CreateComboBox(string defaultVal, int x, int y, int width)
        {
            var cmb = new ComboBox();
            cmb.Location = new Point(x, y);
            cmb.Size = new Size(width, 21);
            cmb.AutoCompleteMode = AutoCompleteMode.SuggestAppend;
            cmb.AutoCompleteSource = AutoCompleteSource.ListItems;
            if (_paramList != null) cmb.Items.AddRange(_paramList.ToArray());
            cmb.Text = defaultVal;
            this.Controls.Add(cmb);
            return cmb;
        }

        private void BtnBrowse_Click(object sender, EventArgs e)
        {
            using (var dlg = new OpenFileDialog())
            {
                dlg.Filter = "Excel Files|*.xlsx;*.xls;*.xlsm";
                if (dlg.ShowDialog() == DialogResult.OK) txtPath.Text = dlg.FileName;
            }
        }

        private void BtnRun_Click(object sender, EventArgs e)
        {
            if (string.IsNullOrEmpty(txtPath.Text))
            {
                MessageBox.Show("Выберите файл!", "Ошибка", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }
            try
            {
                IdxRoom = ColumnLetterToNumber(txtRoomCol.Text);
                if (chkLoad.Checked || !string.IsNullOrEmpty(txtLoadCol.Text)) IdxLoad = ColumnLetterToNumber(txtLoadCol.Text);
                if (chkTemp.Checked || !string.IsNullOrEmpty(txtTempCol.Text)) IdxTemp = ColumnLetterToNumber(txtTempCol.Text);
                if (IdxRoom < 1) throw new Exception();
            }
            catch
            {
                MessageBox.Show("Некорректные буквы колонок!", "Ошибка", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return;
            }
            IsRun = true;
            this.DialogResult = DialogResult.OK;
            this.Close();
        }

        private int ColumnLetterToNumber(string columnName)
        {
            if (string.IsNullOrEmpty(columnName)) return 0;
            columnName = columnName.ToUpperInvariant().Trim();
            int sum = 0;
            foreach (char c in columnName)
            {
                sum *= 26;
                sum += (c - 'A' + 1);
            }
            return sum;
        }
    }
}