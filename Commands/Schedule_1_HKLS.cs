using ATP_Common_Plugin.Services;
using ATP_Common_Plugin.Utils;
using Autodesk.Revit.Attributes;
using Autodesk.Revit.DB;
using Autodesk.Revit.DB.Mechanical;
using Autodesk.Revit.DB.Plumbing;
using Autodesk.Revit.UI;
using System;
using System.Collections.Generic;
using System.Drawing;
using System.Linq;
using System.Runtime.InteropServices;
using System.Windows.Forms;

namespace ATP_Common_Plugin.Commands
{
    [Regeneration(RegenerationOption.Manual)]
    [Transaction(TransactionMode.Manual)]
    public class Schedule_1_HKLS : IExternalCommand
    {
        public Result Execute(ExternalCommandData commandData, ref string message, ElementSet elements)
        {
            UIApplication uiApp = commandData.Application;
            Document doc = uiApp.ActiveUIDocument?.Document;
            string docName = doc.Title;

            var logger = ATP_App.GetService<ILoggerService>();

            bool isUseDict;
            string DictPath;
            bool isUseModel;

            if (doc == null)
            {
                logger.LogError("Активный документ не найден", docName);
                return Result.Failed;
            }

            try
            {
                // Группирование
                // Сбор всех элементов по категориям
                var elementsByCategory = new Dictionary<string, IList<Element>>
                {
                    ["MechanicalEquipment"] = selecttionBuiltInInstance.selectInstanceOfCategory(doc, BuiltInCategory.OST_MechanicalEquipment),
                    ["DuctTerminal"] = selecttionBuiltInInstance.selectInstanceOfCategory(doc, BuiltInCategory.OST_DuctTerminal),
                    ["PlumbingFixtures"] = selecttionBuiltInInstance.selectInstanceOfCategory(doc, BuiltInCategory.OST_PlumbingFixtures),
                    ["DuctAccessory"] = selecttionBuiltInInstance.selectInstanceOfCategory(doc, BuiltInCategory.OST_DuctAccessory),
                    ["PipeAccessory"] = selecttionBuiltInInstance.selectInstanceOfCategory(doc, BuiltInCategory.OST_PipeAccessory),
                    ["DuctCurves"] = selecttionBuiltInInstance.selectInstanceOfCategory(doc, BuiltInCategory.OST_DuctCurves),
                    ["DuctFlexCurves"] = selecttionBuiltInInstance.selectInstanceOfCategory(doc, BuiltInCategory.OST_FlexDuctCurves),
                    ["PipeCurves"] = selecttionBuiltInInstance.selectInstanceOfCategory(doc, BuiltInCategory.OST_PipeCurves),
                    ["PipeFlexCurves"] = selecttionBuiltInInstance.selectInstanceOfCategory(doc, BuiltInCategory.OST_FlexPipeCurves),
                    ["DuctFittings"] = selecttionBuiltInInstance.selectInstanceOfCategory(doc, BuiltInCategory.OST_DuctFitting),
                    ["PipeFittings"] = selecttionBuiltInInstance.selectInstanceOfCategory(doc, BuiltInCategory.OST_PipeFitting),
                    ["DuctInsulation"] = selecttionBuiltInInstance.selectInstanceOfCategory(doc, BuiltInCategory.OST_DuctInsulations),
                    ["PipeInsulation"] = selecttionBuiltInInstance.selectInstanceOfCategory(doc, BuiltInCategory.OST_PipeInsulations)
                };

                IList<Element> pipeSystems = new FilteredElementCollector(doc)
                .OfCategory(BuiltInCategory.OST_PipingSystem)
                .WhereElementIsNotElementType()
                .ToList();

                IList<Element> ductSystems = new FilteredElementCollector(doc)
                .OfCategory(BuiltInCategory.OST_DuctSystem)
                .WhereElementIsNotElementType()
                .ToList();

                IList<Element> systems = pipeSystems.Concat(ductSystems).ToList();

                // Вызов формы
                using (var form = new ExcellDictSelectionForm())
                {
                    if (form.ShowDialog() == DialogResult.OK)
                    {
                        isUseDict = form.UseDictMode;
                        DictPath = form.SelectedFilePath;
                        isUseModel = form.UseRevitData;
                    }
                    else
                    {
                        logger.LogError("Ошибка чтения Excel");
                        return Result.Cancelled;
                    }

                }

                // Группировка операций по логическим блокам
                var operations = new List<CategoryOperation>
                {
                    new CategoryOperation(
                        name: "Группирование оборудование и приборы",
                        groupValue: "1",
                        categoryKeys: new[] { "MechanicalEquipment", "PlumbingFixtures" }
                    ),
                    new CategoryOperation(
                        name: "Группирование арматуры",
                        groupValue: "2",
                        categoryKeys: new[] { "DuctTerminal", "DuctAccessory", "PipeAccessory" }
                    ),
                    new CategoryOperation(
                        name: "Группирование линейных элементов",
                        groupValue: "3",
                        categoryKeys: new[] { "DuctCurves", "PipeCurves"}
                    ),
                    new CategoryOperation(
                        name: "Группирование соединительных элементов",
                        groupValue: "4",
                        categoryKeys: new[] { "DuctFittings", "PipeFittings" }
                    ),
                    new CategoryOperation(
                        name: "Группирование гибких линейных объектов",
                        groupValue: "5",
                        categoryKeys: new[] { "DuctFlexCurves", "PipeFlexCurves" }
                    ),
                    new CategoryOperation(
                        name: "Группирование изоляции",
                        groupValue: "6",
                        categoryKeys: new[] { "DuctInsulation", "PipeInsulation" }
                    )
                };

                // Выполнение операций с индивидуальными транзакциями
                ExecuteCategoryOperations(doc, elementsByCategory, operations);

                // Имена систем
                using (Transaction tr = new Transaction(doc, "Заполнение ИмяСистемы"))
                {
                    logger.LogInfo("Начало заполнения параметра ИмяСистемы", docName);
                    tr.Start();

                    if (isUseDict)
                    {
                        var tableData = ReadDict(DictPath, docName);
                        var linearObjects = new List<Element>();
                        var categories = new List<BuiltInCategory>
                        {
                            BuiltInCategory.OST_DuctCurves,
                            BuiltInCategory.OST_DuctTerminal,
                            BuiltInCategory.OST_DuctAccessory,
                            BuiltInCategory.OST_DuctFitting,
                            BuiltInCategory.OST_DuctInsulations,
                            BuiltInCategory.OST_DuctLinings,
                            BuiltInCategory.OST_FlexDuctCurves,
                            BuiltInCategory.OST_PipeCurves,
                            BuiltInCategory.OST_PipeAccessory,
                            BuiltInCategory.OST_PipeFitting,
                            BuiltInCategory.OST_PipeInsulations,
                            BuiltInCategory.OST_FlexPipeCurves

                        };

                        foreach (var category in categories)
                        {
                            if (selecttionBuiltInInstance.selectInstanceOfCategory(doc, category) != null)
                            {
                                linearObjects.AddRange(selecttionBuiltInInstance.selectInstanceOfCategory(doc, category));
                            }
                        }

                        var mechEquip = new List<Element>();
                        var equipCategories = new List<BuiltInCategory>
                        {
                            BuiltInCategory.OST_MechanicalEquipment,
                            BuiltInCategory.OST_PlumbingFixtures,
                        };

                        foreach (var category in equipCategories)
                        {
                            var list = selecttionBuiltInInstance.selectInstanceOfCategory(doc, category);
                            if (list != null && list.Count > 0)
                                mechEquip.AddRange(list);
                        }

                        if (linearObjects.Count > 0)
                        {
                            SetSystemNameByDictLinearElements(linearObjects, tableData, docName);
                        }
                        else
                        {
                            logger.LogWarning("Не найдено линейных объектов для обработки", docName);
                        }

                        if (mechEquip.Count > 0)
                        {
                            SetSystemNameByDictMechEq(mechEquip, tableData, docName);
                        }
                        else
                        {
                            logger.LogWarning("Не найдено оборудования для обработки", docName);
                        }

                    }
                    else if (isUseModel)
                    {
                        SetSystemName(doc, docName, systems);
                    }

                    tr.Commit();
                }

                //TaskDialog.Show("Готово", "Имя системы и группирование заполнено.");
                logger.LogInfo("Завершение заполнения параметра ИмяСистемы", docName);
                TaskDialog.Show("Успех", "Параметры ИмяСистемы и ADSK_Группирование заполнены");

                return Result.Succeeded;
            }
            catch (Exception ex)
            {
                //message = $"Ошибка: {ex.Message}\n{ex.StackTrace}";
                //TaskDialog.Show("Ошибка", message);
                logger.LogError($"Ошибка: {ex.Message}\n{ex.StackTrace}", docName);
                return Result.Failed;
            }
        }

        // Основной метод выполнения операций
        private void ExecuteCategoryOperations(
            Document doc,
            Dictionary<string, IList<Element>> elementsByCategory,
            List<CategoryOperation> operations)
        {
            var logger = ATP_App.GetService<ILoggerService>();
            string docName = doc.Title;
            foreach (var operation in operations)
            {
                using (Transaction tr = new Transaction(doc, operation.Name))
                {
                    try
                    {
                        tr.Start();

                        foreach (var categoryKey in operation.CategoryKeys)
                        {
                            if (elementsByCategory.TryGetValue(categoryKey, out var elements) && elements?.Count > 0)
                            {
                                ProcessElementsBatch(elements, operation.GroupValue, docName);
                            }
                        }

                        if (tr.HasStarted() && !tr.HasEnded())
                        {
                            tr.Commit();
                        } 
                    }
                    catch (Exception ex)
                    {
                        logger.LogError($"Ошибка при выполнении операции '{operation.Name}': {ex.Message}", docName);
                        if (tr.HasStarted() && !tr.HasEnded())
                        {
                            tr.RollBack();
                        }
                    }
                }
            }
        }

        // Обработка элементов с чанкированием для больших коллекций
        private void ProcessElementsBatch(IList<Element> elements, string groupValue, string docName, int batchSize = 500)
        {
            int totalElements = elements.Count;
            int processed = 0;
            var logger = ATP_App.GetService<ILoggerService>();

            while (processed < totalElements)
            {
                int remaining = totalElements - processed;
                int currentBatchSize = Math.Min(batchSize, remaining);

                var batch = elements
                    .Skip(processed)
                    .Take(currentBatchSize)
                    .ToList();

                foreach (var element in batch)
                {
                    try
                    {
                        RevitUtils.SetParameterValue(element, dictionaryGUID.ADSKGroup, groupValue);
                    }
                    catch (Exception ex)
                    {
                        // Логирование ошибки без прерывания процесса
                        logger.LogError($"Ошибка элемента {element.Id}: {ex.Message}", docName);
                        //Debug.WriteLine($"Ошибка элемента {element.Id}: {ex.Message}");
                    }
                }

                processed += currentBatchSize;
            }
        }

        // Вспомогательный класс для группировки операций
        private class CategoryOperation
        {
            public string Name { get; }
            public string GroupValue { get; }
            public string[] CategoryKeys { get; }
            public CategoryOperation(string name, string groupValue, string[] categoryKeys)
            {
                Name = name;
                GroupValue = groupValue;
                CategoryKeys = categoryKeys;
            }
        }

        private void SetSystemName(Document doc, string docName, IList<Element> systems)
        {
            // Ключи — без учёта регистра, чтобы мэппинг стабильно находился
            var systemsNameMap = new Dictionary<string, string>(StringComparer.InvariantCultureIgnoreCase);
            var systemKindMap = new Dictionary<string, bool>(StringComparer.InvariantCultureIgnoreCase);
            var deferred = new HashSet<ElementId>(); // элементы с несколькими системами или пустыми именами

            // Локальные функции
            Func<string, string> normalize = s => (s ?? string.Empty).Trim();
            Func<Element, bool> IsPipingSystem = e =>
                e.Category != null && e.Category.Id.IntegerValue == (int)BuiltInCategory.OST_PipingSystem;
            Func<Element, bool> IsDuctSystem = e =>
                e.Category != null && e.Category.Id.IntegerValue == (int)BuiltInCategory.OST_DuctSystem;

            // Чем меньше ранг, тем выше приоритет
            // Piping: Н (Напорные системы), К (Канализаця), В (ХВС), Т (ГВС), Х (ХС)
            // Duct:   П (Пприток), В (Вытяжка)
            char[] pipingPriority = { 'Н', 'К', 'В', 'Т', 'Х' };
            char[] ductPriority = { 'П', 'В' };

            Func<string, bool, int> getRank = (name, isPiping) =>
            {
                if (string.IsNullOrEmpty(name)) return int.MaxValue;
                var arr = isPiping ? pipingPriority : ductPriority;
                // Ранг по первой найденной букве-приориту
                for (int i = 0; i < arr.Length; i++)
                    if (name.IndexOf(arr[i].ToString(), StringComparison.InvariantCultureIgnoreCase) >= 0)
                        return i;
                return arr.Length + 1; // «хуже» известных приоритетов
            };

            // 1) Построим мэппинг systemNameKey → (новое имя + тип системы)
            foreach (Element system in systems)
            {
                if (system == null) continue;

                var systemType = doc.GetElement(system.GetTypeId());
                if (systemType == null) continue;

                // Текстовые параметры типа
                var abbreviation = systemType.get_Parameter(BuiltInParameter.RBS_SYSTEM_ABBREVIATION_PARAM)?.AsString() ?? string.Empty;
                var typeComment = systemType.get_Parameter(BuiltInParameter.ALL_MODEL_TYPE_COMMENTS)?.AsString() ?? string.Empty;

                // Имя системы (у самой системы)
                var systemNameRaw = system.get_Parameter(BuiltInParameter.RBS_SYSTEM_NAME_PARAM)?.AsString()
                                    ?? system.get_Parameter(BuiltInParameter.RBS_SYSTEM_NAME_PARAM)?.AsValueString();
                var systemNameKey = normalize(systemNameRaw);

                var newName = $"{abbreviation} - {typeComment}";
                if (newName == "-" || string.IsNullOrWhiteSpace(newName)) newName = abbreviation.Length > 0 ? abbreviation : typeComment;

                // Сохраняем мэппинг имени системы → то, что будем писать в "ИмяСистемы"
                if (systemNameKey.Length > 0)
                {
                    systemsNameMap[systemNameKey] = newName;

                    // Запомним вид системы (piping/duct) для корректного ранжирования дальше
                    if (IsPipingSystem(system)) systemKindMap[systemNameKey] = true;
                    else if (IsDuctSystem(system)) systemKindMap[systemNameKey] = false;
                }

                // 2) Соберём элементы сети (без дублей)
                var members = new HashSet<ElementId>();
                if (IsPipingSystem(system))
                {
                    var ps = system as PipingSystem;
                    if (ps?.PipingNetwork != null)
                        foreach (Element m in ps.PipingNetwork)
                            if (m != null) members.Add(m.Id);
                }
                else
                {
                    var ms = system as MechanicalSystem;
                    if (ms?.DuctNetwork != null)
                        foreach (Element m in ms.DuctNetwork)
                            if (m != null) members.Add(m.Id);
                }

                // 3) Обработаем найденные элементы сети (только разрешённые категории)
                foreach (var id in members)
                {
                    var element = doc.GetElement(id);
                    if (element == null) continue;

                    var names = element.get_Parameter(BuiltInParameter.RBS_SYSTEM_NAME_PARAM)?.AsString()
                                ?? element.get_Parameter(BuiltInParameter.RBS_SYSTEM_NAME_PARAM)?.AsValueString()
                                ?? string.Empty;

                    var hasComma = names.IndexOf(',') >= 0;
                    var isEmpty = string.IsNullOrWhiteSpace(names);

                    if (hasComma || isEmpty)
                    {
                        deferred.Add(id);
                    }
                    else
                    {
                        var key = normalize(names);
                        string mapped;
                        if (key.Length > 0 && systemsNameMap.TryGetValue(key, out mapped))
                        {
                            RevitUtils.SetParameterValue(element, "ИмяСистемы", mapped);
                        }
                        else
                        {
                            // Если ключ не найден
                        }
                    }
                }
            }

            // 4) Финальная фаза: мультисистемные / пустые
            foreach (var elementId in deferred)
            {
                var element = doc.GetElement(elementId);
                if (element == null) continue;

                var names = element.get_Parameter(BuiltInParameter.RBS_SYSTEM_NAME_PARAM)?.AsString()
                            ?? element.get_Parameter(BuiltInParameter.RBS_SYSTEM_NAME_PARAM)?.AsValueString()
                            ?? string.Empty;

                var namesList = names
                    .Split(new[] { ',' }, StringSplitOptions.RemoveEmptyEntries)
                    .Select(s => normalize(s))
                    .Where(s => s.Length > 0)
                    .Distinct(StringComparer.InvariantCultureIgnoreCase)
                    .ToArray();

                string chosenMapped = null;
                int bestRank = int.MaxValue;

                // Попробуем выбрать «лучшую» систему по таблице приоритетов
                foreach (var n in namesList)
                {
                    string mapped;
                    if (!systemsNameMap.TryGetValue(n, out mapped)) continue;

                    bool isPiping;
                    // Если вид системы неизвестен — попробуем угадать по наличию в обоих наборах: по умолчанию хуже любых известных
                    var hasKind = systemKindMap.TryGetValue(n, out isPiping);
                    int rank = hasKind ? getRank(n, isPiping) : int.MaxValue - 1;

                    if (rank < bestRank)
                    {
                        bestRank = rank;
                        chosenMapped = mapped;
                    }
                }

                if (string.IsNullOrEmpty(chosenMapped))
                {
                    // Фолбэк: если ничего не сматчилось, берем первое присутствующее в мэппинге
                    foreach (var n in namesList)
                    {
                        string mapped;
                        if (systemsNameMap.TryGetValue(n, out mapped))
                        {
                            chosenMapped = mapped;
                            break;
                        }
                    }
                }

                if (string.IsNullOrEmpty(chosenMapped))
                {
                    chosenMapped = "Оборудование без системы";
                }

                RevitUtils.SetParameterValue(element, "ИмяСистемы", chosenMapped);
            }
        }

        private void SetSystemNameByDictLinearElements(IList<Element> elements, Dictionary<string, string> dictionary, string docName)
        {
            var logger = ATP_App.GetService<ILoggerService>();

            if (elements == null)
            {
                logger.LogError($"Нет линейных элементов для обработки", docName);
                return; 
            }

            foreach(Element element in elements)
            {
                var nameParam = element.get_Parameter(BuiltInParameter.RBS_SYSTEM_NAME_PARAM);
                if (nameParam == null || !nameParam.HasValue)
                    continue;

                string name = nameParam.AsValueString();
                if (string.IsNullOrWhiteSpace(name))
                    continue;

                string key = name.Trim();
                string newName;
                if (!dictionary.TryGetValue(key, out newName))
                {
                    logger.LogWarning($"Система элемента не добавлена в словарь: «{key}». Id элемента: {element.Id}", docName);
                    newName = key;
                }

                RevitUtils.SetParameterValue(element, "ИмяСистемы", newName);
            }
        }

        private void SetSystemNameByDictMechEq(IList<Element> elements, Dictionary<string, string> dictionary, string docName)
        {
            var logger = ATP_App.GetService<ILoggerService>();

            if (elements == null)
            {
                return;
            }

            foreach (Element element in elements)
            {
                try
                {
                    var nameParam = element.get_Parameter(BuiltInParameter.RBS_SYSTEM_NAME_PARAM);
                    if (nameParam == null || !nameParam.HasValue)
                        continue;

                    string name = nameParam.AsValueString();
                    if (string.IsNullOrEmpty(name))
                        continue;

                    string[] priorityNames = name.Split(',');
                    string systemName = "Оборудование без системы";

                    foreach (string namePart in priorityNames)
                    {
                        var key = (namePart ?? string.Empty).Trim();
                        if (key.Length == 0) continue;

                        if (dictionary.TryGetValue(key, out string dictValue))
                        {
                            systemName = dictValue;
                            break;
                        }
                    }

                    var param = element.LookupParameter("ИмяСистемы");
                    if (param != null && param.IsReadOnly == false)
                    {
                        param.Set(systemName);
                    }
                }
                catch (Exception ex)
                {
                    logger.LogWarning($"Ошибка элемента {element.Id}: {ex.Message}", docName);
                }
            }
        }

        private Dictionary<string, string> ReadDict(string path, string docName)
        {
            var logger = ATP_App.GetService<ILoggerService>();
            try
            {
                return ExcelUtils.ReadSystemDict(
                    filePath: path,
                    worksheetName: "Title"
                );
            }
            catch (COMException ex) when (ex.ErrorCode == -2147221164)
            {
                logger.LogError("Несовместимая версия Excel. Требуется Excel 2010 или новее", docName);
                throw new Exception("Несовместимая версия Excel. Требуется Excel 2010 или новее");
            }
            catch (Exception ex)
            {
                logger.LogError($"Ошибка Excel, {ex.Message}", docName);
                return new Dictionary<string, string>();
            }
        }
    }

    public class ExcellDictSelectionForm : System.Windows.Forms.Form
    {
        private System.Windows.Forms.Button browseButton;
        private System.Windows.Forms.TextBox filePathTextBox;
        private System.Windows.Forms.CheckBox useRevitDataCheckBox;
        private System.Windows.Forms.CheckBox useFileCheckBox;
        private System.Windows.Forms.Button okButton;
        private System.Windows.Forms.Button cancelButton;
        private OpenFileDialog openFileDialog;

        public string SelectedFilePath { get; private set; }
        public bool UseDictMode { get; private set; }
        public bool UseRevitData { get; private set; }

        public ExcellDictSelectionForm()
        {
            InitializeCoomponents();
        }

        private void InitializeCoomponents()
        {
            this.Text = "Заполнение имени системы";
            this.Size = new Size(450, 200);
            this.FormBorderStyle = FormBorderStyle.FixedDialog;
            this.StartPosition = FormStartPosition.CenterScreen;
            this.MaximizeBox = false;
            this.MinimizeBox = false;

            // CheckBox for switch
            useFileCheckBox = new System.Windows.Forms.CheckBox
            {
                Text = "Использовать словарь Excel",
                Checked = false,
                Location = new System.Drawing.Point(20, 25),
                AutoSize = true
            };
            useFileCheckBox.CheckedChanged += UseFileCheckBox_CheckedChanged;

            useRevitDataCheckBox = new System.Windows.Forms.CheckBox
            {
                Text = "Использовать данные в модели",
                Checked = true,
                Location = new System.Drawing.Point(20, 5),
                AutoSize = true
            };
            useRevitDataCheckBox.CheckedChanged += UseRevitDataCheckBox_CheckedChanged;

            filePathTextBox = new System.Windows.Forms.TextBox
            {
                Location = new System.Drawing.Point(20, 55),
                Size = new System.Drawing.Size(300, 20),
                ReadOnly = true
            };

            browseButton = new System.Windows.Forms.Button
            {
                Text = "Browse...",
                Location = new System.Drawing.Point(330, 52),
                Size = new System.Drawing.Size(80, 23)
            };
            browseButton.Click += BrowseButton_Click;

            okButton = new System.Windows.Forms.Button
            {
                Text = "Ok",
                DialogResult = DialogResult.OK,
                Location = new System.Drawing.Point(250, 105),
                Size = new System.Drawing.Size(80, 30)
            };

            cancelButton = new System.Windows.Forms.Button
            {
                Text = "Cancel",
                DialogResult = DialogResult.Cancel,
                Location = new System.Drawing.Point(330, 105),
                Size = new System.Drawing.Size(80, 30)
            };

            openFileDialog = new OpenFileDialog
            {
                Filter = "Excel Files|*.xlsx;*.xls",
                Title = "Выберете словарь систем"
            };

            this.Controls.Add(useRevitDataCheckBox);
            this.Controls.Add(useFileCheckBox);
            this.Controls.Add(filePathTextBox);
            this.Controls.Add(browseButton);
            this.Controls.Add(okButton);
            this.Controls.Add(cancelButton);

            UpdateControls();
        }

        private void UseFileCheckBox_CheckedChanged(object sender, EventArgs e)
        {
            // Если флажок включен - выключаем другой
            if (useFileCheckBox.Checked)
            {
                useRevitDataCheckBox.Checked = false;
            }
            // Если оба флажка выключены - включаем текущий обратно
            else if (!useRevitDataCheckBox.Checked)
            {
                useFileCheckBox.Checked = true;
                return;
            }
            UpdateControls();
        }

        private void UpdateControls()
        {
            bool useDict = useFileCheckBox.Checked;
            filePathTextBox.Enabled = useDict;
            browseButton.Enabled = useDict;
        }

        private void UseRevitDataCheckBox_CheckedChanged(object sender, EventArgs e)
        {
            // Если флажок включен - выключаем другой
            if (useRevitDataCheckBox.Checked)
            {
                useFileCheckBox.Checked = false;
            }
            // Если оба флажка выключены - включаем текущий обратно
            else if (!useFileCheckBox.Checked)
            {
                useRevitDataCheckBox.Checked = true;
                return;
            }
            UpdateControls();
        }

        private void BrowseButton_Click(object sender, EventArgs e)
        {
            if (openFileDialog.ShowDialog() == DialogResult.OK)
            {
                filePathTextBox.Text = openFileDialog.FileName;
                SelectedFilePath = openFileDialog.FileName;
            }
        }

        protected override void OnFormClosing(FormClosingEventArgs e)
        {
            if (this.DialogResult == DialogResult.OK)
            {
                UseDictMode = useFileCheckBox.Checked;
                UseRevitData = useRevitDataCheckBox.Checked;
                if (UseDictMode && string.IsNullOrWhiteSpace(filePathTextBox.Text))
                {
                    MessageBox.Show("Выберите словарь с именами систем или отключите режим его использование", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    e.Cancel = true;
                }
                if (!UseDictMode && !UseRevitData)
                {
                    MessageBox.Show("Выберите режим работы", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    e.Cancel = true;
                }
            }
            base.OnFormClosing(e);
        }
    }
}