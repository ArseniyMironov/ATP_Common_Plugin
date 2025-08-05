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

            if (doc == null)
            {
                //message = "Активный документ не найден";
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
                        isUseDict = form.UseDictMoide;
                        DictPath = form.SelectedFilePath;
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
                        categoryKeys: new[] { "DuctCurves", "DuctFlexCurves", "PipeCurves", "PipeFlexCurves" }
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

                        foreach (var category in categories)
                        {
                            if (selecttionBuiltInInstance.selectInstanceOfCategory(doc, category) != null)
                            {
                                linearObjects.AddRange(selecttionBuiltInInstance.selectInstanceOfCategory(doc, category));
                            }
                        }

                        if (linearObjects.Count > 0)
                        {
                            SetSystemNameByDictLinearElements(linearObjects, tableData, docName);
                        }
                        else
                        {
                            logger.LogWarning("Не найдено линейных объектов для обработки", docName);
                            //Debug.WriteLine("Не найдено линейных объектов для обработки");
                        }

                        if (mechEquip.Count > 0)
                        {
                            SetSystemNameByDictMechEq(mechEquip, tableData, docName);
                        }
                        else
                        {
                            logger.LogWarning("Не найдено оборудования для обработки", docName);
                            //Debug.WriteLine("Не найдено оборудования для обработки");
                        }
                    }
                    else 
                        SetSystemName(doc, docName, systems);

                    tr.Commit();
                }

                //TaskDialog.Show("Готово", "Имя системы и группирование заполнено.");
                logger.LogInfo("Завершение заполнения параметра ИмяСистемы", docName);

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
                    catch
                    {
                        if (tr.HasStarted() && !tr.HasEnded())
                        {
                            tr.RollBack();
                        }
                        throw; // Перебрасываем исключение выше
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
            Dictionary<string, string> SystemsDict = new Dictionary<string, string> { };
            Dictionary<ElementId, bool> priorityElements = new Dictionary<ElementId, bool> { };

            var logger = ATP_App.GetService<ILoggerService>();

            foreach (Element system in systems)
            {
                Element systemType = doc.GetElement(system.GetTypeId());
                string abbreviation = systemType.get_Parameter(BuiltInParameter.RBS_SYSTEM_ABBREVIATION_PARAM).AsValueString();
                string typeComment = systemType.get_Parameter(BuiltInParameter.ALL_MODEL_TYPE_COMMENTS).AsValueString();
                string systemName = system.get_Parameter(BuiltInParameter.RBS_SYSTEM_NAME_PARAM).AsValueString();
                string newName = $"{abbreviation} - {typeComment}";
                SystemsDict[systemName] = newName;

                IList<Element> Contains = new List<Element> { };

                bool isPipingSystem = system.Category.Name == "Piping Systems";
                bool isDuctSystem = system.Category.Name == "Duct Systems";

                if (isPipingSystem)
                {
                    var sys = system as PipingSystem;
                    ElementSet pipeSystemElements = sys.PipingNetwork;
                    foreach (Element pipeSystemElement in pipeSystemElements)
                    {
                        Contains.Add(pipeSystemElement);
                    }
                }
                else
                {
                    var sys = system as MechanicalSystem;
                    ElementSet ductSystemElements = sys.DuctNetwork;
                    foreach (Element ductSystemElement in ductSystemElements)
                    {
                        Contains.Add(ductSystemElement);
                    }
                }

                foreach (Element element in Contains)
                {
                    if (element.Category.Name == "Mechanical Equipment" || element.Category.Name == "Plumbing Fixture" || element.Category.Name == "Pipe Accessories" || element.Category.Name == "Duct Accessories")
                    {
                        string names = element.get_Parameter(BuiltInParameter.RBS_SYSTEM_NAME_PARAM).AsValueString();

                        if (names.Contains(","))
                        {
                            priorityElements[element.Id] = isPipingSystem;
                        }
                        else
                        {
                            RevitUtils.SetParameterValue(element, "ИмяСистемы", newName);
                        }
                    }

                        RevitUtils.SetParameterValue(element, "ИмяСистемы", newName);
                }
            }

            foreach (ElementId elementId in priorityElements.Keys)
            {
                bool isPiping = priorityElements[elementId];
                Element element = doc.GetElement(elementId);
                string names = element.get_Parameter(BuiltInParameter.RBS_SYSTEM_NAME_PARAM).AsValueString();
                string[] namesList = names.Split(',');
                string search = "";
                string newName = "Оборудование без системы";

                if (!isPiping)
                {
                    char[] targetChars = { 'П', 'В' };
                    search = targetChars
                        .Where(c => names.Contains(c.ToString()))
                        .Select(c => namesList.FirstOrDefault(name => name.Contains(c.ToString())))
                        .FirstOrDefault(name => name != null);
                }
                else
                {
                    char[] targetChars = { 'Н', 'К', 'В', 'Т', 'Х' };
                    search = targetChars
                        .Where(c => names.Contains(c.ToString()))
                        .Select(c => namesList.FirstOrDefault(name => name.Contains(c.ToString())))
                        .FirstOrDefault(name => name != null);
                }

                newName = SystemsDict[search];

                RevitUtils.SetParameterValue(element, "ИмяСистемы", newName);
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
                if (string.IsNullOrEmpty(name))
                    continue;

                string newName = name;
                try
                {
                    newName = dictionary[name];
                }
                catch
                {
                    logger.LogWarning($"Система элемента не добавлена в словарь. Id элемента: {element.Id}", docName);
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
                        if (dictionary.TryGetValue(namePart, out string dictValue))
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
        public bool UseDictMoide { get; private set; }

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
                Filter = "Excel Files|* .xlsx;* .xls",
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
                UseDictMoide = useFileCheckBox.Checked;
                bool UseRevitData = useRevitDataCheckBox.Checked;
                if (UseDictMoide && string.IsNullOrWhiteSpace(filePathTextBox.Text))
                {
                    MessageBox.Show("Выберите словарь с именами систем или отключите режим его использование", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    e.Cancel = true;
                }
                if (!UseDictMoide && !UseRevitData)
                {
                    MessageBox.Show("Выберите режим работы", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    e.Cancel = true;
                }
            }
            base.OnFormClosing(e);
        }
    }
}