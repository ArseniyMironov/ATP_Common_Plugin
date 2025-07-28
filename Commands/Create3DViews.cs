using Autodesk.Revit.Attributes;
using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using System;
using System.Collections.Generic;
using System.Drawing;
using System.Linq;
using System.Threading;
using System.Threading.Tasks;
using System.Windows.Forms;
using Application = Autodesk.Revit.ApplicationServices.Application;
using Form = System.Windows.Forms.Form;
using TextBox = System.Windows.Forms.TextBox;

namespace ATP_Common_Plugin.Commands
{
    [Regeneration(RegenerationOption.Manual)]
    [Transaction(TransactionMode.Manual)]
    public class Create3DViews :IExternalCommand
    {
        private Document doc;
        private List<string> viewNames;
        private List<string> systemNameDuctUnique;
        private List<string> systemNamePipeUnique;
        private List<string> checkedNames;

        public Result Execute(ExternalCommandData commandData, ref string message, ElementSet elements)
        {
            if (commandData?.Application?.ActiveUIDocument?.Document == null)
                return Result.Failed;

            this.doc = commandData.Application.ActiveUIDocument.Document;
            UIDocument uidoc = commandData.Application.ActiveUIDocument;
            Application app = commandData.Application.Application;

            try
            {
                // Получаем все элементы нужных категорий
                var massElementsDuct = new List<FilteredElementCollector>
                {
                    new FilteredElementCollector(this.doc).WhereElementIsNotElementType().OfCategory(BuiltInCategory.OST_DuctFitting),
                    new FilteredElementCollector(this.doc).WhereElementIsNotElementType().OfCategory(BuiltInCategory.OST_DuctCurves),
                    new FilteredElementCollector(this.doc).WhereElementIsNotElementType().OfCategory(BuiltInCategory.OST_DuctAccessory),
                    new FilteredElementCollector(this.doc).WhereElementIsNotElementType().OfCategory(BuiltInCategory.OST_FlexDuctCurves),
                    new FilteredElementCollector(this.doc).WhereElementIsNotElementType().OfCategory(BuiltInCategory.OST_FlexPipeCurves)
                };

                var massElementsPipe = new List<FilteredElementCollector>
                {
                    new FilteredElementCollector(this.doc).WhereElementIsNotElementType().OfCategory(BuiltInCategory.OST_PipeFitting),
                    new FilteredElementCollector(this.doc).WhereElementIsNotElementType().OfCategory(BuiltInCategory.OST_PipeCurves),
                    new FilteredElementCollector(this.doc).WhereElementIsNotElementType().OfCategory(BuiltInCategory.OST_PipeAccessory),
                    new FilteredElementCollector(this.doc).WhereElementIsNotElementType().OfCategory(BuiltInCategory.OST_FlexPipeCurves),
                    new FilteredElementCollector(this.doc).WhereElementIsNotElementType().OfCategory(BuiltInCategory.OST_PipeInsulations),
                    new FilteredElementCollector(this.doc).WhereElementIsNotElementType().OfCategory(BuiltInCategory.OST_PlumbingFixtures)
                };

                // Собираем имена систем воздуховодов
                var systemNameDuctPV = new List<string>();
                var systemNameDuctPid = new List<ElementId>();

                foreach (var collection in massElementsDuct)
                {
                    foreach (Element e in collection)
                    {
                        Parameter systemNameParam = e.LookupParameter("ИмяСистемы");
                        if (systemNameParam != null && !string.IsNullOrEmpty(systemNameParam.AsString()))
                        {
                            systemNameDuctPV.Add(systemNameParam.AsString());
                            systemNameDuctPid.Add(systemNameParam.Id);
                        }
                    }
                }

                systemNameDuctUnique = systemNameDuctPV.Distinct().ToList();

                // Собираем имена систем трпубопроводов
                var systemNamePipePV = new List<string>();
                var systemNamePipePid = new List<ElementId>();

                foreach (var collection in massElementsPipe)
                {
                    foreach (Element e in collection)
                    {
                        Parameter systemNameParam = e.LookupParameter("ИмяСистемы");
                        if (systemNameParam != null && !string.IsNullOrEmpty(systemNameParam.AsString()))
                        {
                            systemNamePipePV.Add(systemNameParam.AsString());
                            systemNamePipePid.Add(systemNameParam.Id);
                        }
                    }
                }

                systemNamePipeUnique = systemNamePipePV.Distinct().ToList();

                // Получаем все 3D виды и планы для проверки имен
                var allViews = new FilteredElementCollector(doc)
                    .OfClass(typeof(View3D))
                    .UnionWith(new FilteredElementCollector(doc).OfClass(typeof(ViewPlan)))
                    .ToElements();

                viewNames = allViews.Select(v => v.Name).ToList();

                // Показываем окно выбора систем
                using (var form = new SelectionForm(systemNameDuctUnique, systemNamePipeUnique)) //
                {
                    if (form.ShowDialog() == DialogResult.OK)
                    {
                        checkedNames = form.SelectedItems;
                    }
                    else
                    {
                        return Result.Cancelled;
                    }
                }

                // Категории для фильтра
                var filter1Categories = new List<ElementId>
                {
                    new ElementId(BuiltInCategory.OST_DuctFitting),
                    new ElementId(BuiltInCategory.OST_MechanicalEquipment),
                    new ElementId(BuiltInCategory.OST_DuctCurves),
                    new ElementId(BuiltInCategory.OST_DuctAccessory),
                    new ElementId(BuiltInCategory.OST_DuctTerminal),
                    new ElementId(BuiltInCategory.OST_DuctInsulations),
                    new ElementId(BuiltInCategory.OST_PipeCurves),
                    new ElementId(BuiltInCategory.OST_PipeFitting),
                    new ElementId(BuiltInCategory.OST_PipeAccessory),
                    new ElementId(BuiltInCategory.OST_PipeInsulations),
                    new ElementId(BuiltInCategory.OST_FlexPipeCurves),
                    new ElementId(BuiltInCategory.OST_FlexDuctCurves)
                };

                // Категории для скрытия
                var viewCategoriesToHide = new List<ElementId>
                {
                    new ElementId(BuiltInCategory.OST_Levels),
                    new ElementId(BuiltInCategory.OST_Sprinklers),
                    new ElementId(BuiltInCategory.OST_TelephoneDevices),
                    new ElementId(BuiltInCategory.OST_GenericModel),
                    new ElementId(BuiltInCategory.OST_Lines),
                    new ElementId(BuiltInCategory.OST_VolumeOfInterest),
                    new ElementId(BuiltInCategory.OST_Grids)
                };

                // Базовые точки
                var basicPoints = new FilteredElementCollector(doc)
                    .WhereElementIsNotElementType()
                    .OfCategory(BuiltInCategory.OST_ProjectBasePoint)
                    .UnionWith(new FilteredElementCollector(doc)
                    .WhereElementIsNotElementType()
                    .OfCategory(BuiltInCategory.OST_SharedBasePoint))
                    .ToElements()
                    .Select(e => e.Id)
                    .ToList();

                // Ссылки на RVT
                var rvtLinks = new FilteredElementCollector(doc)
                    .OfCategory(BuiltInCategory.OST_RvtLinks)
                    .OfClass(typeof(RevitLinkInstance))
                    .Select(e => e.Id)
                    .ToList();

                // CAD импорты
                var cadImports = new FilteredElementCollector(doc)
                    .OfClass(typeof(ImportInstance))
                    .Select(e => e.Id)
                    .ToList();

                // Все фильтры параметров
                var allParameterFilters = new FilteredElementCollector(doc)
                    .OfClass(typeof(ParameterFilterElement))
                    .Cast<ParameterFilterElement>()
                    .ToList();

                int counter = 0;

                using (Transaction tr = new Transaction(doc, "Создание 3D Вида"))
                {
                    tr.Start();

                    foreach (string sn in checkedNames)
                    {
                        string name = FindName(sn);
                        View3D view = Create3DView(name);

                        ElementId parameterId = systemNameDuctPid.First();
                        FilterRule filterRule = ParameterFilterRuleFactory.CreateNotEqualsRule(parameterId, sn, false);
                        ElementParameterFilter epf = new ElementParameterFilter(filterRule);

                        string filterName = $"Вентиляция. Имя системы НЕ равно {sn}";
                        ParameterFilterElement pfe;

                        if (ParameterFilterElement.IsNameUnique(doc, filterName))
                        {
                            pfe = ParameterFilterElement.Create(doc, filterName, filter1Categories, epf);
                        }
                        else
                        {
                            pfe = allParameterFilters.FirstOrDefault(f => f.Name == filterName);
                        }

                        view.SetFilterVisibility(pfe.Id, false);
                        HideCategories(view, viewCategoriesToHide);

                        view.DetailLevel = ViewDetailLevel.Fine;
                        view.SaveOrientationAndLock();

                        if (rvtLinks.Any())
                        {
                            view.HideElements(rvtLinks.ToList());
                        }

                        if (cadImports.Any())
                        {
                            view.HideElements(cadImports.ToList());
                        }

                        if (basicPoints.Any())
                        {
                            view.HideElements(basicPoints.ToList());
                        }

                        SetParameterValue(view.LookupParameter("ATP_Браузер проекта_Уровень 1"), "001_Modeling area");

                        if (sn.Contains("ПД") || sn.Contains("ДУ"))
                        {
                            SetParameterValue(view.LookupParameter("ADSK_Штамп Раздел проекта"), "315_Противодымная вентиляция");
                        }
                        else if (sn.Contains("П") || sn.Contains("В"))
                        {
                            SetParameterValue(view.LookupParameter("ADSK_Штамп Раздел проекта"), "310_Общеобменная вентиляция");
                        }
                        else if (sn.Contains("Х1") || sn.Contains("Х2") || sn.Contains("Х3") || sn.Contains("Х4") || sn.Contains("Х5") || sn.Contains("ХA"))
                        {
                            SetParameterValue(view.LookupParameter("ADSK_Штамп Раздел проекта"), "330_Холодоснабжение");
                        }
                        else if (sn.Contains("Т1") || sn.Contains("Т2"))
                        {
                            SetParameterValue(view.LookupParameter("ADSK_Штамп Раздел проекта"), "320_Отопление");
                        }
                        else if (sn.Contains("В21") || sn.Contains("В22") || sn.Contains("В2"))
                        {
                            SetParameterValue(view.LookupParameter("ADSK_Штамп Раздел проекта"), "355_Автоматическая система пожаротушения");
                        }
                        else if (sn.Contains("В1") || sn.Contains("Т3") || sn.Contains("Т4") || sn.Contains("В3"))
                        {
                            SetParameterValue(view.LookupParameter("ADSK_Штамп Раздел проекта"), "340_Водоснабжение");
                        }
                        else if (sn.Contains("К1") || sn.Contains("К2") || sn.Contains("К3") || sn.Contains("К4"))
                        {
                            SetParameterValue(view.LookupParameter("ADSK_Штамп Раздел проекта"), "345_Канализация");
                        }
                        else
                        {
                            SetParameterValue(view.LookupParameter("ADSK_Штамп Раздел проекта"), "300_ATP ID");
                        }

                        counter++;
                    }

                    tr.Commit();
                }

                TaskDialog.Show("Готово", $"Колличество успешно созданых 3D видов - {counter}");
                return Result.Succeeded;
            }
            catch (Exception ex)
            {
                TaskDialog.Show("Error", ex.Message);
                return Result.Failed;
            }
        }

        /// <summary>
        /// Вспомогательный метод по созданию 3D вида
        /// </summary>
        /// <param name="name"></param>
        /// <param name="existingView"></param>
        /// <returns></returns>
        /// <exception cref="InvalidOperationException"></exception>
        private View3D Create3DView(string name, View3D existingView = null)
        {
            if (existingView == null)
            {
                if (doc == null)
                    throw new InvalidOperationException("Document is not initialized!");

                var viewFamilyTypes = new FilteredElementCollector(this.doc)
                    .OfClass(typeof(ViewFamilyType))
                    .Cast<ViewFamilyType>()
                    .FirstOrDefault(vft => vft.ViewFamily == ViewFamily.ThreeDimensional);

                View3D view = View3D.CreateIsometric(this.doc, viewFamilyTypes.Id);
                view.Name = name;
                return view;
            }
            return existingView;
        }

        /// <summary>
        /// Вспомогательный метод по генерации имени для 3D вида
        /// </summary>
        /// <param name="systemName"></param>
        /// <param name="counter"></param>
        /// <returns></returns>
        private string FindName(string systemName, int counter = 0)
        {
            string name;
            if (counter == 0)
            {
                name = $"Схема системы {systemName}";
            }
            else
            {
                name = $"Схема системы {systemName} ({counter})";
            }

            if (viewNames.Contains(name))
            {
                counter++;
                return FindName(systemName, counter);
            }
            return name;
        }

        /// <summary>
        /// Вспомогательный метод для скрытия элементов категории
        /// </summary>
        /// <param name="view"></param>
        /// <param name="categories"></param>
        private void HideCategories(Autodesk.Revit.DB.View view, List<ElementId> categories)
        {
            foreach (ElementId categoryId in categories)
            {
                view.SetCategoryHidden(categoryId, true);
            }
        }

        private void SetParameterValue(Parameter parameter, string value)
        {
            if (parameter != null && parameter.StorageType == StorageType.String)
            {
                parameter.Set(value);
            }
        }
    }

    public class SelectionForm : Form
    {
        private ListView listView;
        private TextBox txtFilter;
        private Button btnConfirm;
        private System.Windows.Forms.Label lblStatus;
        private Button btnShowDuctSystems;
        private Button btnShowPipeSystems;

        private List<string> ductSystems;
        private List<string> pipeSystems;
        private List<string> currentItems = new List<string>();
        private Dictionary<string, bool> selectionState = new Dictionary<string, bool>();
        private CancellationTokenSource filterCancellationTokenSource;
        private const string PlaceholderText = "Фильтр...";
        private bool isPlaceholderActive = true;
        private bool showDuctSystems = true;

        public List<string> SelectedItems => selectionState
            .Where(kv => kv.Value)
            .Select(kv => kv.Key)
            .ToList();

        public SelectionForm(List<string> systemNamesDuct, List<string> systemNamesPipe) //
        {
            ductSystems = systemNamesDuct;
            pipeSystems = systemNamesPipe;
            InitializeSelectionState();
            InitializeComponents();
            LoadCurrentSystems();
        }

        public void InitializeSelectionState()
        {
            foreach (var item in ductSystems.Concat(pipeSystems).Distinct())
            {
                if (!selectionState.ContainsKey(item))
                    selectionState[item] = false;
            }
        }

        private void InitializeComponents() 
        {
            this.Text = "Выбор систем для создания схем";
            this.Size = new Size(650, 600);
            this.MinimumSize = new Size(500, 500);
            this.MaximumSize = new Size(600, 600);
            this.StartPosition = FormStartPosition.CenterScreen;


            // Control panel
            var filterPanel = new System.Windows.Forms.Panel { Dock = DockStyle.Top, Height = 40 };
            txtFilter = new TextBox { Dock = DockStyle.Fill, Margin = new Padding(5) };
            SetupPlaceholder();
            txtFilter.GotFocus += RemovePlaceholder;
            txtFilter.LostFocus += SetPlaceholder;
            txtFilter.TextChanged += OnFilterTextChangedAsync;

            // switch system panel
            var switchPanel = new System.Windows.Forms.Panel { Dock = DockStyle.Top, Height = 40 };
            btnShowDuctSystems = new Button
            {
                Text = "Вентиляционные системы",
                Dock = DockStyle.Left,
                Width = 180,
                BackColor = System.Drawing.Color.LightGreen
            };
            btnShowPipeSystems = new Button
            {
                Text = "Трубопроводные системы",
                Dock = DockStyle.Left,
                Width = 180,
                BackColor = System.Drawing.Color.LightGreen
            };
            btnShowDuctSystems.Click += (s, e) => SwitchSystemType(true);
            btnShowPipeSystems.Click += (s, e) => SwitchSystemType(false);
            switchPanel.Controls.Add(btnShowDuctSystems);
            switchPanel.Controls.Add(btnShowPipeSystems);


            // status bar
            lblStatus = new System.Windows.Forms.Label
            {
                Dock = DockStyle.Bottom,
                Height = 20,
                TextAlign = ContentAlignment.MiddleLeft
            };

            // List view
            listView = new ListView // перевести на VirtualMod?
            {
                Dock = DockStyle.Fill,
                CheckBoxes = true,
                FullRowSelect = true,
                View = System.Windows.Forms.View.Details,
                Columns = { new ColumnHeader { Width = -2, Text = "Системы"} },
                BorderStyle = BorderStyle.FixedSingle,
                HideSelection = false
            };
            listView.ItemChecked += OnItemChecked;

            // Confirm button
            btnConfirm = new Button
            {
                Dock = DockStyle.Bottom,
                Height = 40,
                Text = "Создать",
                BackColor = System.Drawing.Color.LightGreen
            };
            btnConfirm.Click += (s, e) => this.DialogResult = DialogResult.OK;

            // Set
            filterPanel.Controls.Add(txtFilter);
            this.Controls.Add(listView);
            this.Controls.Add(switchPanel);
            this.Controls.Add(filterPanel);
            this.Controls.Add(btnConfirm);
            this.Controls.Add(lblStatus);
        }

        private void SwitchSystemType(bool showDuct)
        {
            SaveCurrentSelections();

            showDuctSystems = showDuct;
            btnShowDuctSystems.BackColor = showDuct ? System.Drawing.Color.LightGreen : SystemColors.Control;
            btnShowPipeSystems.BackColor = showDuct ? SystemColors.Control : System.Drawing.Color.LightGreen;

            LoadCurrentSystems();
            ApplyFilter(txtFilter.Text);
        }

        private void LoadCurrentSystems()
        {
            currentItems = showDuctSystems ? ductSystems : pipeSystems;
            UpdateStatusText();
        }

        private void SaveCurrentSelections()
        {
            foreach (ListViewItem item in listView.Items)
            {
                selectionState[item.Text] = item.Checked;
            }
        }
        private void SetupPlaceholder()
        {
            txtFilter.Text = PlaceholderText;
            txtFilter.ForeColor = SystemColors.GrayText;
            isPlaceholderActive = true;
        }
        private void RemovePlaceholder(object sender, EventArgs e)
        {
            if (isPlaceholderActive)
            {
                txtFilter.Text = "";
                txtFilter.ForeColor = SystemColors.WindowText;
                isPlaceholderActive = false;
            }
        }

        private void SetPlaceholder(object sender, EventArgs e)
        {
            if (string.IsNullOrWhiteSpace(txtFilter.Text))
            {
                SetupPlaceholder();
            }
        }

        private async void OnFilterTextChangedAsync(object sender, EventArgs e)
        {
            if (isPlaceholderActive) return;

            filterCancellationTokenSource?.Cancel();
            filterCancellationTokenSource = new CancellationTokenSource();

            try
            {
                await Task.Delay(300, filterCancellationTokenSource.Token);
                ApplyFilter(txtFilter.Text);
            }
            catch (TaskCanceledException)
            {
                // Фильтрация была отменена новым вводом
            }
        }

        private void ApplyFilter(string filter)
        {
            if (this.InvokeRequired)
            {
                this.Invoke(new Action<string>(ApplyFilter), filter);
                return;
            }

            try
            {
                lblStatus.Text = "Фильтрация...";

                listView.BeginUpdate();
                listView.Items.Clear();

                var filtered = string.IsNullOrWhiteSpace(filter) || isPlaceholderActive 
                    ? currentItems 
                    : currentItems.Where(item => item.IndexOf(filter, StringComparison.OrdinalIgnoreCase) >= 0)
                    .ToList();

                foreach (var item in filtered.OrderBy(x => x))
                {
                    var listItem = new ListViewItem(item)
                    {
                        Checked = selectionState.TryGetValue(item, out bool isChecked) && isChecked
                    };
                    listView.Items.Add(listItem);
                }

                UpdateStatusText();
            }
            finally
            {
                listView.EndUpdate();
            }
        }

        private void OnItemChecked(object sender, ItemCheckedEventArgs e)
        {
            if (e.Item == null) return;
            selectionState[e.Item.Text] = e.Item.Checked;
        }
        private void UpdateStatusText()
        {
            int total = showDuctSystems ? ductSystems.Count : pipeSystems.Count;
            int selected = selectionState.Count(kv => kv.Value && (showDuctSystems ? ductSystems.Contains(kv.Key) : pipeSystems.Contains(kv.Key)));

            lblStatus.Text = $"Выбрано: {selected} | Показано: {listView.Items.Count} из {total}";
        }

        protected override void OnFormClosing(FormClosingEventArgs e)
        {
            filterCancellationTokenSource?.Cancel();
            base.OnFormClosing(e);
        }
    }
}