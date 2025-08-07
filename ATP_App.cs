using ATP_Common_Plugin.Services;
using Autodesk.Revit.UI;
using Microsoft.Extensions.DependencyInjection;
using System;
using System.Reflection;
using System.Windows.Media.Imaging;
using PushButton = Autodesk.Revit.UI.PushButton;

namespace ATP_Common_Plugin
{
    class ATP_App : IExternalApplication
    {

        private static IServiceProvider _serviceProvider;

        public Result OnStartup(UIControlledApplication app)
        { 
            try
            {
                // Реализуем логгер
                var services = new ServiceCollection();

                services.AddSingleton<ILoggerService, LoggerService>();

                _serviceProvider = services.BuildServiceProvider();

                string assemblyName = Assembly.GetExecutingAssembly().Location;

                string commandNamespace = "ATP_Common_Plugin.Commands.";
                // Иконки
                BitmapImage numbering = new BitmapImage(new Uri("pack://application:,,,/ATP_Common_Plugin;component/Images/numbering.ico"));
                string numberingIconAuthor = "Freepik";
                BitmapImage fix = new BitmapImage(new Uri("pack://application:,,,/ATP_Common_Plugin;component/Images/FixIcon.ico"));
                string fixIconAuthor = "Freepik";
                BitmapImage SchemeIco = new BitmapImage(new Uri("pack://application:,,,/ATP_Common_Plugin;component/Images/Scheme.ico"));
                string shemeIcoAuthor = "Sympnoiaicon";
                BitmapImage errorCenterIco = new BitmapImage(new Uri("pack://application:,,,/ATP_Common_Plugin;component/Images/ErrorCenter.ico"));
                string errorCenterIcoAuthor = "Sympnoiaicon";

                // Создание вкладки
                string HKLStabName = "AT•P HKLS";
                app.CreateRibbonTab(HKLStabName);

                // HKLS
                // Создание панелей на вкладки
                RibbonPanel ribbonPamelHKLSMark = app.CreateRibbonPanel(HKLStabName, "Оформление");
                RibbonPanel ribbonPamelHKLSUtils = app.CreateRibbonPanel(HKLStabName, "Утилиты");
                RibbonPanel ribbonPanelHKLSSchedule = app.CreateRibbonPanel(HKLStabName, "Спецификация");
                RibbonPanel ribbonPanelHKLSTask = app.CreateRibbonPanel(HKLStabName, "Задания");


                // Создание кнопки для управления логгером
                var toggleBtnData = new PushButtonData(
                    name: "ToggleLogger",
                    text: "Error center",
                    assemblyName: assemblyName,
                    commandNamespace + "ToggleLoggerCommand"
                )
                {
                    LargeImage = errorCenterIco,
                    LongDescription = $"Центр ошибок собирающий всю информацию об ошибках, предупреждениях и процессах выполнения. Made by EDD & ARMI, Icon by {errorCenterIcoAuthor}"
                };
                PushButton toggleBtn = ribbonPamelHKLSUtils.AddItem(toggleBtnData) as PushButton;
                // Маркировка
                // Мариковрка элементов HKLS
                string mark_description = "Заполняет параметр ADSK_Обозначение по стандартам оформления ATP-TLP для элементов категории ";
                PushButtonData mark_air_termi_BtnData = new PushButtonData(name: "Mark Air Terminals", text: "Air Terminals", assemblyName: assemblyName, commandNamespace + "MarkAirTerm")
                {
                    LargeImage = numbering,
                    LongDescription = $"{mark_description} Air Terminal. Made by ARMI, Icon by {numberingIconAuthor}"
                };

                PushButtonData mark_duct_acc_BtnData = new PushButtonData(name: "Mark Duct Accessory", text: "Duct Accessory", assemblyName: assemblyName, commandNamespace + "MarkDuctAccesories")
                {
                    LargeImage = numbering,
                    LongDescription = $"{mark_description} Duct Accessory. Made by ARMI, Icon by {numberingIconAuthor}"
                };

                SplitButtonData splitButtonDataMarking = new SplitButtonData("Marking", "Marking MEP")
                {
                    LongDescription = "Маркировка элементов по стандартам ATP-TLP." // За подробностью о маркировке, общращайтесь к правиласм моделирования HKLS
                };
                SplitButton splitButtonMarking = ribbonPamelHKLSMark.AddItem(splitButtonDataMarking) as SplitButton;
                splitButtonMarking.AddPushButton(mark_air_termi_BtnData); 
                splitButtonMarking.AddPushButton(mark_duct_acc_BtnData);

                PushButtonData mark_opennings_BtnData = new PushButtonData(name: "Mark Opennings", text: "Mark Opennings", assemblyName: assemblyName, commandNamespace + "MarkOpennings")
                {
                    LargeImage = numbering,
                    ToolTip = "Маркировка заданий на отверстия на текущем виде",
                    LongDescription = "Первичная маркировка заданий на отверстия"
                };
                PushButton mark_opennings = ribbonPamelHKLSMark.AddItem(mark_opennings_BtnData) as PushButton;

                // Создание 3D видов
                PushButtonData create_3D_Scheme_BtnData = new PushButtonData(name: "Create 3d Scheme", text: "3D Scheme", assemblyName: assemblyName, commandNamespace + "Create3DViews")
                {
                    LargeImage = SchemeIco,
                    ToolTip = "Создание схемы для оформления",
                    LongDescription = $"Создание изометрических схем по системам. Требует заполненного параметра ИмяСистемы. Made by ARMI, Icon by {shemeIcoAuthor}"
                };
                PushButton create_3D_Scheme_Btn = ribbonPamelHKLSMark.AddItem(create_3D_Scheme_BtnData) as PushButton;

                // Корректировка изоляции в модели
                PushButtonData fix_insulation_BtnData = new PushButtonData(name: "Fix Duct and pipe insulation workset", text: "Fix Insulation", assemblyName: assemblyName, commandNamespace + "FixInsulation")
                {
                    LargeImage = fix,
                    ToolTip = "Корректировка изоляции.",
                    LongDescription = $"Корректирует рабочие наборы у изоляции, если он отличеный от рабочего набора основы и удаляет изоляцию без основы. Made by EDD & ARMI, Icon by {fixIconAuthor}"
                };
                PushButton fix_insulation_Btn = ribbonPamelHKLSUtils.AddItem(fix_insulation_BtnData) as PushButton;

                // Заполнение уклона
                var fill_slope_BtnData = new PushButtonData(name: "FillSlope", text: "Fill slope", assemblyName: assemblyName, commandNamespace + "FillSlope")
                {
                    LargeImage = fix,
                    ToolTip = "Заполнение уклона трубопроода.",
                    LongDescription = $"Заполняет параметр ADSK_Уклон у трубопроводов в модели, на основе их уклона в модели (Показанного в параметре Slope). Made by EDD & ARMI. Icon by {fixIconAuthor}"
                };
                PushButton fill_slope_Btn = ribbonPamelHKLSUtils.AddItem(fill_slope_BtnData) as PushButton;

                // Специфицирование элементов в модели
                // Заполнение группирования
                PushButtonData HVAC_Schedule_1_BtnData = new PushButtonData(name: "Fill Group", text: "Name and Group", assemblyName: assemblyName, commandNamespace + "Schedule_1_HKLS")
                {
                    LargeImage = numbering,
                    ToolTip = "Заполнение группирования элементов.",
                    LongDescription = $"Заполняет параметры ADSK_Группирование и ИмяСистемы. Обрабатывает категории Plumbing Fixtures, Mechanical Equipment, Air Terminals, Duct Accessory, Ducts, Duct Fitting, Flex Ducts, Duct Insulationm, Pipe Accessory< Pipes, Pipe Fitting, Flex Pipes и Pipe Insulation. Made by ARMI, Icon by {numberingIconAuthor}"
                };
                PushButton HVAC_Schedule_1_Btn = ribbonPanelHKLSSchedule.AddItem(HVAC_Schedule_1_BtnData) as PushButton;

                // Заполнение вложенных
                PushButtonData HVAC_Schedule_2_BtnData = new PushButtonData(name: "Fill Subelements", text: "Subelements", assemblyName: assemblyName, commandNamespace + "Schedule_2_HKLS")
                {
                    LargeImage = numbering,
                    ToolTip = "Обработка вложенных элементов.",
                    LongDescription = $"Заполняет параметры ADSK_Группирование, ИмяСистемы и ADSK_Комплект во все вложенные элементы, а так же задаёт новое значение для основного элемента для комплектной сборки в спецификации. Обрабатывает категории Plumbing Fixtures, Mechanical Equipment, Duct Accessory, Duct Fitting, Pipe Accessory и Pipe Fitting. Made by ARMI, Icon by {numberingIconAuthor}"
                };
                PushButton HVAC_Schedule_2_Btn = ribbonPanelHKLSSchedule.AddItem(HVAC_Schedule_2_BtnData) as PushButton;

                // Заполнение линейных
                PushButtonData HVAC_Schedule_3_BtnData = new PushButtonData(name: "Fill element", text: "Linear Elements", assemblyName: assemblyName, commandNamespace + "Schedule_3_HKLS")
                {
                    LargeImage = numbering,
                    ToolTip = "Заполнение труб и воздуховодов.",
                    LongDescription = $"Заполнение значений параметров ADSK_Наименование, ADSK_Колличество, ADSK_Обозначение и ADSK_Толщина стенки. Обрабатывает категории Duct, Flex Duct, Duct Fitting, Duct Insulation, Pipe, Flex Pipe и Pipe Insulation. Made by ARMI, Icon by {numberingIconAuthor}"
                };
                PushButton Hvac_Schedule_3_Btn = ribbonPanelHKLSSchedule.AddItem(HVAC_Schedule_3_BtnData) as PushButton;
                
                // Task
                // Задание на решетки для АР
                PushButtonData HVAC_AirTerminal_Task_BtnData = new PushButtonData(name: "Task_Mark_AirTerm", text: "Task air terminal in ceilings", assemblyName: assemblyName, commandNamespace + "FilterAirTerminalsCommand")
                {
                    LargeImage = numbering,
                    ToolTip = "Заполнение основы воздухораспределителей.",
                    LongDescription = $"Заполнение параметра ATP_Основа в воздухораспределителях необходимого для задания АР. Made by SHKA & ARMI, Icon by {numberingIconAuthor}" 
                };
                PushButton HVAC_AirTerminal_Task_Btn = ribbonPanelHKLSTask.AddItem(HVAC_AirTerminal_Task_BtnData) as PushButton;

                return Result.Succeeded;

            }
            catch (Exception ex)
            {
                TaskDialog.Show("Ошибка инициализации", ex.ToString());
                return Result.Failed;
            }
        }
        public Result OnShutdown(UIControlledApplication app)
        {
            // Очистка ресурсов
            if (_serviceProvider is IDisposable disposable)
            {
                disposable.Dispose();
            }

            return Result.Succeeded;
        }
        public static T GetService<T>() => _serviceProvider.GetService<T>();
    }
}