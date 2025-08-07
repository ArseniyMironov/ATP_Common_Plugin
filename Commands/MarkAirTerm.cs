using System;
using System.Collections.Generic;
using System.Linq;
using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using Autodesk.Revit.Attributes;
using ATP_Common_Plugin.Utils;
using ATP_Common_Plugin.Services;


namespace ATP_Common_Plugin.Commands
{
    [Regeneration(RegenerationOption.Manual)]
    [Transaction(TransactionMode.Manual)]
    public class MarkAirTerm : IExternalCommand
    {

        public Result Execute(ExternalCommandData commandData, ref string message, ElementSet elements)
        {
            UIApplication uiApp = commandData.Application;
            UIDocument uiDoc = uiApp.ActiveUIDocument;
            Document doc = uiDoc.Document;
            string docName = doc.Title;

            var logger = ATP_App.GetService<ILoggerService>();

            try
            {
                // Получаем все Air Terminals
                IList<Element> airTerminals = selecttionBuiltInInstance.selectInstanceOfCategory(doc, BuiltInCategory.OST_DuctTerminal);

                if (airTerminals.Count == 0)
                {
                    //TaskDialog.Show("Ошибка", "В модели нет элементов категории Air Terminal. (Проверьте рабочие наборы)");
                    logger.LogWarning("В модели нет элементов категории Air Terminal. (Проверьте рабочие наборы)", docName);
                    return Result.Cancelled;
                }

                // Группируем по типу (Приточный, Вытяжной, Заборный, Выбросной)
                var groupedByType = airTerminals
                    .GroupBy(x => GetTerminalType(x.get_Parameter(BuiltInParameter.ELEM_FAMILY_PARAM).AsValueString()))
                    .ToDictionary(g => g.Key, g => g.ToList());

                logger.LogInfo("Начало маркировки воздухораспределителей", docName);

                // Маркируем элементы
                using (Transaction tr = new Transaction(doc, "Маркировка Air Terminals"))
                {
                    tr.Start();

                    foreach (var typeGroup in groupedByType)
                    {
                        string terminalType = typeGroup.Key;
                        List<Element> terminals = typeGroup.Value;

                        // Получаем префикс для текущего типа
                        string prefix = dictionaryHvacElements.GetPrefix(terminalType);

                        // Группируем по ADSK_Марка и сортируем по алфавиту
                        var groupedByMark = terminals
                            .GroupBy(x => RevitUtils.GetSharedParameterValue(x, dictionaryGUID.ADSKMark))
                            .OrderBy(g => g.Key)
                            .ToList();

                        // Создаем словарь с порядковыми номерами групп

                        int groupNumber = 1; // Начинаем нумерацию с 1

                        foreach (var markGroup in groupedByMark)
                        {
                            string fullMark = $"{groupNumber}.{prefix}";

                            foreach (var terminal in markGroup)
                            {
                                RevitUtils.SetParameterValue(terminal, dictionaryGUID.ADSKSign, fullMark);
                            }
                            groupNumber++;
                        }
                    }

                    tr.Commit();
                }

                //TaskDialog.Show("Готово", "Маркировка воздухораспределителей выполнена!");
                logger.LogInfo("Маркировка воздухораспределителей выполнена.", docName);
                return Result.Succeeded;
            }
            catch (Exception ex)
            {
                //TaskDialog.Show("Ошибка", ex.ToString());
                logger.LogError($"Ошибка! {ex.Message}", docName);
                return Result.Failed;
            }
        }

        // Определяем тип воздухораспределителя по имени семейства
        private string GetTerminalType(string familyName)
        {
            if (string.IsNullOrEmpty(familyName))
                return "Неизвестный";
            if (familyName.IndexOf("риточн", StringComparison.OrdinalIgnoreCase) >= 0)
                return "Приточный";
            if (familyName.IndexOf("ытяжн", StringComparison.OrdinalIgnoreCase) >= 0)
                return "Вытяжной";
            if (familyName.IndexOf("аборн", StringComparison.OrdinalIgnoreCase) >= 0)
                return "Заборный";
            if (familyName.IndexOf("ыбросн", StringComparison.OrdinalIgnoreCase) >= 0)
                return "Выбросной";

            return "Неизвестный";
        }
    }
}