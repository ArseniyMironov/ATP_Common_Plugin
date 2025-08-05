using ATP_Common_Plugin.Services;
using ATP_Common_Plugin.Utils;
using Autodesk.Revit.Attributes;
using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using System;
using System.Collections.Generic;
using System.Linq;

namespace ATP_Common_Plugin.Commands
{
    [Regeneration(RegenerationOption.Manual)]
    [Transaction(TransactionMode.Manual)]

    class MarkDuctAccesories : IExternalCommand
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
                // Получаем все Duct Accesory
                IList<Element> ductAcessories = selecttionBuiltInInstance.selectInstanceOfCategory(doc, BuiltInCategory.OST_DuctAccessory);

                if (ductAcessories.Count == 0)
                {
                    //TaskDialog.Show("Ошибка", "В модели нет элементов категории Duct Accesory'. (Проверьте рабочие наборы)");
                    logger.LogWarning("В модели нет элементов категории Duct Accesory'. (Проверьте рабочие наборы)", docName);
                    return Result.Cancelled;
                }

                var groupedByModel = ductAcessories
                    .GroupBy(x => GetAccesoriesType(doc.GetElement(x.GetTypeId()).get_Parameter(BuiltInParameter.ALL_MODEL_MODEL).AsValueString()))
                    .ToDictionary(g => g.Key, g => g.ToList());

                logger.LogInfo("Начато выполнение Mark Duct Accesory", docName);
                using (Transaction tr = new Transaction(doc, "Маркировка Duct Accessories"))
                {
                    tr.Start();

                    foreach (var typeGroup in groupedByModel)
                    {
                        string accessoryType = typeGroup.Key;
                        IList<Element> accessories = typeGroup.Value;

                        string prefix = dictionaryHvacElements.GetPrefix(accessoryType);

                        var groupedByMark = accessories
                            .GroupBy(x => RevitUtils.GetSharedParameterValue(x, dictionaryGUID.ADSKMark))
                            .OrderBy(g => g.Key)
                            .ToList();

                        int groupNumber = 1;

                        foreach (var markGroup in groupedByMark)
                        {
                            if (prefix.Contains("ППК."))
                            {
                                continue;
                            }

                            string fullMark = $"{groupNumber}.{prefix}";

                            foreach (var accessory in markGroup)
                            {
                                RevitUtils.SetParameterValue(accessory, dictionaryGUID.ADSKSign, fullMark);
                            }
                            groupNumber++;
                        }
                    }

                    tr.Commit();
                }

                //TaskDialog.Show("Готово", "Маркировка воздуховодной арматуры выполнена!");
                logger.LogInfo("Маркировка воздуховодной арматуры выполнена.", docName);
                return Result.Succeeded;
            }
            catch (Exception ex)
            {
                //TaskDialog.Show("Ошибка", ex.ToString());
                logger.LogError($"Ошибка! {ex.Message}", docName);
                return Result.Failed;
            }
        }

        // Определяем тип клапана по параметру Model
        private string GetAccesoriesType(string familyModel)
        {
            if (string.IsNullOrEmpty(familyModel))
                return "Неизвестный";
            if (familyModel.IndexOf("Зонт", StringComparison.OrdinalIgnoreCase) >= 0)
                return "Зонт прямоугольный";
            if (familyModel.IndexOf("Лючок для замеров", StringComparison.OrdinalIgnoreCase) >= 0)
                return "Лючок для замеров";
            if (familyModel.IndexOf("Лючок для прочистки воздуховода", StringComparison.OrdinalIgnoreCase) >= 0)
                return "Лючок для прочистки воздуховода";
            if (familyModel.IndexOf("_421_", StringComparison.OrdinalIgnoreCase) >= 0)
                return "Дефлектор";
            if (familyModel.IndexOf("_430_", StringComparison.OrdinalIgnoreCase) >= 0)
                return "Фильтр воздушный канальный";
            if (familyModel.IndexOf("_442-1_", StringComparison.OrdinalIgnoreCase) >= 0)
                return "Клапан противопожарный нормально открытый";
            if (familyModel.IndexOf("_442-2_", StringComparison.OrdinalIgnoreCase) >= 0)
                return "Клапан противопожарный нормально закрытый";
            if (familyModel.IndexOf("_442-3_", StringComparison.OrdinalIgnoreCase) >= 0)
                return "Клапан противопожарный двойного действия";
            if (familyModel.IndexOf("_442-4_", StringComparison.OrdinalIgnoreCase) >= 0)
                return "Дроссель-клапан воздушный";
            if (familyModel.IndexOf("_442-5_", StringComparison.OrdinalIgnoreCase) >= 0)
                return "Клапан регулирующий";
            if (familyModel.IndexOf("_442-6_", StringComparison.OrdinalIgnoreCase) >= 0)
                return "Клапан обратный воздушный";
            if (familyModel.IndexOf("_442-7_", StringComparison.OrdinalIgnoreCase) >= 0)
                return "Клапан сброса избыточного давления";
            if (familyModel.IndexOf("_429_", StringComparison.OrdinalIgnoreCase) >= 0)
                return "Воздухонагреватель канальный";
            if (familyModel.IndexOf("_431_", StringComparison.OrdinalIgnoreCase) >= 0)
                return "Воздухоохладитель канальный";
            if (familyModel.IndexOf("_424_", StringComparison.OrdinalIgnoreCase) >= 0)
                return "Шумоглушитель";

            return "Не маркируется";
        }
    }
}
