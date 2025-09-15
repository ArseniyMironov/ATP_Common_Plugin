using System;
using System.Collections.Generic;
using System.Linq;
using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using Autodesk.Revit.Attributes;
using ATP_Common_Plugin.Utils;
using Autodesk.Revit.ApplicationServices;
using System.IO;
using ATP_Common_Plugin.Services;

namespace ATP_Common_Plugin.Commands
{
    [Transaction(TransactionMode.Manual)]
    public class FilterAirTerminalsCommand : IExternalCommand
    {
        public Result Execute(ExternalCommandData commandData, ref string message, ElementSet elements)
        {
            UIDocument uiDoc = commandData.Application.ActiveUIDocument;
            Document doc = uiDoc.Document;
            Application app = doc.Application;
            string docName = doc.Title;
            var logger = ATP_App.GetService<ILoggerService>();
            var BuiltInCat = BuiltInCategory.OST_DuctTerminal;

            try
            {
                // 1. Получаем все воздухораспределители
                var airTerminals = new FilteredElementCollector(doc)
                    .OfCategory(BuiltInCat)
                    .WhereElementIsNotElementType()
                    .Cast<FamilyInstance>()
                    .ToList();

                if (airTerminals.Count == 0)
                {
                    logger.LogError("Не найдено воздухораспределителей", docName);
                    return Result.Failed;
                }

                // 2. Проверяем параметр
                string paramName = "ATP_Основа";
                Parameter testParam = airTerminals[0].LookupParameter(paramName);
                if (testParam == null)
                {
                    logger.LogWarning($"Параметр '{paramName}' не найден", docName);

                    try
                    {
                        RevitUtils.AddSharedParameter(doc, paramName, dictionaryGUID.ATPHost, BuiltInCat);
                        logger.LogInfo("Добавлен параметр ATP_Основа", docName);
                    }
                    catch (Exception ex)
                    {
                        logger.LogError($"Ошибка при добавлении общего параметра: {ex.Message}", docName);
                        return Result.Failed;
                    }
                }

                // 3. Получаем связанные модели
                var archLinks = new FilteredElementCollector(doc)
                    .OfClass(typeof(RevitLinkInstance))
                    .Cast<RevitLinkInstance>()
                    .Where(x => x.Name.Contains("_AR_") || x.Name.Contains("_Arch"))
                    .ToList();

                if (archLinks.Count == 0)
                {
                    logger.LogError("Не найдено связанных моделей архитектуры", docName);
                    return Result.Failed;
                }

                // 4. Обработка в транзакции
                using (Transaction t = new Transaction(doc, "Заполнение параметров"))
                {
                    t.Start();

                    int markedCount = 0;
                    int clearedCount = 0;
                    int totalCeilingsChecked = 0;

                    // Очищаем параметры
                    foreach (var terminal in airTerminals)
                    {
                        Parameter param = terminal.LookupParameter(paramName);
                        if (param == null || param.IsReadOnly) continue;
                        param.Set("");
                        clearedCount++;
                    }

                    // Проверяем пересечения
                    foreach (var link in archLinks)
                    {
                        var linkDoc = link.GetLinkDocument();
                        if (linkDoc == null) continue;

                        // Получаем подвесные И реечные потолки
                        var ceilings = new FilteredElementCollector(linkDoc)
                            .OfCategory(BuiltInCategory.OST_Ceilings)
                            .WhereElementIsNotElementType()
                            .Where(c => IsCeilingTypeMatch(c))
                            .ToList();

                        var suspendedCeilings = new List<Ceiling>();

                        totalCeilingsChecked += ceilings.Count;

                        if (ceilings.Count == 0)
                        {
                            logger.LogError($"В модели {link.Name} нет подходящих потолков", docName);
                            continue;
                        }

                        foreach (var terminal in airTerminals)
                        {
                            Parameter param = terminal.LookupParameter(paramName);
                            if (param == null || param.IsReadOnly || param.AsString() == "да")
                                continue;

                            if (CheckIntersection(terminal, ceilings, link.GetTotalTransform()))
                            {
                                param.Set("да");
                                markedCount++;
                            }
                        }
                    }

                    t.Commit();

                    logger.LogInfo($"Обработано моделей: {archLinks.Count}\n", docName);
                    logger.LogInfo($"Проверено потолков: {totalCeilingsChecked}", docName);
                    logger.LogInfo($"Найдено пересечений: {markedCount}", docName);
                }

                return Result.Succeeded;
            }
            catch (Exception ex)
            {
                logger.LogError(ex.Message, docName);
                return Result.Failed;
            }
        }

        // Проверяем тип потолка (подвесной или реечный)
        private bool IsCeilingTypeMatch(Element ceiling)
        {
            string typeName = GetCeilingTypeName(ceiling).ToLower();
            return typeName.Contains("подвесной") || typeName.Contains("реечный");
        }

        private string GetCeilingTypeName(Element ceiling)
        {
            ElementId typeId = ceiling.GetTypeId();
            if (typeId == null) return "";

            Element ceilingType = ceiling.Document.GetElement(typeId);
            return ceilingType?.Name ?? "";
        }

        private bool CheckIntersection(FamilyInstance terminal, List<Element> ceilings, Transform linkTransform)
        {
            BoundingBoxXYZ bbTerminal = terminal.get_BoundingBox(null);
            if (bbTerminal == null) return false;

            foreach (Element ceiling in ceilings)
            {
                BoundingBoxXYZ bbCeiling = ceiling.get_BoundingBox(null);
                if (bbCeiling == null) continue;

                XYZ min = linkTransform.OfPoint(bbCeiling.Min);
                XYZ max = linkTransform.OfPoint(bbCeiling.Max);

                if (bbTerminal.Min.X < max.X && bbTerminal.Max.X > min.X &&
                    bbTerminal.Min.Y < max.Y && bbTerminal.Max.Y > min.Y &&
                    bbTerminal.Min.Z < max.Z && bbTerminal.Max.Z > min.Z)
                {
                    return true;
                }
            }
            return false;
        }
    }
}