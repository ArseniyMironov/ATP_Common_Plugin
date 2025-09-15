using ATP_Common_Plugin.Utils;
using Autodesk.Revit.Attributes;
using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using Microsoft.Office.Interop.Excel;
using System;
using System.Collections.Generic;
using System.Linq;

namespace ATP_Common_Plugin.Commands
{
    [Regeneration(RegenerationOption.Manual)]
    [Transaction(TransactionMode.Manual)]
    class MarkPlumbFix : IExternalCommand
    {
        private const string SearchToken = "Трап";  // что ищем в параметре Model
        private const string Prefix = "ВВ";        // префикс перед аббревиатурой системы

        public Result Execute(ExternalCommandData commandData, ref string message, ElementSet elements)
        {
            UIDocument uiDoc = commandData.Application.ActiveUIDocument;
            if (uiDoc == null) { message = "Нет активного документа."; return Result.Failed; }
            Document doc = uiDoc.Document;

            // 1) Сбор элементов категории Plumbing Fixtures
            var col = new FilteredElementCollector(doc)
                .OfCategory(BuiltInCategory.OST_PlumbingFixtures)
                .WhereElementIsNotElementType();

            // 2) Фильтр по встроенному параметру "Model" (ALL_MODEL_MODEL) содержит "710"
            //    Используем параметр-фильтр по подстроке (без учета регистра).
            var pvp = new ParameterValueProvider(new ElementId(BuiltInParameter.ALL_MODEL_MODEL));
            var contains = new FilterStringContains();
            var rule = new FilterStringRule(pvp, contains, SearchToken, false);
            var filter = new ElementParameterFilter(rule);

            IList<Element> candidates = col
                .WherePasses(filter)
                .Where(e => (e as FamilyInstance)?.SuperComponent == null)
                .ToList();

            if (candidates.Count == 0)
            {
                TaskDialog.Show("Нумерация", "Не найдено сантехприборов с Model, содержащим \"" + SearchToken + "\".");
                return Result.Succeeded;
            }

            // 3) Подготовим набор с координатами и аббревиатурой системы
            //    (минимизируем повторы обращений к параметрам/геометрии).
            var items = new List<TagItem>(capacity: candidates.Count);
            foreach (var e in candidates)
            {
                var fi = e as FamilyInstance;
                if (fi == null) continue;
                if (fi.SuperComponent != null) continue;

                XYZ p = GetStablePoint(e);
                if (p == null) continue;

                string abbr = GetSystemAbbreviation(fi, doc);
                items.Add(new TagItem
                {
                    Instance = fi,
                    SystemAbbr = string.IsNullOrWhiteSpace(abbr) ? "NA" : abbr.Trim(),
                    P = p
                });
            }

            if (items.Count == 0)
            {
                TaskDialog.Show("Нумерация", "Подходящие элементы не имеют валидной геометрии (точки).");
                return Result.Succeeded;
            }

            // 4) Группировка по System Abbreviation
            var groups = items.GroupBy(i => i.SystemAbbr);

            int totalChanged = 0;

            using (var t = new Transaction(doc, "ВВ: нумерация сантехприборов"))
            {
                t.Start();

                foreach (var g in groups)
                {
                    // 5) Сортировка: Z ↑, затем X ↑, затем Y ↑
                    var ordered = g.OrderByDescending(i => i.P.Z)
                                   .ThenBy(i => i.P.X)
                                   .ThenBy(i => i.P.Y);

                    int index = 1; // порядковый номер внутри группы
                    foreach (var it in ordered)
                    {
                        string value = $"{Prefix}-{g.Key}-{index}";
                        var markParam = it.Instance.get_Parameter(dictionaryGUID.ATPMarkScriot);
                        if (markParam != null && !markParam.IsReadOnly)
                        {
                            markParam.Set(value);
                            totalChanged++;
                        }
                        index++;
                    }
                }

                t.Commit();
            }

            TaskDialog.Show("Нумерация", $"Обработано: {items.Count}\nИзменено Mark: {totalChanged}");
            return Result.Succeeded;
        }

        /// <summary>
        /// Возвращает аббревиатуру системы (System Abbreviation) для FamilyInstance.
        /// Приоритет: значение на системе → значение на типе системы → имя системы → "NA".
        /// </summary>
        private static string GetSystemAbbreviation(FamilyInstance fi, Document doc)
        {
            // Для сантехприборов разъёмы доступны через MEPModel.ConnectorManager.
            var mep = fi.MEPModel;
            var cm = mep?.ConnectorManager;
            if (cm == null) return "NA";

            foreach (Connector c in cm.Connectors)
            {
                var sys = c.MEPSystem;
                if (sys == null) continue;

                // Пытаемся прочитать встроенный параметр аббревиатуры системы
                string abbr = ReadString(sys, BuiltInParameter.RBS_DUCT_PIPE_SYSTEM_ABBREVIATION_PARAM);
                if (!string.IsNullOrEmpty(abbr)) return abbr;

                // Если на экземпляре системы пусто — пробуем на типе системы
                Element sysType = sys.Document.GetElement(sys.GetTypeId());
                if (sysType != null)
                {
                    abbr = ReadString(sysType, BuiltInParameter.RBS_DUCT_PIPE_SYSTEM_ABBREVIATION_PARAM);
                    if (!string.IsNullOrEmpty(abbr)) return abbr;
                }

                // Фолбэк — имя системы
                if (!string.IsNullOrWhiteSpace(sys.Name)) return sys.Name;
            }

            return "NA";
        }

        /// <summary>
        /// Безопасное чтение строкового параметра по BuiltInParameter.
        /// </summary>
        private static string ReadString(Element e, BuiltInParameter bip)
        {
            if (e == null) return null;
            var p = e.get_Parameter(bip);
            return p != null ? (p.AsString() ?? null) : null;
        }

        /// <summary>
        /// Получить устойчивую точку для сортировки:
        /// LocationPoint.Point, иначе центр BoundingBoxXYZ в активном виде/документе.
        /// </summary>
        private static XYZ GetStablePoint(Element e)
        {
            var lp = e.Location as LocationPoint;
            if (lp != null && lp.Point != null) return lp.Point;

            // Фолбэк — центр бокса в модели
            var bb = e.get_BoundingBox(null);
            if (bb == null) return null;
            return (bb.Min + bb.Max) * 0.5;
        }

        private class TagItem
        {
            public FamilyInstance Instance;
            public string SystemAbbr;
            public XYZ P;
        }
    }
}
