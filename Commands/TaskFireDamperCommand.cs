using System;
using System.Collections.Generic;
using System.Linq;
using System.Text.RegularExpressions;
using Autodesk.Revit.Attributes;
using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using ATP_Common_Plugin.Utils;

namespace ATP_Common_Plugin.Commands
{
    [Transaction(TransactionMode.Manual)]
    public class TaskFireDamperCommand : IExternalCommand
    {
        private static readonly Guid GuidMark = dictionaryGUID.ATPMarkScriot; 
        private static readonly Guid GuidDesignation = dictionaryGUID.ADSKSign;

        private readonly Guid[] ElecParamGuids = 
        {
            dictionaryGUID.ADSKVoltage,
            dictionaryGUID.ADSKNominalPower,
            dictionaryGUID.ADSKPhaseCount
        };

        private const string DEFAULT_SYSTEM = "ПЕР";

        private readonly Dictionary<string, string> SEARCH_DICTIONARY = new Dictionary<string, string>
        {
            { "Клапан противопожарный нормально открытый", "ППК.НО" },
            { "Клапан противопожарный нормально закрытый", "ППК.НЗ" },
            { "Клапан противопожарный двойного действия", "ППК.ДД" },
            { "Клапан регулирующий", "РК" },
            { "Клапан", "ВК" }
        };

        public Result Execute(ExternalCommandData commandData, ref string message, ElementSet elements)
        {
            Document doc = commandData.Application.ActiveUIDocument.Document;

            var excludeIds = new FilteredWorksetCollector(doc)
                .OfKind(WorksetKind.UserWorkset)
                .Where(ws => ws.Name.Contains("020_Временные элементы") || ws.Name.StartsWith("000_"))
                .Select(ws => ws.Id.IntegerValue)
                .ToList();

            ElementMulticategoryFilter catFilter = new ElementMulticategoryFilter(new List<BuiltInCategory>
            {
                BuiltInCategory.OST_DuctAccessory,
                BuiltInCategory.OST_DuctTerminal
            });

            var collection = new FilteredElementCollector(doc)
                .WherePasses(catFilter)
                .WhereElementIsNotElementType()
                .Where(e => !excludeIds.Contains(e.WorksetId.IntegerValue))
                .ToList();

            using (Transaction tr = new Transaction(doc, "ATP: Маркировка клапанов"))
            {
                tr.Start();

                var allProcessedElements = new List<ElementData>();

                foreach (var e in collection)
                {
                    if (e.Category.Id.IntegerValue == (int)BuiltInCategory.OST_DuctAccessory && !HasElecParams(e, doc))
                        continue;

                    var data = new ElementData(e, doc, this);
                    if (data.Prefix == null) continue;

                    allProcessedElements.Add(data);
                }

                // ==========================================
                // ЭТАП 1: Заполнение ATP_Маркировка_Скрипт
                // Логика: Уникальный номер экземпляра внутри СИСТЕМЫ
                // ==========================================

                var systemGroups = allProcessedElements
                    .GroupBy(x => new { x.SystemName, x.Prefix });

                foreach (var group in systemGroups)
                {
                    int nextNumMark = 1;

                    var usedMarks = group
                        .Select(d => ParseNum(d.CurrentMark))
                        .Where(n => n.HasValue)
                        .Select(n => n.Value)
                        .ToList();

                    if (usedMarks.Any()) nextNumMark = usedMarks.Max() + 1;

                    foreach (var item in group.OrderBy(i => i.Id)) 
                    {
                        int num = ParseNum(item.CurrentMark) ?? nextNumMark++;
                        string mark = $"{item.SystemName}-{item.Prefix}-{num}";
                        item.El.get_Parameter(GuidMark)?.Set(mark);
                    }
                }

                // ==========================================
                // ЭТАП 2: Заполнение ADSK_Обозначение
                // Логика: Одинаковый номер для одинаковых ТИПОВ (независимо от системы)
                // ==========================================

                var prefixGroups = allProcessedElements.GroupBy(x => x.Prefix);

                foreach (var prefGroup in prefixGroups)
                {
                    var uniqueTypes = prefGroup
                        .GroupBy(x => x.TypeId)
                        .OrderBy(g => doc.GetElement(g.Key).Name);

                    int typeCounter = 1;
                    foreach (var typeGroup in uniqueTypes)
                    {
                        int currentTypeNum = typeCounter++;
                        foreach (var item in typeGroup)
                        {
                            string designationValue;
                            if (item.IsFireDamper)
                            {
                                // Пример: ППК.НО1
                                designationValue = $"ППК.{item.Prefix}{currentTypeNum}";
                            }
                            else
                            {
                                // Пример: ВР1
                                designationValue = $"{item.Prefix}.{currentTypeNum}";
                            }

                            item.El.get_Parameter(GuidDesignation)?.Set(designationValue);
                        }
                    }
                }

                tr.Commit();
            }

            return Result.Succeeded;
        }

        private bool HasElecParams(Element e, Document doc)
        {
            Element t = doc.GetElement(e.GetTypeId());
            return ElecParamGuids.All(guid => (e.get_Parameter(guid) ?? t?.get_Parameter(guid)) != null);
        }

        private int? ParseNum(string s)
        {
            if (string.IsNullOrEmpty(s)) return null;
            Match m = Regex.Match(s, @"-(\d+)$");
            return m.Success ? int.Parse(m.Groups[1].Value) : (int?)null;
        }

        private class ElementData
        {
            public Element El { get; }
            public int Id { get; }
            public ElementId TypeId { get; }
            public string SystemName { get; }
            public string Prefix { get; }
            public string CurrentMark { get; }
            public bool IsFireDamper { get; }

            public ElementData(Element e, Document doc, TaskFireDamperCommand cmd)
            {
                El = e;
                Id = e.Id.IntegerValue;
                TypeId = e.GetTypeId();
                CurrentMark = e.get_Parameter(GuidMark)?.AsString() ?? "";

                string sys = e.get_Parameter(BuiltInParameter.RBS_SYSTEM_NAME_PARAM)?.AsString() ?? "";
                int lastIndex = sys.LastIndexOf('.');
                string sysResult = lastIndex == -1 ? sys : sys.Substring(0, lastIndex);
                SystemName = string.IsNullOrWhiteSpace(sys) ? DEFAULT_SYSTEM : sysResult;

                Element typeElem = doc.GetElement(TypeId);
                string modelVal = typeElem?.get_Parameter(BuiltInParameter.ALL_MODEL_MODEL)?.AsString();

                if (string.IsNullOrWhiteSpace(modelVal))
                {
                    Prefix = null;
                }
                else
                {
                    var match = cmd.SEARCH_DICTIONARY
                        .OrderByDescending(kv => kv.Key.Length)
                        .FirstOrDefault(kv => modelVal.IndexOf(kv.Key, StringComparison.OrdinalIgnoreCase) >= 0);

                    Prefix = match.Key != null ? match.Value : null;
                }
            }
        }
    }
}
