using Autodesk.Revit.Attributes;
using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using System;
using System.Collections.Generic;
using System.Linq;

namespace ATP_Common_Plugin.Commands.Tasks
{
    [Regeneration(RegenerationOption.Manual)]
    [Transaction(TransactionMode.Manual)]
    public class FireDampers : IExternalCommand
    {
        public Result Execute(ExternalCommandData commandData, ref string message, ElementSet elements)
        {
            UIDocument uiDoc = commandData.Application.ActiveUIDocument;
            Document doc = uiDoc.Document;

            var provider = new ParameterValueProvider(new ElementId(BuiltInParameter.ELEM_TYPE_PARAM));

            var dampersType = new Dictionary<string, string>()
            {
                ["Клапан противопожарный нормально открытый"] = "ППК.НО",
                ["Клапан противопожарный нормально закрытый"] = "ППК.НЗ",
                ["Клапан противопожарный двойного действия"] = "ППК.ДД",
                ["Клапан регулирующий"] = "РК"
            };

            var sortedDampersTypeId = new Dictionary<string, List<Element>>();

            IList<string> keys = new List<string>() {"Клапан противопожарный нормально открыты", "Клапан противопожарный нормально закрытый", "Клапан противопожарный двойного действи", "Клапан регулирующий" };

            IList<Element> typeIds = new FilteredElementCollector(doc)
                .WhereElementIsNotElementType()
                .OfCategory(BuiltInCategory.OST_DuctAccessory)
                .ToElements();

            Dictionary<string, List<Element>> dampers = GroupByModel(typeIds, keys);



            //foreach (Element type in typeIds)
            //{
            //    string typeModel = type.get_Parameter(BuiltInParameter.ALL_MODEL_MODEL).AsString();

            //    if (string.IsNullOrWhiteSpace(typeModel))
            //        continue;
                 
            //    foreach (string value in dampersType.Keys)
            //    {
            //        if(typeModel.IndexOf(value, StringComparison.OrdinalIgnoreCase) >= 0)
            //        {
            //            sortedDampersTypeId[value].Add(type.Id);
            //        }
            //        else
            //        {
            //            continue;
            //        }
            //    }       
            //}


            //foreach (string type in sortedDampersTypeId.Keys)
            //{
            //    string prefix = dampersType[type];

            //    IList<Element> dampers = new List<Element>() { };

            //    foreach (ElementId id in sortedDampersTypeId[type])
            //    {
            //        var rule = new FilterElementIdRule(provider, new FilterNumericEquals(), id);
            //        var typeFilter = new ElementParameterFilter(rule);

            //        var damperCollector = new FilteredElementCollector(doc).WhereElementIsNotElementType().OfClass(typeof(FamilyInstance)).WherePasses(typeFilter);

            //        foreach (FamilyInstance fi in damperCollector)
            //            dampers.Add(fi);
            //    }

            //    dampers.OrderBy(d => d.get_Parameter(Utils.dictionaryGUID.ADSKMark).AsValueString())
            //        .ThenBy(d => d.get_Parameter(BuiltInParameter.RBS_CALCULATED_SIZE).AsValueString());

            //    foreach (Element damper in dampers)
            //    {
            //        int counter = 1;
            //        string newSign = $"{prefix}.{counter}";
            //        RevitUtils.SetParameterValue(damper, Utils.dictionaryGUID.ADSKSign, newSign);
            //        counter++;
            //    }
            //}

            return Result.Succeeded;
        }

        /// <summary>
        /// Группирует элементы по ключам:
        /// если строковое значение параметра ТИПА (BuiltInParameter.ALL_MODEL_MODEL) содержит key,
        /// то элемент попадает в groups[key].
        /// Элементы без совпадений никуда не добавляются.
        ///
        /// Важно: если совпало несколько ключей, элемент попадает в ПЕРВУЮ группу по порядку groupKeys.
        /// </summary>
        public static Dictionary<string, List<Element>> GroupByModel(IList<Element> elements, IList<string> groupKeys)
        {
            if (elements == null) throw new ArgumentNullException(nameof(elements));
            if (groupKeys == null) throw new ArgumentNullException(nameof(groupKeys));

            // Инициализируем словарь групп (только по ключам)
            var groups = new Dictionary<string, List<Element>>(StringComparer.OrdinalIgnoreCase);

            for (int i = 0; i < groupKeys.Count; i++)
            {
                var key = groupKeys[i];
                if (string.IsNullOrEmpty(key)) continue;

                if (!groups.ContainsKey(key))
                    groups[key] = new List<Element>();
            }

            // Основной проход
            for (int i = 0; i < elements.Count; i++)
            {
                var el = elements[i];
                if (el == null) continue;

                // Берём тип элемента
                ElementId typeId = el.GetTypeId();
                if (typeId == ElementId.InvalidElementId) continue;

                var type = el.Document.GetElement(typeId) as ElementType;
                if (type == null) continue;

                // Получаем значение параметра Model из типа
                Parameter p = type.get_Parameter(BuiltInParameter.ALL_MODEL_MODEL);
                if (p == null || p.StorageType != StorageType.String) continue;

                string value = p.AsString();
                if (string.IsNullOrWhiteSpace(value)) continue;

                // Порядок groupKeys = приоритет (первое совпадение)
                for (int k = 0; k < groupKeys.Count; k++)
                {
                    var key = groupKeys[k];
                    if (string.IsNullOrWhiteSpace(key)) continue;

                    if (value.IndexOf(key, StringComparison.OrdinalIgnoreCase) >= 0)
                    {
                        List<Element> bucket;
                        if (!groups.TryGetValue(key, out bucket))
                        {
                            bucket = new List<Element>();
                            groups[key] = bucket;
                        }

                        bucket.Add(el);
                        break;
                    }
                }
            }

            return groups;
        }
    }
}
