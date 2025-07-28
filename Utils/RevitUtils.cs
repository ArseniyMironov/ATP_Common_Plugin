using ATP_Common_Plugin.Services;
using Autodesk.Revit.DB;
using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;
using System.Text.RegularExpressions;

namespace ATP_Common_Plugin
{
    public static class RevitUtils
    {

        /// <summary>
        /// Фильтрует элементы, исключая элементы из запрещённых рабочих наборов.
        /// </summary>
        public static IList<Element> FilterTempFamilyInstance(Document doc , IList<Element> elements)
        {
            IList<Workset> worksets = new FilteredWorksetCollector(doc)
                .OfKind(WorksetKind.UserWorkset)
                .ToWorksets();

            IList<WorksetId> wsBlackList = worksets
                .Where(ws => ws.Name.Contains("000_") ||
                            ws.Name.Contains("010_") ||
                            ws.Name.Contains("020_") ||
                            ws.Name.Contains("030_") ||
                            ws.Name.Contains("040_") ||
                            ws.Name.Contains("050_"))
                .Select(ws => ws.Id)
                .ToList();

            return elements
                .Where(elem => !wsBlackList.Contains(elem.WorksetId))
                .ToList();
        }

        /// <summary>
        /// Устанавливает строковое значение параметра, если он доступен и отличается от текущего значения.
        /// </summary>
        public static void SetParameterValue(Element element, Guid paramGuid, string value)
        {
            var logger = ATP_App.GetService<ILoggerService>();

            try
            {
                Parameter param = element.get_Parameter(paramGuid);
                bool isElemInGroup = element.GroupId.IntegerValue != -1;

                if (param == null && (!param.IsReadOnly))
                {
                    //TaskDialog.Show("Ошибка", $"Параметр {param.Name} отсутствует у элемента {elemetn.Id}"); // ПЕРЕНЕСТИ В ЛОГЕР
                    logger.LogWarning($"Параметр {param} отсутствует у элемента {element.Id}");
                    return;
                }

                if (isElemInGroup)
                {
                    logger.LogWarning($"Элемент {element.Id} в группе");
                    return;
                }

                string oldValue = element.get_Parameter(paramGuid).AsValueString();

                if (!param.IsReadOnly && param.StorageType == StorageType.String && value != oldValue)
                {
                    param.Set(value);
                    return;
                }
            }
            catch
            {
                return;
            }
        }
        public static void SetParameterValue(Element element, string paramName, string value)
        {
            Parameter param = element.LookupParameter(paramName);
            bool isElemInGroup = element.GroupId.IntegerValue != -1;
            var logger = ATP_App.GetService<ILoggerService>();

            if (param == null && (!param.IsReadOnly))
            {
                //TaskDialog.Show("Ошибка", $"Параметр {param.Name} отсутствует у элемента {elemetn.Id}");
                logger.LogWarning($"Параметр {param} отсутствует у элемента {element.Id}");
                return;
            }

            if (isElemInGroup)
            {
                logger.LogWarning($"Элемент {element.Id} в группе");
                return;
            }

            string oldValue = param.AsValueString();

            if (!param.IsReadOnly && param.StorageType == StorageType.String && value != oldValue)
            {
                param.Set(value);
            }
        }
        public static void SetParameterValue(Element element, Guid paramGuid, double value)
        {
            Parameter param = element.get_Parameter(paramGuid);
            bool isElemInGroup = element.GroupId.IntegerValue != -1;
            var logger = ATP_App.GetService<ILoggerService>();

            if (param == null && (!param.IsReadOnly))
            {
                //TaskDialog.Show("Ошибка", $"Параметр {param.Name} отсутствует у элемента {elemetn.Id}");
                logger.LogWarning($"Параметр {param} отсутствует у элемента {element.Id}");
                return;
            }


            double oldValue = element.get_Parameter(paramGuid).AsDouble();

            if (!param.IsReadOnly && value != oldValue) 
            {
                param.Set(value);
                return;
            }
        }

        /// <summary>
        /// Получаем значение параметра (безопасно)
        /// </summary>
        public static string GetSharedParameterValue(Element element, Guid paramGuid)
        {
            Parameter param = element.get_Parameter(paramGuid);
            string value = param?.AsValueString();

            if (!string.IsNullOrEmpty(value))
                return value;

            Element elementType = element.Document.GetElement(element.GetTypeId());
            if (elementType != null)
            {
                Parameter typeParam = elementType.get_Parameter(paramGuid);
                return typeParam?.AsValueString();
            }

            return $"Проверьте наличие параметра {param}";
        }

        public static string GetProjectParameterValue(Element element, string paramName)
        {
            Parameter param = element.LookupParameter(paramName);
            string value = param?.AsValueString();

            if (!string.IsNullOrEmpty(value))
                return value;

            Element elementType = element.Document.GetElement(element.GetTypeId());
            if (elementType != null)
            {
                Parameter typeParam = elementType.LookupParameter(paramName);
                return typeParam?.AsValueString();
            }

            return $"Проверьте наличие параметра {paramName}";
        }

        public static void CheckElement(Element elem, ref string exeption)
        {
            bool isGroup = elem.GroupId.IntegerValue != -1;
            bool isNull = elem == null;
            return ;
        }

        /// <summary>
        /// Вспомогательный метод для того, что юы убрать единицы измерения из значения типа string
        /// </summary>
        /// <param name="input"></param>
        /// <returns></returns>
        public static string CleanSizeString(string input)
        {
            if (string.IsNullOrWhiteSpace(input))
                return "";

            // Оставляем только цифры и точку
            string cleaned = Regex.Replace(input, @"[^\d.]", "");

            if (Regex.IsMatch(input, @"[xх*×]")) // Если есть разделитель
            {
                var numbers = Regex.Split(input, @"[^\d.]+")
                                  .Where(s => !string.IsNullOrEmpty(s))
                                  .Select(s => double.Parse(s, CultureInfo.InvariantCulture))
                                  .ToList();

                return numbers.Count >= 2
                    ? Math.Max(numbers[0], numbers[1]).ToString(CultureInfo.InvariantCulture)
                    : "";
            }

            return cleaned; // Если диаметр
        }

        /// <summary>
        /// WIP Вспомогательный метод для проверки элемента 
        /// </summary>
        /// <param name="elem"></param>
        /// <returns></returns>
        public static bool CheckElement(Element elem)
        {
            bool isElementInGroup = elem.GroupId.IntegerValue != -1;

            return isElementInGroup;
        }
    }
}
