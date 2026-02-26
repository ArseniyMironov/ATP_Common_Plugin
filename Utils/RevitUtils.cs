using ATP_Common_Plugin.Services;
using ATP_Common_Plugin.Utils;
using Autodesk.Revit.ApplicationServices;
using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
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

                if (param == null)
                {
                    logger.LogWarning($"Параметр по GUID {paramGuid} отсутствует у элемента {element.Id}", element.Document.Title);
                    return;
                }
                if (param.IsReadOnly)
                {
                    logger.LogWarning($"Параметр {param.Definition?.Name} у элемента {element.Id} только для чтения", element.Document.Title);
                    return;
                }

                if (param.StorageType == StorageType.String)
                {
                    string oldValue = param.AsString();
                    if (!string.Equals(value, oldValue))
                        param.Set(value);
                }
            }
            catch (Exception ex)
            {
                logger.LogError($"Не предвиденная ошибка! {ex.Message}");
                return;
            }
        }
        public static void SetParameterValue(Element element, string paramName, string value)
        {
            var logger = ATP_App.GetService<ILoggerService>();

            try
            {
                Parameter param = element.LookupParameter(paramName);

                if (param == null)
                {
                    logger.LogWarning($"Параметр '{paramName}' отсутствует у элемента {element.Id}", element.Document.Title);
                    return;
                }
                if (param.IsReadOnly)
                {
                    logger.LogWarning($"Параметр '{paramName}' у элемента {element.Id} только для чтения", element.Document.Title);
                    return;
                }

                if (param.StorageType == StorageType.String)
                {
                    string oldValue = param.AsString();
                    if (!string.Equals(value, oldValue))
                        param.Set(value);
                }
            }
            catch (Exception ex)
            {
                logger.LogError($"Не предвиденная ошибка! {ex.Message}");
                return;
            }
        }
        public static void SetParameterValue(Element element, Guid paramGuid, double value)
        {
            Parameter param = element.get_Parameter(paramGuid);
            var logger = ATP_App.GetService<ILoggerService>();

            try
            {
                if (param == null)
                {
                    logger.LogWarning($"Параметр по GUID {paramGuid} отсутствует у элемента {element.Id}", element.Document.Title);
                    return;
                }
                if (param.IsReadOnly)
                {
                    logger.LogWarning($"Параметр {param.Definition?.Name} у элемента {element.Id} только для чтения", element.Document.Title);
                    return;
                }
                if (param.StorageType != StorageType.Double)
                {
                    logger.LogWarning($"Параметр {param.Definition?.Name} у элемента {element.Id} не Double", element.Document.Title);
                    return;
                }

                double oldValue = param.AsDouble();
                const double Tol = 1e-9;
                if (Math.Abs(value - oldValue) > Tol)
                    param.Set(value);
            }
            catch (Exception ex)
            {
                logger.LogError($"Не предвиденная ошибка! {ex.Message}");
                return;
            }
        }

        /// <summary>
        /// Получаем значение параметра (безопасно)
        /// </summary>
        public static string GetSharedParameterValue(Element element, Guid paramGuid)
        {
            Parameter param = element.get_Parameter(paramGuid);
            string value = GetParameterValue(param);

            if (!string.IsNullOrEmpty(value))
                return value;

            Element elementType = element.Document.GetElement(element.GetTypeId());
            if (elementType != null)
            {
                Parameter typeParam = elementType.get_Parameter(paramGuid);
                string typeValue = GetParameterValue(typeParam);
                if (!string.IsNullOrEmpty(typeValue))
                    return typeValue;
            }

            return $"Проверьте наличие параметра {param?.Definition?.Name ?? paramGuid.ToString()}";
        }

        public static string GetProjectParameterValue(Element element, string paramName)
        {
            Parameter param = element.LookupParameter(paramName);
            string value = GetParameterValue(param);

            if (!string.IsNullOrEmpty(value))
                return value;

            Element elementType = element.Document.GetElement(element.GetTypeId());
            if (elementType != null)
            {
                Parameter typeParam = elementType.LookupParameter(paramName);
                string typeValue = GetParameterValue(typeParam);
                if (!string.IsNullOrEmpty(typeValue))
                    return typeValue;
            }

            return $"Проверьте наличие параметра {paramName}";
        }

        private static string GetParameterValue(Parameter param)
        {
            if (param == null) return null;
            if (param.StorageType == StorageType.String)
                return param.AsString();
            return param.AsValueString();
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

        public static void AddSharedParameter(Document doc, string SharedParameterName, Guid SharedParameterGuid, BuiltInCategory cat)
        {
            Application app = doc.Application;
            string originalFile = app.SharedParametersFilename;
            app.SharedParametersFilename = dictionaryGUID.SharedParameterFilePath;

            DefinitionFile defFile = app.OpenSharedParameterFile();
            if (defFile == null)
                throw new Exception("Не удалось открыть файл общих параметров.");

            Definition definition = null;

            foreach (DefinitionGroup group in defFile.Groups)
            {
                foreach (Definition def in group.Definitions)
                {
                    if (def.Name == SharedParameterName &&
                        def is ExternalDefinition extDef &&
                        extDef.GUID == SharedParameterGuid)
                    {
                        definition = def;
                        break;
                    }
                }

                if (definition != null)
                    break;
            }

            if (definition == null)
                throw new Exception($"Параметр '{SharedParameterName}' с GUID {SharedParameterGuid} не найден в файле.");

            Category targetCategory = doc.Settings.Categories.get_Item(cat);
            BindingMap map = doc.ParameterBindings;

            bool bindingExist = map.Contains(definition);
            bool categoryAlreadyBound = false;

            using (Transaction tr = new Transaction(doc, "Обновить/добавить общий параметр"))
            {
                tr.Start();

                if (bindingExist)
                {
                    Binding existingBinding = map.get_Item(definition);

                    // Проверяем тип привязки
                    CategorySet existingCategories = null;
                    if (existingBinding is InstanceBinding ib)
                        existingCategories = ib.Categories;
                    else if (existingBinding is TypeBinding tb)
                        existingCategories = tb.Categories;

                    // Если категория уже есть — ничего не делаем
                    if (existingCategories.Cast<Category>().Any(c => c.Id == targetCategory.Id))
                    {
                        categoryAlreadyBound = true;
                    }
                    else
                    {
                        // Добавляем новую категорию в существующий CategorySet
                        CategorySet newCatSet = app.Create.NewCategorySet();
                        foreach (Category c in existingCategories)
                            newCatSet.Insert(c);
                        newCatSet.Insert(targetCategory);

                        Binding newBinding = (existingBinding is InstanceBinding) 
                            ? (Binding)new InstanceBinding(newCatSet) 
                            : new TypeBinding(newCatSet);

                        map.ReInsert(definition, newBinding, BuiltInParameterGroup.PG_DATA);
                    }
                }
                else
                {
                    // Параметр ещё не был привязан — создаём новую привязку
                    CategorySet newCatSet = app.Create.NewCategorySet();
                    newCatSet.Insert(targetCategory);

                    Binding binding = new InstanceBinding(newCatSet);
                    map.Insert(definition, binding, BuiltInParameterGroup.PG_DATA);
                }

                tr.Commit();
            }

            app.SharedParametersFilename = originalFile; // Восстанавливаем оригинальный путь
        }

        public static void AddSharedParameter(Document doc, string SharedParameterName, Guid SharedParameterGuid, Element elem)
        {
            Application app = doc.Application;
            string originalFile = app.SharedParametersFilename;
            app.SharedParametersFilename = dictionaryGUID.SharedParameterFilePath;
            BuiltInCategory cat = (BuiltInCategory)elem.Category.Id.IntegerValue;

            if (cat.Equals(null))
            {
                return;
            }

            DefinitionFile defFile = app.OpenSharedParameterFile();
            if (defFile == null)
                throw new Exception("Не удалось открыть файл общих параметров.");

            Definition definition = null;

            foreach (DefinitionGroup group in defFile.Groups)
            {
                foreach (Definition def in group.Definitions)
                {
                    if (def.Name == SharedParameterName &&
                        def is ExternalDefinition extDef &&
                        extDef.GUID == SharedParameterGuid)
                    {
                        definition = def;
                        break;
                    }
                }

                if (definition != null)
                    break;
            }

            if (definition == null)
                throw new Exception($"Параметр '{SharedParameterName}' с GUID {SharedParameterGuid} не найден в файле.");

            Category targetCategory = doc.Settings.Categories.get_Item(cat);
            BindingMap map = doc.ParameterBindings;

            bool bindingExist = map.Contains(definition);
            bool categoryAlreadyBound = false;

            if (bindingExist)
            {
                Binding existingBinding = map.get_Item(definition);

                // Проверяем тип привязки
                CategorySet existingCategories = null;
                if (existingBinding is InstanceBinding ib)
                    existingCategories = ib.Categories;
                else if (existingBinding is TypeBinding tb)
                    existingCategories = tb.Categories; 

                // Если категория уже есть — ничего не делаем
                if (existingCategories.Cast<Category>().Any(c => c.Id == targetCategory.Id))
                {
                    categoryAlreadyBound = true;
                }
                else
                {
                    // Добавляем новую категорию в существующий CategorySet
                    CategorySet newCatSet = app.Create.NewCategorySet();
                    foreach (Category c in existingCategories)
                        newCatSet.Insert(c);
                    newCatSet.Insert(targetCategory);

                    Binding newBinding = (existingBinding is InstanceBinding)
                        ? (Binding)new InstanceBinding(newCatSet)
                        : new TypeBinding(newCatSet);

                    map.ReInsert(definition, newBinding, BuiltInParameterGroup.PG_DATA);
                }
            }
            else
            {
                // Параметр ещё не был привязан — создаём новую привязку
                CategorySet newCatSet = app.Create.NewCategorySet();
                newCatSet.Insert(targetCategory);

                Binding binding = new InstanceBinding(newCatSet);
                map.Insert(definition, binding, BuiltInParameterGroup.PG_DATA);
            }

            app.SharedParametersFilename = originalFile; // Восстанавливаем оригинальный путь
        }

        private const string ExcludedWorksetPrefix = "000_";
        private const string ExcludedWorksetExactName = "020_Временные элементы";
        private const string DefaultEmptyGroupKey = "(EMPTY)";

        /// <summary>
        /// Collect all instance elements of the given category, excluding elements that are placed
        /// on worksets whose names start with "000_" or equal "020_Временные элементы".
        /// Returns a List-based IList for speed / memory and convenient downstream grouping.
        /// </summary>
        public static IList<ElementId> CollectElementIdsByCategoryExludingWorksets(Document doc, BuiltInCategory category)
        {
            if (doc == null) throw new ArgumentNullException(nameof(doc));

            HashSet<WorksetId> excludeWorksetIds = doc.IsWorkshared
                ? GetExcludedWorksetIds(doc)
                : new HashSet<WorksetId>();

            var result = new List<ElementId>(capacity: 1024);

            // Single pass over category instances (O(N))
            FilteredElementCollector collector = new FilteredElementCollector(doc)
                .WhereElementIsNotElementType()
                .OfCategory(category);

            foreach (Element e in collector)
            {
                if (e == null) continue;

                if (doc.IsWorkshared && excludeWorksetIds.Count > 0)
                {
                    WorksetId wsId = e.WorksetId;
                    if (wsId != null && wsId != WorksetId.InvalidWorksetId && excludeWorksetIds.Contains(wsId))
                        continue;
                }

                result.Add(e.Id);
            }

            return result;
        }

        private static HashSet<WorksetId> GetExcludedWorksetIds(Document doc)
        {
            var result = new HashSet<WorksetId>();

            FilteredWorksetCollector worksets = new FilteredWorksetCollector(doc)
                .OfKind(WorksetKind.UserWorkset);

            foreach (Workset ws in worksets)
            {
                if (ws == null) continue;

                string name = ws.Name ?? string.Empty;

                if (name.StartsWith(ExcludedWorksetPrefix, StringComparison.OrdinalIgnoreCase) ||
                    string.Equals(name, ExcludedWorksetExactName, StringComparison.OrdinalIgnoreCase))
                {
                    result.Add(ws.Id);
                }
            }

            return result;
        }
    }
}
