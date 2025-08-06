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

        public static void AddSharedParameter(Application app, Document doc, string SharedParameterName, Guid SharedParameterGuid, BuiltInCategory cat)
        {
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

            //CategorySet categories = app.Create.NewCategorySet();
            //categories.Insert(doc.Settings.Categories.get_Item(cat));

            //Binding binding = new InstanceBinding(categories);

            //BindingMap map = doc.ParameterBindings;
            //using (Transaction tx = new Transaction(doc, "Добавить общий параметр"))
            //{
            //    tx.Start();

            //    map.Insert(definition, binding, BuiltInParameterGroup.PG_DATA);

            //    tx.Commit();
            //}

            app.SharedParametersFilename = originalFile; // Восстанавливаем оригинальный путь
        }
    }
}
