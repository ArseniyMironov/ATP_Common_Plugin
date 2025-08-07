using ATP_Common_Plugin.Utils;
using Autodesk.Revit.ApplicationServices;
using Autodesk.Revit.Attributes;
using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using System.Collections.Generic;
using System.Linq;

namespace ATP_Common_Plugin.Commands
{
    [Regeneration(RegenerationOption.Manual)]
    [Transaction(TransactionMode.Manual)]
    class FillLevel : IExternalCommand
    {
        private const double Tolerance = 100.0 / 304.8;
        public Result Execute(ExternalCommandData commandData, ref string message, ElementSet elements)
        {
            UIDocument uidoc = commandData.Application.ActiveUIDocument;
            Document doc = uidoc.Document;
            Application app = doc.Application;
            string docName = doc.Title;
            //var logger = ATP_App.GetService<ILoggerService>();

            // Получаем уровни
            List<Level> levels = new FilteredElementCollector(doc)
                .OfClass(typeof(Level))
                .Cast<Level>()
                .OrderBy(lvl => lvl.Elevation)
                .ToList();

            if (levels.Count == 0)
            {
                TaskDialog.Show("Ошибка", "Уровни в модели не найдены.");
                //logger.LogWarning("Уровни в модели не найдены.", docName);
                return Result.Failed;
            }

            // Создаем словарь: уровень -> значение ADSK_Этаж
            Dictionary<Level, string> levelDict = new Dictionary<Level, string>();
            foreach (var lvl in levels)
            {
                string[] parts = lvl.Name.Split('_');
                if (parts.Length > 0)
                    levelDict[lvl] = parts[0];
            }

            // Получаем все элементы, у которых есть параметр ADSK_Этаж
            var collector = new FilteredElementCollector(doc)
                .WhereElementIsNotElementType()
                .ToList();

            int total = collector.Count;
            int processed = 0;

            using (Transaction tr = new Transaction(doc, "Заполнение ADSK_Этаж"))
            {
                tr.Start();

                foreach (var elem in collector)
                {
                    processed++;
                    if (processed % 500 == 0)
                    {
                        TaskDialog.Show("Прогресс", $"Обработано {processed} из {total} элементов...");
                    }

                    // Игнорируем вложенные компоненты — они будут обработаны через родителя
                    if (elem is FamilyInstance fi && fi.SuperComponent != null)
                        continue;

                    List<Element> targets = new List<Element>();

                    // Если это FamilyInstance — обрабатываем вложенные
                    if (elem is FamilyInstance instance)
                    {
                        foreach (ElementId id in instance.GetSubComponentIds())
                        {
                            Element sub = doc.GetElement(id);
                            if (sub != null && sub.get_Parameter(dictionaryGUID.ADSKLevel) != null)
                                targets.Add(sub);
                        }
                    }

                    BoundingBoxXYZ bbox = elem.get_BoundingBox(null);
                    if (bbox == null)
                        continue;

                    double z = bbox.Min.Z;

                    Level nearestLevel = levels
                        .Where(lvl => z >= lvl.Elevation)
                        .FirstOrDefault();

                    if (nearestLevel == null || !levelDict.TryGetValue(nearestLevel, out string floorValue))
                        continue;

                    foreach (var target in targets)
                    {
                        Parameter param = target.get_Parameter(dictionaryGUID.ADSKLevel);
                        if (param == null)
                        {
                            var builtInCat = (BuiltInCategory)target.Category.Id.IntegerValue;
                            RevitUtils.AddSharedParameter(doc, "ADSK_Этаж", dictionaryGUID.ADSKLevel, builtInCat);
                        }
                        if (param != null && !param.IsReadOnly)
                            param.Set(floorValue);
                    }
                }

                tr.Commit();
            }

            TaskDialog.Show("Готово", $"Обработано {processed} элементов.\nПараметр ADSK_Этаж задан.");
            return Result.Succeeded;
        }
    }
}
