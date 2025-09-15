using ATP_Common_Plugin.Utils;
using Autodesk.Revit.ApplicationServices;
using Autodesk.Revit.Attributes;
using Autodesk.Revit.DB;
using Autodesk.Revit.Exceptions;
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
                .OrderBy(l => l.Elevation)
                .ToList();

            if (levels.Count == 0)
            {
                TaskDialog.Show("Ошибка", "Уровни в модели не найдены");
                //logger.LogWarning("Уровни в модели не найдены", docName);
                return Result.Failed;
            }

            // Сопоставление Elevation -> строка ADSK_Этаж (до первого "_")
            var levelDict = new Dictionary<Level, string>();
            foreach (var level in levels)
            {
                var nameParts = level.Name.Split('_');
                if (nameParts.Length > 0)
                    levelDict[level] = nameParts[0];
            }

            // Коллекция всех элементов модели, кроме типов и уровней
            var elementsToProcess = new FilteredElementCollector(doc)
                .WhereElementIsNotElementType()
                .Where(e => !(e is Level))
                .ToList();

            int processed = 0;
            int processedGood = 0;

            using (Transaction tr = new Transaction(doc, "Заполнение ADSK_Этаж"))
            {
                tr.Start();

                foreach (var elem in elementsToProcess)
                {
                    if (elem == null && elem.Location == null && elem.get_BoundingBox(null) == null)
                        continue;

                    // Пробуем получить нижнюю Z координату
                    BoundingBoxXYZ bbox = elem.get_BoundingBox(null);
                    if (bbox == null)
                        continue;

                    processed++;

                    double z = bbox.Min.Z;

                    // Учитываем допуск (элемент может быть немного ниже уровня)
                    Level nearestLevel = levels
                        .Where(lvl => z >= lvl.Elevation - Tolerance)
                        .OrderByDescending(lvl => lvl.Elevation)
                        .FirstOrDefault();

                    if (nearestLevel == null || !levelDict.ContainsKey(nearestLevel))
                    {
                        continue;
                    }

                    string floorValue = levelDict[nearestLevel];


                    // Параметр ADSK_Этаж
                    Parameter param = elem.get_Parameter(dictionaryGUID.ADSKLevel);
                    if (param == null)
                    {
                        try
                        {
                            //RevitUtils.AddSharedParameter(doc, "ADSK_Этаж", dictionaryGUID.ADSKLevel, elem);
                        }
                        catch
                        {
                            continue;
                        }
                    }

                    if (param.IsReadOnly)
                        continue;

                    param.Set(floorValue);
                    processedGood++;
                }

                tr.Commit();
            }

            TaskDialog.Show("Готово", $"Параметр ADSK_Этаж заполнен для {processed} элементов, нормально {processedGood}");
            return Result.Succeeded;
        }
    }
}
