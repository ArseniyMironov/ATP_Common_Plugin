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
    class FillSlope : IExternalCommand
    {
        public Result Execute(ExternalCommandData commandData, ref string message, ElementSet elements)
        {
            UIApplication uiApp = commandData.Application;
            Document doc = uiApp.ActiveUIDocument.Document;
            string docName = doc.Title;
            var logger = ATP_App.GetService<ILoggerService>();

            try
            {

                logger.LogInfo($"Начало выполнения Fill slope.", docName);

                IList<Element> pipes = new FilteredElementCollector(doc)
                    .OfCategory(BuiltInCategory.OST_PipeCurves)
                    .WhereElementIsNotElementType()
                    .ToElements();

                if (!pipes.Any())
                {
                    logger.LogWarning("Трубопроводы не найдены", docName);
                    return Result.Succeeded;
                }

                int successCount = 0;

                using (Transaction tr = new Transaction(doc, "Заполнение уклона"))
                {
                    tr.Start();

                    foreach (Element pipe in pipes)
                    {
                        Parameter slopeParam = pipe.get_Parameter(BuiltInParameter.RBS_PIPE_SLOPE);
                        if (slopeParam == null || !slopeParam.HasValue) continue;

                        double slope = slopeParam.AsDouble();
                        Parameter adskParam = pipe.get_Parameter(dictionaryGUID.ADSKSlope);
                        if (adskParam == null || adskParam.IsReadOnly) continue;

                        adskParam.Set(slope);
                        successCount++;
                    }

                    tr.Commit();
                }

                logger.LogInfo($"Уклон заполнен у {successCount} труб", docName);
                logger.LogInfo("Команда Fill slope успешно завершена", docName);
                return Result.Succeeded;
            }
            catch (Exception ex)
            {
                logger.LogError($"Ошибка: {ex.Message}", docName);
                return Result.Failed;
            }
        }
    }
}
