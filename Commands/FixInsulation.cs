using ATP_Common_Plugin.Services;
using Autodesk.Revit.Attributes;
using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using System.Collections.Generic;
using System.Linq;

namespace ATP_Common_Plugin.Commands
{
    [Regeneration(RegenerationOption.Manual)]
    [Transaction(TransactionMode.Manual)]
    class FixInsulation : IExternalCommand
    {
        public Result Execute(ExternalCommandData commandData, ref string message, ElementSet elements)
        {
            // Получаем доступ к интерфейсу и документу Revit
            UIApplication uiApp = commandData.Application;
            UIDocument uiDoc = uiApp.ActiveUIDocument;
            Document doc = uiDoc.Document;
            string docName = doc.Title;
            var logger = ATP_App.GetService<ILoggerService>();
            // Список для хранения элементов на удаление
            List<ElementId> elementsToDelete = new List<ElementId>();

            // Создаем список категорий для фильтрации (изоляция труб и воздуховодов)
            IList<BuiltInCategory> cats = new List<BuiltInCategory>
            {
                BuiltInCategory.OST_DuctInsulations,
                BuiltInCategory.OST_PipeInsulations
            };

            ElementMulticategoryFilter insulation_filter = new ElementMulticategoryFilter(cats);

            // Получаем коллектор элементов изоляции
            FilteredElementCollector insulation = new FilteredElementCollector(doc)
                .WherePasses(insulation_filter)
                .WhereElementIsNotElementType();

            // Начинаем транзакцию для изменений в модели
            using (Transaction tr = new Transaction(doc, "Корректировка изоляции"))
            {
                tr.Start();

                int deletedCount = 0;
                int movedCount = 0;
                logger.LogInfo("Начало выполнения Fix Insulation.", docName);
                 
                // Обрабатываем каждый элемент изоляции
                foreach (InsulationLiningBase ins in insulation.Cast<InsulationLiningBase>())
                {
                    // Получаем элемент-основу для изоляции
                    Element host = doc.GetElement(ins.HostElementId);
                    WorksetId ins_workset_id = ins.WorksetId;

                    if (host != null) // Если основа существует
                    {
                        WorksetId host_workset_id = host.WorksetId;

                        // Проверяем соответствие рабочих наборов
                        if (ins_workset_id != host_workset_id)
                        {
                            // Корректируем рабочий набор основы
                            Parameter host_workset = host.get_Parameter(BuiltInParameter.ELEM_PARTITION_PARAM);
                            host_workset.Set(ins_workset_id.IntegerValue);

                            doc.Regenerate(); // Регенерируем модель

                            // Возвращаем исходный рабочий набор
                            host_workset = host.get_Parameter(BuiltInParameter.ELEM_PARTITION_PARAM);
                            host_workset.Set(host_workset_id.IntegerValue);
                            movedCount++;
                        }
                    }
                    else // Если основа отсутствует
                    {
                        elementsToDelete.Add(ins.Id); // Добавляем в список на удаление
                        deletedCount++;
                    }
                }

                // Удаляем элементы без основы
                if (elementsToDelete.Count > 0)
                {
                    doc.Delete(elementsToDelete);
                }

                tr.Commit(); // Завершаем транзакцию

                // Выводим сообщение о завершении
                logger.LogInfo($"Изоляция скорректирована. Перенесено {movedCount} элементов, удален {deletedCount} элементов.", docName);
                return Result.Succeeded;
            }
        }
    }
}