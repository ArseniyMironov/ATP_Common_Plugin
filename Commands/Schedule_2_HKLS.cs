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
    class Schedule_2_HKLS : IExternalCommand
    {
        public Result Execute(ExternalCommandData commandData, ref string message, ElementSet elements)
        {
            UIApplication uiApp = commandData.Application;
            UIDocument uiDoc = uiApp.ActiveUIDocument;
            Document doc = uiDoc.Document;
            string docName = doc.Title;

            var logger = ATP_App.GetService<ILoggerService>();

            string log = "";

            IList<Element> mechEquip = selecttionBuiltInInstance.selectInstanceOfCategory(doc, BuiltInCategory.OST_MechanicalEquipment);
            IList<Element> plumbFix = selecttionBuiltInInstance.selectInstanceOfCategory(doc, BuiltInCategory.OST_PlumbingFixtures);
            IList<Element> ductFittings = selecttionBuiltInInstance.selectInstanceOfCategory(doc, BuiltInCategory.OST_DuctFitting);
            IList<Element> ductAccessory = selecttionBuiltInInstance.selectInstanceOfCategory(doc, BuiltInCategory.OST_DuctAccessory);
            IList<Element> pipeFittings = selecttionBuiltInInstance.selectInstanceOfCategory(doc, BuiltInCategory.OST_PipeFitting);
            IList<Element> pipeAccessory = selecttionBuiltInInstance.selectInstanceOfCategory(doc, BuiltInCategory.OST_PipeAccessory);
            try
            {
                using (Transaction tr = new Transaction(doc, "Обработка вложенных элементов сантехнических приборов"))
                {
                    logger.LogInfo("Начало обработки вложенных элементов сантехнических приборов", docName);
                    tr.Start();

                    SetDependentElemParam(doc, plumbFix, ref log);

                    tr.Commit();
                    logger.LogInfo("Завершение обработки вложенных элементов сантехнических приборов", docName);
                }

                using (Transaction tr = new Transaction(doc, "Обработка вложенных элементов оборудования"))
                {
                    logger.LogInfo("Начало обработки вложенных элементов оборудования", docName);
                    tr.Start();

                    SetDependentElemParam(doc, mechEquip, ref log);

                    tr.Commit();
                    logger.LogInfo("Завершение обработки вложенных элементов оборудования", docName);
                }

                using (Transaction tr = new Transaction(doc, "Обработка вложенных элементов арматуры воздуховодов"))
                {
                    logger.LogInfo("Начало обработки вложенных элементов арматуры воздуховодов", docName);
                    tr.Start();

                    SetDependentElemParam(doc, ductAccessory, ref log);

                    tr.Commit();
                    logger.LogInfo("Завершение обработки вложенных элементов арматуры воздуховодов", docName);
                }

                using (Transaction tr = new Transaction(doc, "Обработка вложенных элементов фитингов воздуховодов"))
                {
                    logger.LogInfo("Начало обработки вложенных элементов фитингов воздуховодов", docName);
                    tr.Start();

                    SetDependentElemParam(doc, ductFittings, ref log);

                    tr.Commit();
                    logger.LogInfo("Завершение обработки вложенных элементов фитингов воздуховодов", docName);
                }

                using (Transaction tr = new Transaction(doc, "Обработка вложенных элементов арматуры трубопроводов"))
                {
                    logger.LogInfo("Начало обработки вложенных элементов арматуры трубопроводов", docName);
                    tr.Start();

                    SetDependentElemParam(doc, pipeAccessory, ref log);

                    tr.Commit();
                    logger.LogInfo("Завершение обработки вложенных элементов арматуры трубопроводов", docName);
                }

                using (Transaction tr = new Transaction(doc, "Обработка вложенных элементов фитингов трубопроводов"))
                {
                    logger.LogInfo("Начало обработки вложенных элементов фитингов трубопроводов", docName);
                    tr.Start();

                    SetDependentElemParam(doc, pipeFittings, ref log);

                    tr.Commit();
                    logger.LogInfo("Завершение обработки вложенных элементов фитингов трубопроводов", docName);
                }
            }
            catch
            {
                logger.LogError($"Непредвиденная ошибка: {log}", docName);
                return Result.Failed;
            }
            logger.LogInfo("Параметры переданы во вложенные семейства", docName);
            return Result.Succeeded;
        }

        /// <summary>
        /// Метод для обработки основного и вложенных элементов
        /// </summary>
        /// <param name="doc"></param>
        /// <param name="mass_elements"></param>
        /// <param name="log"></param>
        private void SetDependentElemParam(Document doc, IList<Element> mass_elements, ref string log)
        {
            string docName = doc.Title;
            double komp_index = 0.01;
            Dictionary<string, Dictionary<string, double>> setDict = new Dictionary<string, Dictionary<string, double>>();

            var logger = ATP_App.GetService<ILoggerService>();
            if (mass_elements.Count > 0)
            {
                foreach (Element host in mass_elements)
                {
                    try
                    {
                        // Get all parameters from host
                        if (!(host is FamilyInstance hostInstance)) continue;

                        string hostName = RevitUtils.GetSharedParameterValue(host, dictionaryGUID.ADSKName) ?? "У основного элемнта не заполнен ADSK_Наименования";
                        string hostSystemName = RevitUtils.GetProjectParameterValue(host, "ИмяСистемы");
                        string hostKomp = RevitUtils.GetSharedParameterValue(host, dictionaryGUID.ADSKKomp) ?? "Оборудование без компекта";
                        string hostGroupString = RevitUtils.GetSharedParameterValue(host, dictionaryGUID.ADSKGroup).Split('.').First();
                        bool isGroupCorrect = Double.TryParse($"{hostGroupString}", out double hostGroup);

                        if (isGroupCorrect == false)
                        {
                            logger.LogWarning($"{host.Id} - Параметр ADSK_Группирование заполнене некорректно", docName);
                            continue;
                        }

                        double setNumber = hostGroup;

                        // Get and set all parameters of sub elements
                        IList<Element> subElements = hostInstance.GetSubComponentIds()
                            .Select(id => doc.GetElement(id))
                            .Where(e => e != null)
                            .ToList();

                        if (subElements.Count > 0)
                        {
                            double element_position = 0.0001;
                            foreach (Element subElem in subElements)
                            {
                                RevitUtils.SetParameterValue(subElem, dictionaryGUID.ADSKKomp, hostKomp);
                                RevitUtils.SetParameterValue(subElem, "ИмяСистемы", hostSystemName);

                                if (hostGroup == 0)
                                {
                                    logger.LogWarning($"{host.Id} - У основы не заполнен параметр ADSK_Группирование", docName);
                                    continue;
                                }

                                if (host.Category.Name.Contains("Fitting") || host.Category.Name.Contains("оединительн"))
                                {
                                    RevitUtils.SetParameterValue(subElem, dictionaryGUID.ADSKGroup, hostGroup); 
                                    continue;
                                }

                                // Get number of set and put new number to setDictionary
                                setNumber += komp_index;
                                if (setDict.TryGetValue(hostSystemName, out var innerDict))
                                {
                                    if (innerDict.TryGetValue(hostName, out double existingValue))
                                    {
                                        setNumber = existingValue;
                                    }
                                    else
                                    {
                                        setNumber = (innerDict.Count > 0) ? innerDict.Values.Max() + 0.01 : 0.01;
                                        innerDict.Add(hostName, setNumber);
                                    }
                                }
                                else
                                {
                                    setNumber = hostGroup + 0.01;
                                    setDict.Add(hostSystemName, new Dictionary<string, double> { { hostName, setNumber } });
                                }

                                RevitUtils.SetParameterValue(subElem, dictionaryGUID.ADSKGroup, $"{setNumber + element_position}");
                                element_position += 0.0001;
                            }
                            RevitUtils.SetParameterValue(host, dictionaryGUID.ADSKGroup, $"{setNumber}00");
                            element_position = 0.0001;
                        }
                    }
                    catch (Exception ex)
                    {
                        logger.LogError($"Ошибка при обработке элемента - {host.Id}. {ex.Message}", docName);
                    }
                }
            }
            else
            {
                logger.LogWarning("Нет элементов для обработки" , docName);
            }
        }
    }
}
