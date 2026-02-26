using ATP_Common_Plugin.Services;
using ATP_Common_Plugin.Utils;
using Autodesk.Revit.Attributes;
using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;
using Parameter = Autodesk.Revit.DB.Parameter;

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

            IList<Element> airTerms = selecttionBuiltInInstance.selectInstanceOfCategory(doc, BuiltInCategory.OST_DuctTerminal);
            IList<Element> mechEquip = selecttionBuiltInInstance.selectInstanceOfCategory(doc, BuiltInCategory.OST_MechanicalEquipment);
            IList<Element> plumbFix = selecttionBuiltInInstance.selectInstanceOfCategory(doc, BuiltInCategory.OST_PlumbingFixtures);
            IList<Element> ductFittings = selecttionBuiltInInstance.selectInstanceOfCategory(doc, BuiltInCategory.OST_DuctFitting);
            IList<Element> ductAccessory = selecttionBuiltInInstance.selectInstanceOfCategory(doc, BuiltInCategory.OST_DuctAccessory);
            IList<Element> pipeFittings = selecttionBuiltInInstance.selectInstanceOfCategory(doc, BuiltInCategory.OST_PipeFitting);
            IList<Element> pipeAccessory = selecttionBuiltInInstance.selectInstanceOfCategory(doc, BuiltInCategory.OST_PipeAccessory);

            var batches = new (string title, IList<Element> elems)[]
            {
                ("Обработка вложенных элементов сантехнических приборов", plumbFix),
                ("Обработка вложенных элементов оборудования", mechEquip),
                ("Обработка вложенных элементов воздухораспределителей", airTerms),
                ("Обработка вложенных элементов арматуры воздуховодов", ductAccessory),
                ("Обработка вложенных элементов фитингов воздуховодов", ductFittings),
                ("Обработка вложенных элементов арматуры трубопроводов", pipeAccessory),
                ("Обработка вложенных элементов фитингов трубопроводов", pipeFittings),
            };

            try
            {
                foreach (var (title, elems) in batches)
                {
                    using (var tr = new Transaction(doc, title))
                    {
                        logger.LogInfo("Начало: " + title, docName);
                        tr.Start();

                        SetDependentElemParam(doc, elems, ref log);

                        tr.Commit();
                        logger.LogInfo("Завершение: " + title, docName);
                    }
                }
            }
            catch (Exception ex)
            {
                logger.LogError($"Непредвиденная ошибка: {ex}", docName);
                return Result.Failed;
            }
            logger.LogInfo("Параметры переданы во вложенные семейства", docName);

            TaskDialog.Show("Успех", "Параметры переданы во вложенные семейства");
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

            const double KompStep = 0.01;
            const double SubStep = 0.0001; 
            Dictionary<string, double> setDict = new Dictionary<string, double>();

            var logger = ATP_App.GetService<ILoggerService>();
            if (mass_elements.Count > 0)
            {
                foreach (Element host in mass_elements)
                {
                    try
                    {
                        // Get all parameters from host
                        if (!(host is FamilyInstance hostInstance)) continue;

                        string hostName = RevitUtils.GetSharedParameterValue(host, dictionaryGUID.ADSKName) ?? "У основного элемента не заполнен ADSK_Наименование";
                        string hostMark = RevitUtils.GetSharedParameterValue(host, dictionaryGUID.ADSKMark) ?? "У основного элемента не заполнен ADSK_Марка";
                        string hostKey = hostName + " | " + hostMark;
                        string hostSystemName = RevitUtils.GetProjectParameterValue(host, "ИмяСистемы") ?? "Элемент без системы";
                        string hostKomp = RevitUtils.GetSharedParameterValue(host, dictionaryGUID.ADSKKomp) ?? "У основного элемента не заполнен ADSK_Комплект";
                        string hostLvl = RevitUtils.GetSharedParameterValue(host, dictionaryGUID.ADSKLevel) ?? "У основного элемента не заполнен ADSK_Этаж";
                        string hostGroupString = RevitUtils.GetSharedParameterValue(host, dictionaryGUID.ADSKGroup);
                        double hostGroup; 
                        double.TryParse(hostGroupString, NumberStyles.Float, CultureInfo.InvariantCulture, out hostGroup);

                        double setNumber = hostGroup;

                        // Get and set all parameters of sub elements
                        var subElements = hostInstance.GetSubComponentIds()
                            .Select(id => doc.GetElement(id))
                            .Where(e => e != null)
                            .ToList();

                        if (subElements.Count == 0) continue;

                        Parameter pHostGroup = host.get_Parameter(dictionaryGUID.ADSKGroup);
                        bool hostGroupWritable = pHostGroup != null && !pHostGroup.IsReadOnly;
                        if (!hostGroupWritable)
                        {
                            foreach (var id in ((FamilyInstance)host).GetSubComponentIds())
                            {
                                var sub = doc.GetElement(id);
                                if (sub == null) continue;
                                RevitUtils.SetParameterValue(sub, "ИмяСистемы", hostSystemName);
                                RevitUtils.SetParameterValue(sub, dictionaryGUID.ADSKKomp, hostKomp);
                                RevitUtils.SetParameterValue(sub, dictionaryGUID.ADSKLevel, hostLvl);
                            }
                            log += $"{host.Id} - Параметр ADSK_Группирование недоступен для записи\n";
                            continue; 
                        }

                        int subIndex = 1;
                        var subNameMarkMap = new Dictionary<string, double>();

                        foreach (Element subElem in subElements)
                        {
                            if (hostGroup == 0)
                            {
                                logger.LogWarning($"{host.Id} - У основы не заполнен параметр ADSK_Группирование", docName);
                                continue;
                            }

                            if (host.Category.Id.IntegerValue == (int)BuiltInCategory.OST_DuctFitting 
                                || host.Category.Id.IntegerValue == (int)BuiltInCategory.OST_PipeFitting)
                            {
                                RevitUtils.SetParameterValue(subElem, "ИмяСистемы", hostSystemName);
                                RevitUtils.SetParameterValue(subElem, dictionaryGUID.ADSKGroup, "4");
                                continue;
                            }

                            // Get number of set and put new number to setDictionary
                            double existing;
                            if (setDict.TryGetValue(hostKey, out existing))
                            {
                                setNumber = existing;
                            }
                            else
                            {
                                double next = setDict.Count > 0
                                    ? Math.Round(setDict.Values.Max() + KompStep, 4)
                                    : Math.Round(hostGroup + KompStep, 4);

                                setNumber = next;
                                setDict[hostKey] = setNumber;
                            }

                            string subName = RevitUtils.GetSharedParameterValue(subElem, dictionaryGUID.ADSKName)
                                ?? RevitUtils.GetSharedParameterValue(doc.GetElement(subElem.GetTypeId()), dictionaryGUID.ADSKName)
                                ?? string.Empty;

                            string subMark = RevitUtils.GetSharedParameterValue(subElem, dictionaryGUID.ADSKMark)
                                ?? RevitUtils.GetSharedParameterValue(doc.GetElement(subElem.GetTypeId()), dictionaryGUID.ADSKMark)
                                ?? string.Empty;

                            string subKey = subName + " | " + subMark;

                            double subNumber;
                            if (subNameMarkMap.TryGetValue(subKey, out subNumber))
                            {
                                // уже назначали для такой пары — оставляем прежнее число
                            }
                            else
                            {
                                subNumber = Math.Round(setNumber + subIndex * SubStep, 4);
                                subNameMarkMap[subKey] = subNumber;
                                subIndex++; // следующий «новый» sub получит следующий хвост
                            }

                            // 3) Записываем (ровно 4 знака, инвариантная точка)
                            string subVal = subNumber.ToString("0.0000", CultureInfo.InvariantCulture);
                            RevitUtils.SetParameterValue(subElem, dictionaryGUID.ADSKGroup, subVal);
                            RevitUtils.SetParameterValue(subElem, "ИмяСистемы", hostSystemName);
                            RevitUtils.SetParameterValue(subElem, dictionaryGUID.ADSKKomp, hostKomp);
                            RevitUtils.SetParameterValue(subElem, dictionaryGUID.ADSKLevel, hostLvl);
                        }

                        string hostVal = setNumber.ToString("0.0000", CultureInfo.InvariantCulture);
                        RevitUtils.SetParameterValue(host, dictionaryGUID.ADSKGroup, hostVal);
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