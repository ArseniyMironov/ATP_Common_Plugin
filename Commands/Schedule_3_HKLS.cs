using ATP_Common_Plugin.Services;
using ATP_Common_Plugin.Utils;
using Autodesk.Revit.Attributes;
using Autodesk.Revit.DB;
using Autodesk.Revit.DB.Mechanical;
using Autodesk.Revit.UI;
using System;
using System.Collections.Generic;
using System.Linq;

namespace ATP_Common_Plugin.Commands
{
    [Regeneration(RegenerationOption.Manual)]
    [Transaction(TransactionMode.Manual)]
    class Schedule_3_HKLS : IExternalCommand
    {
        
        public Result Execute(ExternalCommandData commandData, ref string message, ElementSet elements)
        {
            UIApplication uiApp = commandData.Application;
            UIDocument uiDoc = uiApp.ActiveUIDocument;
            Document doc = uiDoc.Document;
            string docName = doc.Title;

            var logger = ATP_App.GetService<ILoggerService>();

            IList<Element> ducts = selecttionBuiltInInstance.selectInstanceOfCategory(doc, BuiltInCategory.OST_DuctCurves);
            IList<Element> ductsFlex = selecttionBuiltInInstance.selectInstanceOfCategory(doc, BuiltInCategory.OST_FlexDuctCurves);
            IList<Element> ductFittings = selecttionBuiltInInstance.selectInstanceOfCategory(doc, BuiltInCategory.OST_DuctFitting);
            IList<Element> ductInsulation = selecttionBuiltInInstance.selectInstanceOfCategory(doc, BuiltInCategory.OST_DuctInsulations);
            IList<Element> pipes = selecttionBuiltInInstance.selectInstanceOfCategory(doc, BuiltInCategory.OST_PipeCurves);
            IList<Element> pipesFlex = selecttionBuiltInInstance.selectInstanceOfCategory(doc, BuiltInCategory.OST_FlexPipeCurves);
            IList<Element> pipeInsulation = selecttionBuiltInInstance.selectInstanceOfCategory(doc, BuiltInCategory.OST_PipeInsulations);

            string testLog = "";
            double koefDucts = 1.1; // Коэфициент запаса воздуховодов
            double koef = 1.2; // Коэфициент запаса

            // Ducts
            if (ducts.Count != 0)
            {
                using (Transaction tr = new Transaction(doc, "Обработка воздуховодов"))
                {
                    logger.LogInfo("Начало обработки воздуховодов", docName);
                    tr.Start();

                    foreach (Element duct in ducts)
                    {
                        try
                        {
                            if (RevitUtils.CheckElement(duct))
                            {
                                continue;
                            }

                            // Получение существующих параметров 
                            string ductSize = duct.get_Parameter(BuiltInParameter.RBS_CALCULATED_SIZE).AsValueString();
                            string thickness = "0.9";
                            Parameter paramLength = duct.get_Parameter(BuiltInParameter.CURVE_ELEM_LENGTH);
                            GetDuctThickness(duct, ref ductSize, ref thickness);
                            Element ductType = doc.GetElement(duct.GetTypeId());
                            string ductUnits = ductType.get_Parameter(dictionaryGUID.ADSKUnit).AsValueString();
                            ForgeTypeId countParmUnit = paramLength.GetUnitTypeId();
                            ForgeTypeId countUnit = UnitTypeId.Meters;

                            // Генерация новых значений
                            string newName = $"{ductType.get_Parameter(BuiltInParameter.ALL_MODEL_TYPE_COMMENTS).AsValueString()} {ductSize}  δ= {thickness}";
                            double newCount = Int32.Parse(paramLength.AsValueString());
                            double.TryParse(thickness.Trim(), out double newThickness);

                            if (ductUnits == "м²") // Корректировка колличества, если единицы измерения не метры
                            {
                                Parameter param = duct.get_Parameter(BuiltInParameter.RBS_CURVE_SURFACE_AREA);
                                ForgeTypeId unitType = param.GetUnitTypeId();
                                countUnit = UnitTypeId.SquareMeters;
                                newCount = UnitUtils.ConvertFromInternalUnits(param.AsDouble(), unitType);
                            }
                            newCount = UnitUtils.Convert(newCount, countParmUnit, countUnit);

                            // Обработка параметров 
                            RevitUtils.SetParameterValue(duct, dictionaryGUID.ADSKName, newName);
                            RevitUtils.SetParameterValue(duct, dictionaryGUID.ADSKSign, ductSize);
                            RevitUtils.SetParameterValue(duct, dictionaryGUID.ADSKCount, newCount * koefDucts);
                            RevitUtils.SetParameterValue(duct, dictionaryGUID.ADSKThicknes, UnitUtils.ConvertToInternalUnits(newThickness, UnitTypeId.Millimeters));
                        }
                        catch (Exception ex)
                        {
                            logger.LogError($"Ошибка при обработке воздуховодов {duct.Id} {ex}", docName);
                            //testLog.Concat($"Ошибка при обработке воздуховодов {duct.Id} {ex}");
                            //TaskDialog.Show("Ошибка", $"Ошибка при обработке воздуховодов {duct.Id} {ex.ToString()}");
                        }
                    }
                    tr.Commit();
                    logger.LogInfo("Завершение обработки воздуховодов", docName);
                }
            }

            // Duct Fittings
            if (ductFittings.Count != 0)
            {
                using (Transaction tr = new Transaction(doc, "Обработка соединительных деталей воздуховодов"))
                {
                    logger.LogInfo("Начало обработки соединительных деталей воздуховодов", docName);
                    tr.Start();

                    foreach (Element ductFitting in ductFittings)
                    {
                        try
                        {
                            if (RevitUtils.CheckElement(ductFitting))
                            {
                                continue;
                            }
                            //if (ductFitting.)
                            // Получение сеществующих параметров
                            string ductSize = ductFitting.get_Parameter(BuiltInParameter.RBS_CALCULATED_SIZE).AsValueString();
                            string thickness = "0.9";
                            GetDuctThickness(ductFitting, ref ductSize, ref thickness);

                            Element ductFittingType = doc.GetElement(ductFitting.GetTypeId());
                            string ductFittingUnits = ductFitting.get_Parameter(dictionaryGUID.ADSKUnit).AsValueString();
                            double.TryParse(thickness.Trim(), out double newThickness);

                            // Генерация новых значений 
                            if (ductFittingUnits == "м²")
                            {
                                Parameter param = ductFitting.get_Parameter(dictionaryGUID.ADSKSizeArea);
                                ForgeTypeId unitType = param.GetTypeId();
                                ForgeTypeId countUnit = UnitTypeId.SquareMeters;
                                double ductFittingValue = UnitUtils.ConvertFromInternalUnits(param.AsDouble(), unitType);
                                RevitUtils.SetParameterValue(ductFitting, dictionaryGUID.ADSKCount, ductFittingValue);
                            }

                            string newName = $"{ductFittingType.get_Parameter(BuiltInParameter.ALL_MODEL_TYPE_COMMENTS).AsValueString()} {ductSize}  δ = {thickness}";

                            RevitUtils.SetParameterValue(ductFitting, dictionaryGUID.ADSKName, newName);
                            RevitUtils.SetParameterValue(ductFitting, dictionaryGUID.ADSKSign, ductSize);
                            RevitUtils.SetParameterValue(ductFitting, dictionaryGUID.ADSKThicknes, UnitUtils.ConvertToInternalUnits(newThickness, UnitTypeId.Millimeters));
                        }
                        catch (Exception ex)
                        {
                            testLog.Concat($"Ошибка при обработке соединительных деталей воздуховодов {ductFitting.Id} {ex}");
                            //TaskDialog.Show("Ошибка", $"Ошибка при обработке соединительных деталей воздуховодов {ductFitting.Id} {ex.ToString()}");
                        }
                    }
                    tr.Commit();
                    logger.LogInfo("Завершение обработки соединительных деталей воздуховодов", docName);
                }
            }

            // Duct insulation
            if (ductInsulation.Count > 0)
            {
                using (Transaction tr = new Transaction(doc, "Обработка изоляции воздуховодов"))
                {
                    logger.LogInfo("Начало обработки изоляции воздуховодов", docName);
                    tr.Start();

                    foreach (Element insulation in ductInsulation)
                    {
                        try
                        {
                            if (RevitUtils.CheckElement(insulation))
                            {
                                continue;
                            }

                            // Получение сеществующих параметров
                            Element insulationType = doc.GetElement(insulation.GetTypeId());
                            Element insulationHost = doc.GetElement((insulation as InsulationLiningBase).HostElementId);
                            bool isHostFitting = insulationHost.Category.Name.IndexOf("Fittings", StringComparison.OrdinalIgnoreCase) >= 0 || insulationHost.Category.Name.Contains("Cоединительные");
                            string insulationUnit = RevitUtils.GetSharedParameterValue(insulation, dictionaryGUID.ADSKUnit);
                            double thickness = insulation.get_Parameter(BuiltInParameter.RBS_INSULATION_THICKNESS_FOR_DUCT).AsDouble();

                            // Генерация новых значений
                            string newName = $"{insulationType.get_Parameter(BuiltInParameter.ALL_MODEL_TYPE_COMMENTS).AsValueString()} толшиной {UnitUtils.ConvertFromInternalUnits(thickness, UnitTypeId.Millimeters)} мм";
                            double count = 0;
                            if (insulationUnit == "м")
                            {
                                count = UnitUtils.ConvertFromInternalUnits(insulation.get_Parameter(BuiltInParameter.CURVE_ELEM_LENGTH).AsDouble(), UnitTypeId.Millimeters);
                            }
                            else if (insulationUnit == "м²" || insulationUnit == "m²")
                            {
                                if (isHostFitting)
                                {
                                    count = UnitUtils.ConvertFromInternalUnits(insulationHost.get_Parameter(dictionaryGUID.ADSKSizeArea).AsDouble(), UnitTypeId.SquareMeters);
                                }
                                else
                                {
                                    count = UnitUtils.ConvertFromInternalUnits(insulation.get_Parameter(BuiltInParameter.RBS_CURVE_SURFACE_AREA).AsDouble(), UnitTypeId.SquareMeters);
                                }
                            }

                            // Обработка параметров
                            RevitUtils.SetParameterValue(insulation, dictionaryGUID.ADSKName, newName);
                            RevitUtils.SetParameterValue(insulation, dictionaryGUID.ADSKCount, count * koef);
                            RevitUtils.SetParameterValue(insulation, dictionaryGUID.ADSKThicknes, thickness);

                        }
                        catch (Exception ex)
                        {
                            testLog.Concat($"Ошибка при обработке изоляции воздуховодов {ex}");
                            //TaskDialog.Show("Ошибка", $"Ошибка при обработке изоляции воздуховодов /n {ex.ToString()}");
                        }
                    }
                    tr.Commit();
                    logger.LogInfo("Завершение обработки изоляции воздуховодов", docName);
                }
            }

            // Flex ducts
            if (ductsFlex.Count > 0)
            {
                using (Transaction tr = new Transaction(doc, "Обработка гибких воздуховодов"))
                {
                    logger.LogInfo("Начало обработки гибких воздуховодов", docName);
                    tr.Start();

                    foreach (Element flexDuct in ductsFlex)
                    {
                        try
                        {
                            if (RevitUtils.CheckElement(flexDuct))
                            {
                                continue;
                            }
                            // Получение сеществующих параметров
                            string flexSize = flexDuct.get_Parameter(BuiltInParameter.RBS_CURVE_DIAMETER_PARAM).AsValueString();
                            Element flexType = doc.GetElement(flexDuct.GetTypeId());

                            // Генерация новых значений
                            string newName = $"{flexType.get_Parameter(BuiltInParameter.ALL_MODEL_TYPE_COMMENTS).AsValueString()} ⌀{flexSize}";
                            string sign = $"⌀{flexSize}";
                            double count = 1;
                            string Unit = flexDuct.get_Parameter(dictionaryGUID.ADSKUnit).AsValueString();
                            ForgeTypeId countUnitts = UnitTypeId.Meters;
                            if (Unit == "м")
                            {
                                countUnitts = UnitTypeId.Meters;
                                Parameter countParam = flexDuct.get_Parameter(BuiltInParameter.CURVE_ELEM_LENGTH);
                                count = UnitUtils.ConvertFromInternalUnits(countParam.AsDouble(), countParam.GetUnitTypeId()) * koef;
                            }

                            // Обработка параметров
                            RevitUtils.SetParameterValue(flexDuct, dictionaryGUID.ADSKName, newName);
                            RevitUtils.SetParameterValue(flexDuct, dictionaryGUID.ADSKSign, sign);
                            RevitUtils.SetParameterValue(flexDuct, dictionaryGUID.ADSKCount, count);
                        }
                        catch (Exception ex)
                        {
                            testLog.Concat($"Ошибка при обработке гибких воздуховодов {ex}");
                            //TaskDialog.Show("Ошибка", $"Ошибка при обработке гибких воздуховодов {ex.ToString()}");
                        }
                    }
                    tr.Commit();
                    logger.LogInfo("Завершение обработки гибких воздуховодов", docName);
                }
            }

            // Pipes
            if (pipes.Count > 0)
            {
                using (Transaction tr = new Transaction(doc, "Обработка трубопроводов"))
                {
                    logger.LogInfo("Начало обработки трубопроводов", docName);
                    tr.Start();

                    foreach (Element pipe in pipes)
                    {
                        try
                        {
                            if (RevitUtils.CheckElement(pipe))
                            {
                                continue;
                            }
                            // Получение существующих параметров 
                            ForgeTypeId pipeUnits = UnitTypeId.Meters;
                            ForgeTypeId thicknessUnits = UnitTypeId.Millimeters;
                            string pipeMark = RevitUtils.GetSharedParameterValue(pipe, dictionaryGUID.ADSKMark);
                            string pipeName = doc.GetElement(pipe.GetTypeId()).get_Parameter(BuiltInParameter.ALL_MODEL_TYPE_COMMENTS).AsValueString();
                            Parameter outSideDimParam = pipe.get_Parameter(BuiltInParameter.RBS_PIPE_OUTER_DIAMETER);
                            Parameter inSideDimParam = pipe.get_Parameter(BuiltInParameter.RBS_PIPE_INNER_DIAM_PARAM);
                            string diameter = pipe.get_Parameter(BuiltInParameter.RBS_PIPE_DIAMETER_PARAM).AsValueString();
                            double outSideDiameter = outSideDimParam.AsDouble();
                            string stringOusideDiam = outSideDimParam.AsValueString();
                            double inSideDiameter = inSideDimParam.AsDouble();

                            // Генерация новых значений
                            string newSign = "Проверить параметр ADSK_Марка";
                            double countThickness = (outSideDiameter - inSideDiameter) / 2.0;
                            string thickness = Math.Round(UnitUtils.ConvertFromInternalUnits(countThickness, thicknessUnits), 2).ToString();
                            double newCount = pipe.get_Parameter(BuiltInParameter.CURVE_ELEM_LENGTH).AsDouble();
                            string marker = "⌀";

                            if (pipeMark == "ГОСТ Р 52134-2003"
                                || pipeMark == "ГОСТ 52134-2003"
                                || pipeMark == "ГОСТ 8732-78"
                                || pipeMark == "ГОСТ 10704-91"
                                || pipeMark == "ГОСТ Р 70628.2-2023"
                                || pipeMark == "ГОСТ Р 54475-2011"
                                || pipeMark == "ГОСТ 18599-2001"
                                || pipeMark == "ГОСТ 32414-2013"
                                || pipeMark == "ГОСТ 32415-2013"
                                || pipeMark == "RAUTITAN flex"
                                || pipeMark == "RAUTITAN pink") 
                                newSign = $"{marker}{stringOusideDiam}x{thickness}";

                            else if (pipeMark == "ГОСТ 3262-75"
                                || pipeMark == "ЕN 12735-2")
                                newSign = $"{marker}{stringOusideDiam}x{thickness}";

                            else if (pipeMark == "ГОСТ Р 52318-2005" 
                                || pipeMark == "ГОСТ Р 52318-2005")
                                newSign = $"{marker}{diameter},0x{thickness}";

                            else
                                newSign = $"{marker}{diameter},0x{thickness}";

                            string newName = $"{pipeName} {newSign}";

                            // Обработка параметров 
                            RevitUtils.SetParameterValue(pipe, dictionaryGUID.ADSKName, newName);
                            RevitUtils.SetParameterValue(pipe, dictionaryGUID.ADSKSign, newSign);
                            RevitUtils.SetParameterValue(pipe, dictionaryGUID.ADSKThicknes, thickness);
                            RevitUtils.SetParameterValue(pipe, dictionaryGUID.ADSKCount, newCount * koef);

                        }

                        catch (Exception ex)
                        {
                            testLog.Concat($"Ошибка при обработке трубопроводов {pipe.Id} {ex}");
                            //TaskDialog.Show("Ошибка", $"Ошибка при обработке трубопроводов {ex.ToString()}");
                        }
                    }

                    tr.Commit();
                    logger.LogInfo("Завершение обработки трубопроводов", docName);
                }
            }

            // Flex pipes
            if (pipesFlex.Count > 0)
            {
                using( Transaction tr = new Transaction(doc, "Обработка гибких трубопроводов"))
                {
                    logger.LogInfo("Начало обработки гибких трубопроводов", docName);
                    tr.Start();

                    foreach (var flex in pipesFlex)
                    {
                        try
                        {
                            if (RevitUtils.CheckElement(flex))
                            {
                                continue;
                            }
                            // Получение существующих параметров 
                            Element flexType = doc.GetElement(flex.GetTypeId());
                            string flesTypeComment = flexType.get_Parameter(BuiltInParameter.ALL_MODEL_TYPE_COMMENTS).AsValueString();
                            string flexUnit = flexType.get_Parameter(dictionaryGUID.ADSKUnit).AsValueString();
                            string marker = "⌀";
                            double diameter = UnitUtils.ConvertFromInternalUnits(flex.get_Parameter(BuiltInParameter.RBS_PIPE_DIAMETER_PARAM).AsDouble(), UnitTypeId.Millimeters);
                            double count = 0;

                            // Генерация новых значений
                            string newName = $"{flesTypeComment} {marker}{diameter}";
                            if (flexUnit == "м" || flexUnit == "m")
                            {
                                count = UnitUtils.ConvertFromInternalUnits(flex.get_Parameter(BuiltInParameter.CURVE_ELEM_LENGTH).AsDouble(), UnitTypeId.Millimeters) / 1000 * koef;
                            }
                            else if (flexUnit.Contains("шт"))
                            {
                                count = 1.0;
                            }

                            // Обработка параметров 
                            RevitUtils.SetParameterValue(flex, dictionaryGUID.ADSKName, newName);
                            RevitUtils.SetParameterValue(flex, dictionaryGUID.ADSKCount, count);

                        }
                        catch (Exception ex)
                        {
                            testLog.Concat($"Ошибка при обработке гибких трубопроводов {flex.Id} {ex}");
                            //TaskDialog.Show("Ошибка", $"Ошибка при обработке трубопроводов {ex.ToString()}");
                        }
                    }

                    tr.Commit();
                    logger.LogInfo("Завершение обработки гибких трубопроводов", docName);
                }
            }

            // Pipes insulation
            if (pipeInsulation.Count > 0)
            {
                using (Transaction tr = new Transaction(doc, "Обработка изоляции трубопроводов"))
                {
                    logger.LogInfo("Начало обработки изоляции трубопроводов", docName);
                    tr.Start();
                    foreach (Element insulation in pipeInsulation)
                    {
                        try
                        {
                            if (insulation != null)
                            {
                                if (RevitUtils.CheckElement(insulation))
                                {
                                    continue;
                                }
                                
                                // Получение существующих параметров 
                                InsulationLiningBase insLinBase = insulation as InsulationLiningBase;
                                Element host = doc.GetElement(insLinBase.HostElementId);
                                Element insType = doc.GetElement(insulation.GetTypeId());
                                bool isCategotyPypeAcc = host.Category.Id.IntegerValue == ((int)BuiltInCategory.OST_PipeAccessory);
                                string insUnit = RevitUtils.GetSharedParameterValue(insulation, dictionaryGUID.ADSKUnit);
                                string insMark = insulation.get_Parameter(dictionaryGUID.ADSKMark).AsValueString();
                                string insTypeComment = insType.get_Parameter(BuiltInParameter.ALL_MODEL_TYPE_COMMENTS).AsValueString();
                                string hostFabric = RevitUtils.GetSharedParameterValue(host, dictionaryGUID.ADSKFabricName);
                                string hostOutsideDiam = host.get_Parameter(BuiltInParameter.RBS_PIPE_OUTER_DIAMETER).AsValueString();
                                double thickness = UnitUtils.ConvertFromInternalUnits(insulation.get_Parameter(BuiltInParameter.RBS_INSULATION_THICKNESS_FOR_PIPE).AsDouble(), UnitTypeId.Millimeters);
                                double count = 1;

                                // Генерация новых значений
                                string newName = $"{insTypeComment} толщиной {thickness}";

                                if (isCategotyPypeAcc)
                                {
                                    count = 1.0;
                                    RevitUtils.SetParameterValue(insulation, dictionaryGUID.ADSKCount, count);
                                }
                                else
                                {
                                    ForgeTypeId countUnits = UnitTypeId.Meters;
                                    count = UnitUtils.ConvertFromInternalUnits(insulation.get_Parameter(BuiltInParameter.CURVE_ELEM_LENGTH).AsDouble(), UnitTypeId.Meters);
                                    if (insUnit == "м²")
                                    {
                                        countUnits = UnitTypeId.SquareMeters;
                                        count = UnitUtils.ConvertFromInternalUnits(insulation.get_Parameter(BuiltInParameter.RBS_CURVE_SURFACE_AREA).AsDouble(), UnitTypeId.SquareMeters);
                                        RevitUtils.SetParameterValue(insulation, dictionaryGUID.ADSKCount, count);
                                    }
                                    else
                                    {
                                        RevitUtils.SetParameterValue(insulation, dictionaryGUID.ADSKCount, count);
                                    }
                                }

                                if (insMark.IndexOf("k-flex st", StringComparison.OrdinalIgnoreCase) >= 0
                                    || insMark.IndexOf("Energocell HT", StringComparison.OrdinalIgnoreCase) >= 0 
                                    || insTypeComment.IndexOf("k-flex st", StringComparison.OrdinalIgnoreCase) >= 0 
                                    || insTypeComment.IndexOf("Energocell HT", StringComparison.OrdinalIgnoreCase) >= 0)
                                    newName = $"{insTypeComment} {thickness}{CalculateRightInsulationName(hostOutsideDiam, hostFabric)} для трубопровода {TransformFabric(hostFabric)}";

                                if (insMark.IndexOf("PE Compact", StringComparison.OrdinalIgnoreCase) >= 0 || insTypeComment.IndexOf("PE Compact", StringComparison.OrdinalIgnoreCase) >= 0)
                                    newName = $"{insTypeComment} {thickness}{CalculateRightInsulationName(hostOutsideDiam, hostFabric)} из вспененного полиэтилена с наружным слоем из полимерной армирующей пленки для трубопровода {TransformFabric(hostFabric)}";

                                // Обработка параметров 
                                RevitUtils.SetParameterValue(insulation, dictionaryGUID.ADSKName, newName);
                                RevitUtils.SetParameterValue(insulation, dictionaryGUID.ADSKThicknes, UnitUtils.ConvertToInternalUnits(thickness, UnitTypeId.Millimeters) * koef);
                            }
                        }
                        catch (Exception ex)
                        {
                            testLog.Concat($"Ошибка при обработке изоляции трубопроводов {ex}");
                            //TaskDialog.Show("Ошибка", $"Ошибка при обработке изоляции трубопроводов {ex.ToString()}");
                        }
                    }
                    tr.Commit();
                    logger.LogInfo("Завершение обработки изоляции трубопроводов", docName);
                }
            }

            // Finish
            if (testLog.Length > 0)
            {
                //TaskDialog.Show("Ошибки", $"{testLog}");
                return Result.Succeeded;
            }
            else
            {
                //TaskDialog.Show("Готово", "Параметры для спецификации заполнены!");
                logger.LogInfo("Завершение заполнение параметров для специфицкации", docName);
                return Result.Succeeded; // Подумать, может можно не заканчивать, а пропустить?
            }
        }

        /// <summary>
        /// Метод для нахождения толщины стенки воздуховодов и их соединительных деталей
        /// </summary>
        /// <param newName="elem"></param>
        /// <param newName="ductSize"></param>
        /// <param newName="thickness"></param>
        private void GetDuctThickness(Element elem, ref string ductSize, ref string thickness)
        {
            var logger = ATP_App.GetService<ILoggerService>();
            try
            {
                //  Определение типа воздуховода (вынесено в отдельные методы для читаемости)
                bool isRectangular = IsRectangularDuct(elem);
                bool isFitting = IsDuctFitting(elem);

                // Получение основного размера
                int mainSize = GetDuctSize(elem, isRectangular, isFitting, ref ductSize);

                if (mainSize <= 0)
                {
                    //TaskDialog.Show("Ошибка", $"Не удалось определить размер {elem.Category.Name} - ID:{elem.Id}");
                    logger.LogWarning($"Не удалось определить размер {elem.Category.Name} - ID:{elem.Id}", elem.Document.Title);
                    thickness = "0.0";
                    return;
                }

                // Определение толщины стенки
                thickness = CalculateThickness(elem, isRectangular, mainSize);
            }
            catch (Exception ex)
            {
                //TaskDialog.Show("Критическая ошибка", $"Ошибка при расчете толщины: {ex.Message}");
                logger.LogError($"Ошибка при расчете толщины: {ex.Message}", elem.Document.Title);
                thickness = "0.0";
            }
        }

        /// <summary>
        /// Вспомогательный метод, для определения прямоугольный ли воздуховод.
        /// </summary>
        /// <param newName="elem"></param>
        /// <returns></returns>
        private bool IsRectangularDuct(Element elem)
        {
            string familyName = elem.get_Parameter(BuiltInParameter.RBS_CALCULATED_SIZE).AsValueString() ?? "";
            return familyName.IndexOf("x", StringComparison.OrdinalIgnoreCase) >= 0 ||
                   familyName.IndexOf("х", StringComparison.OrdinalIgnoreCase) >= 0;
        }

        /// <summary>
        /// Вспомогательный метод, для определения является ли элемент соединительной деталью.
        /// </summary>
        /// <param newName="elem"></param>
        /// <returns></returns>
        private bool IsDuctFitting(Element elem)
        {
            string categoryName = elem.Category?.Name ?? "";
            return categoryName.IndexOf("Fitting", StringComparison.OrdinalIgnoreCase) >= 0 ||
                   categoryName.IndexOf("детали", StringComparison.OrdinalIgnoreCase) >= 0;
        }

        /// <summary>
        /// Вспомогательный метод, для опредение корректного расметра воздуховода
        /// </summary>
        /// <param newName="elem"></param>
        /// <param newName="isRectangular"></param>
        /// <param newName="isFitting"></param>
        /// <param newName="ductSize"></param>
        /// <returns></returns>
        private int GetDuctSize(Element elem, bool isRectangular, bool isFitting, ref string ductSize)
        {
            if (isFitting)
            {
                string sizeString = elem.get_Parameter(BuiltInParameter.RBS_CALCULATED_SIZE).AsValueString() ?? "";
                int size = ParseSize(elem, sizeString, isRectangular, ref ductSize);
                return size;
            }
            else
            {
                if (isRectangular)
                {
                    Int32.TryParse(elem.get_Parameter(BuiltInParameter.RBS_CURVE_WIDTH_PARAM).AsValueString(), out int width);
                    Int32.TryParse(elem.get_Parameter(BuiltInParameter.RBS_CURVE_HEIGHT_PARAM).AsValueString(), out int height);
                    ductSize = $"{Math.Min(width, height)}x{Math.Max(width, height)}";
                    return Math.Max(width, height);
                }
                else
                {
                    Int32.TryParse(elem.get_Parameter(BuiltInParameter.RBS_CURVE_DIAMETER_PARAM).AsValueString(), out int diameter);
                    ductSize = $"⌀{diameter}";
                    return diameter;
                }
            }
        }

        /// <summary>
        /// Вспомогательный метод для обработки 
        /// </summary>
        /// <param newName="sizeString"></param>
        /// <param newName="isRectangular"></param>
        /// <param newName="ductSize"></param>
        /// <returns></returns>
        private int ParseSize(Element elem, string sizeString, bool isRectangular, ref string ductSize)
        {
            if (isRectangular)
            {
                var numbers = sizeString.Trim().Replace("-", "-").Split('-')
                    .Where(part => (part.Contains('x') || part.Contains('х')) && !part.Contains('⌀'))
                    .SelectMany(part => part.Replace('х', 'x').Replace("x", "x").Split('x'))
                    .Select(numStr => int.TryParse(RevitUtils.CleanSizeString(numStr.Trim()), out int num) ? num : 0)
                    .Where(num => num > 0)
                    .ToList();

                if (numbers.Count == 0) return 0;

                ProcessSpecialCases(elem, sizeString, ref ductSize);
                return numbers.Max();
            }
            else
            {
                var diameters = sizeString.Split('-')
                    .Where(part => part.Contains('⌀'))
                    .Select(part => int.TryParse(RevitUtils.CleanSizeString(part.Replace("⌀", "").Trim()), out int num) ? num : 0)
                    .Where(num => num > 0)
                    .ToList();

                if (diameters.Count == 0) return 0;

                ProcessSpecialCases(elem, sizeString, ref ductSize);
                return diameters.Max();
            }
        }

        /// <summary>
        /// Вспомогальный метод, для установки отдельного размера для отводов и врезок
        /// </summary>
        /// <param newName="elem"></param>
        /// <param newName="sizeString"></param>
        /// <param newName="ductSize"></param>
        private void ProcessSpecialCases(Element elem, string sizeString, ref string ductSize)
        {
            string familyName = elem.get_Parameter(BuiltInParameter.ELEM_FAMILY_PARAM).AsValueString();
            if (familyName.IndexOf("Отвод", StringComparison.OrdinalIgnoreCase) >= 0 ||
                familyName.IndexOf("Врезка", StringComparison.OrdinalIgnoreCase) >= 0)
            {
                ductSize = sizeString.Split('-').FirstOrDefault();
            }
        }

        /// <summary>
        /// Вспомогательный метод для подсчёта толщины стелки воздуховода в соответствии с СП 7 и СП 60
        /// </summary>
        /// <param newName="elem"></param>
        /// <param newName="isRectangular"></param>
        /// <param newName="mainSize"></param>
        /// <returns></returns>
        private string CalculateThickness(Element elem, bool isRectangular, int mainSize)
        {
            bool isEI = elem.get_Parameter(BuiltInParameter.RBS_REFERENCE_INSULATION_TYPE)?.Element?.Name.Contains("EI") ?? false;
            bool isWelded = elem.Name.IndexOf("Сварной", StringComparison.OrdinalIgnoreCase) >= 0;

            if (isEI && isWelded) return "1.5";

            if (isRectangular)
            {
                if (isEI) return mainSize <= 1000 ? "0.8" : "0.9";

                if (mainSize <= 250) return "0.5";
                if (mainSize <= 1000) return "0.7";
                return "0.9";
            }
            else
            {
                if (isEI)
                {
                    if (mainSize <= 800) return "0.8";
                    if (mainSize <= 1250) return "1.0";
                    if (mainSize <= 1600) return "1.2";
                    return "1.4";
                }
                else
                {
                    if (mainSize <= 200) return "0.5";
                    if (mainSize <= 450) return "0.6";
                    if (mainSize <= 800) return "0.7";
                    if (mainSize <= 1250) return "1.0";
                    if (mainSize <= 1600) return "1.2";
                    return "1.4";
                }
            }
        }

        /// <summary>
        /// Вспомогательный метод для генерации части наименования изоляции трубопровода, на основе параметра ADSK_Материал наименование его основы
        /// </summary>
        /// <param name="fabric"></param>
        /// <returns></returns>
        private string TransformFabric(string fabric)
        {
            switch (fabric)
            {
                case "Сшитый полиэтилен": return "из сшитого полиэтилена";
                case "Сталь": return "из оцинкованной стали";
                case "Чугун": return "чугунного";
                case "Полиэтилен": return "из полиэтилена";
                case "Полипропилен": return "из полипропилена";
                default:return "Не заполнен ADSK_Материал наименование у основы";
            }
        }

        /// <summary>
        /// Вспомогательный метод для подбора части наименования для трубопроводов марки PE Compact, K-Fles ST и Energoflex
        /// </summary>
        /// <param name="outSide"></param>
        /// <param name="hostFabric"></param>
        /// <returns></returns>
        private string CalculateRightInsulationName(string outSide, string hostFabric)
        {
            string insulDiam = "Не удалось подобрать диаметр трубопровода под типоразмер изоляцию";
            Double.TryParse(outSide, out double outsideDiam);

            if (outsideDiam < 18)
                insulDiam = "18";
            else if (outsideDiam >= 18 && outsideDiam < 22)
                insulDiam = "22";
            else if (outsideDiam >= 22 && outsideDiam < 28)
                insulDiam = "28";
            else if (outsideDiam >= 28 && outsideDiam < 35)
                insulDiam = "35";
            else if (outsideDiam >= 35 && outsideDiam < 42)
                insulDiam = "42";
            else if (outsideDiam >= 42 && outsideDiam < 48)
                insulDiam = "48";
            else if (outsideDiam >= 48 && outsideDiam < 54)
                insulDiam = "54";
            else if (outsideDiam >= 54 && outsideDiam < 60)
                insulDiam = "60";
            else if (outsideDiam >= 60 && outsideDiam < 76)
                insulDiam = "76";
            else if (outsideDiam >= 76 && outsideDiam < 89)
                insulDiam = "89";
            else if (outsideDiam >= 89 && outsideDiam < 102)
                insulDiam = "102";
            else if (outsideDiam >= 102 && outsideDiam < 108)
                insulDiam = "108";
            else if (outsideDiam >= 108 && outsideDiam < 114)
                insulDiam = "114";
            else if (outsideDiam >= 114 && outsideDiam < 125)
                insulDiam = "125";
            else if (outsideDiam >= 125 && outsideDiam < 133)
                insulDiam = "133";
            else if (outsideDiam >= 133 && outsideDiam < 140)
                insulDiam = "140";
            else if (outsideDiam >= 140 && outsideDiam < 160)
                insulDiam = "160";

            string newName = $"x{insulDiam}";
            return newName;
        }
    }
}