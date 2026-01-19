using ATP_Common_Plugin.Services;
using ATP_Common_Plugin.Utils;
using Autodesk.Revit.Attributes;
using Autodesk.Revit.DB;
using Autodesk.Revit.DB.Mechanical;
using Autodesk.Revit.UI;
using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;
using System.Net;
using System.Text.RegularExpressions;

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
                    logger.LogInfo($"Начало обработки воздуховодов. Найдено {ducts.Count} элементов.", docName);
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
                            string ductSize = duct.get_Parameter(BuiltInParameter.RBS_CALCULATED_SIZE)?.AsValueString() ?? "";
                            Parameter paramLength = duct.get_Parameter(BuiltInParameter.CURVE_ELEM_LENGTH);
                            string thickness = "0.9";
                            GetDuctThickness(duct, ref ductSize, ref thickness);

                            Element ductType = doc.GetElement(duct.GetTypeId());
                            if (ductType == null)
                            {
                                logger.LogWarning($"Тип воздуховода для элемента {duct.Id} не найден", docName);
                                continue;
                            }

                            string ductUnits = ductType.get_Parameter(dictionaryGUID.ADSKUnit)?.AsValueString() ?? "";

                            // qtyNumber: длина (м) или площадь (м²) по единицам
                            double qtyNumber;
                            if (ductUnits == "м²")
                            {
                                // площадь поверхности
                                var areaParam = duct.get_Parameter(BuiltInParameter.RBS_CURVE_SURFACE_AREA);
                                double areaInt = areaParam?.AsDouble() ?? 0; // внутр. ед. (кв. фут)
                                qtyNumber = UnitUtils.ConvertFromInternalUnits(areaInt, UnitTypeId.SquareMeters);
                            }
                            else
                            {
                                // длина
                                double lenInt = paramLength?.AsDouble() ?? 0; // внутр. ед. (фут)
                                qtyNumber = UnitUtils.ConvertFromInternalUnits(lenInt, UnitTypeId.Meters);
                            }

                            // Генерация новых значений
                            string typeComments = ductType.get_Parameter(BuiltInParameter.ALL_MODEL_TYPE_COMMENTS)?.AsValueString() ?? "";
                            string newName = $"{typeComments} {ductSize} δ= {thickness}";
                            double.TryParse(thickness.Trim(), out double newThickness);

                            // Обработка параметров 
                            RevitUtils.SetParameterValue(duct, dictionaryGUID.ADSKName, newName);
                            RevitUtils.SetParameterValue(duct, dictionaryGUID.ADSKSign, ductSize);

                            // ADSK_Количество — уважаем собственные единицы параметра
                            var countParam = duct.get_Parameter(dictionaryGUID.ADSKCount);
                            if (countParam != null && countParam.StorageType == StorageType.Double)
                            {
                                var u = countParam.GetUnitTypeId();
                                double toSet;
                                if (u == UnitTypeId.Meters)
                                    toSet = UnitUtils.ConvertToInternalUnits(qtyNumber * koefDucts, UnitTypeId.Meters);
                                else if (u == UnitTypeId.SquareMeters)
                                    toSet = UnitUtils.ConvertToInternalUnits(qtyNumber * koefDucts, UnitTypeId.SquareMeters);
                                else
                                    toSet = qtyNumber * koefDucts; // безразмерный

                                RevitUtils.SetParameterValue(duct, dictionaryGUID.ADSKCount, toSet);
                            }

                            // Толщина стенки (мм → футы внутр. ед.)
                            RevitUtils.SetParameterValue(duct, dictionaryGUID.ADSKThicknes, UnitUtils.ConvertToInternalUnits(newThickness, UnitTypeId.Millimeters));
                        }
                        catch (Exception ex)
                        {
                            logger.LogError($"Ошибка при обработке воздуховодов {duct.Id} {ex}", docName);
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
                    logger.LogInfo($"Начало обработки соединительных деталей воздуховодов. Найдено {ductFittings.Count} элементов.", docName);
                    tr.Start();

                    foreach (Element ductFitting in ductFittings)
                    {
                        try
                        {
                            if (RevitUtils.CheckElement(ductFitting))
                            {
                                continue;
                            }

                            // Получение существующих параметров
                            string ductSize = ductFitting.get_Parameter(BuiltInParameter.RBS_CALCULATED_SIZE)?.AsValueString() ?? "";
                            string thickness = "0.9";
                            GetDuctThickness(ductFitting, ref ductSize, ref thickness);

                            Element ductFittingType = doc.GetElement(ductFitting.GetTypeId());
                            if (ductFittingType == null)
                            {
                                logger.LogWarning($"Тип фасонной детали для элемента {ductFitting.Id} не найден", docName);
                                continue;
                            }

                            string ductFittingUnits = ductFitting.get_Parameter(dictionaryGUID.ADSKUnit)?.AsValueString() ?? "";

                            double thicknessMmParsed = 0.0;
                            double.TryParse(thickness.Trim(), NumberStyles.Any, CultureInfo.InvariantCulture, out thicknessMmParsed);
                            double thicknessMm2 = Math.Round(thicknessMmParsed, 2, MidpointRounding.AwayFromZero);
                            string thicknessStr = thicknessMm2.ToString("0.##", CultureInfo.InvariantCulture);

                            // Генерация новых значений 
                            if (ductFittingUnits == "м²")
                            {
                                Parameter param = ductFitting.get_Parameter(dictionaryGUID.ADSKSizeArea);
                                if (param != null)
                                {
                                    ForgeTypeId unitType = param.GetUnitTypeId();
                                    double ductFittingValue = UnitUtils.ConvertFromInternalUnits(param.AsDouble(), unitType);
                                    RevitUtils.SetParameterValue(ductFitting, dictionaryGUID.ADSKCount, ductFittingValue);
                                }
                            }

                            // Имя с углом для отводов (если параметр угла есть)
                            string baseTypeComment = ductFittingType.get_Parameter(BuiltInParameter.ALL_MODEL_TYPE_COMMENTS)?.AsValueString() ?? "";
                            string famName = ductFitting.get_Parameter(BuiltInParameter.ELEM_FAMILY_PARAM)?.AsValueString() ?? "";

                            string newName = $"{baseTypeComment} {ductSize}  δ= {thicknessStr}";
                            if (famName.IndexOf("Отвод", StringComparison.OrdinalIgnoreCase) >= 0)
                            {
                                // !!! при необходимости укажи верный GUID параметра угла
                                Parameter pAngle = ductFitting.get_Parameter(dictionaryGUID.ADSKSizeAngle);
                                if (pAngle != null && pAngle.StorageType == StorageType.Double)
                                {
                                    // Внутренние единицы угла — радианы. Конвертим в градусы.
                                    double angDeg = UnitUtils.ConvertFromInternalUnits(pAngle.AsDouble(), UnitTypeId.Degrees);
                                    int angInt = (int)Math.Round(angDeg, 0);
                                    newName = $"{baseTypeComment} {angInt}° {ductSize}  δ= {thickness}";
                                }
                            }

                            // Запись параметров
                            RevitUtils.SetParameterValue(ductFitting, dictionaryGUID.ADSKName, newName);
                            RevitUtils.SetParameterValue(ductFitting, dictionaryGUID.ADSKSign, ductSize);
                            RevitUtils.SetParameterValue(ductFitting, dictionaryGUID.ADSKThicknes,
                                UnitUtils.ConvertToInternalUnits(thicknessMm2, UnitTypeId.Millimeters));
                        }
                        catch (Exception ex)
                        {
                            testLog += $"Ошибка при обработке соединительных деталей воздуховодов {ductFitting.Id} {ex}\n";
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
                    logger.LogInfo($"Начало обработки изоляции воздуховодов. Найдено {ductInsulation.Count} элементов.", docName);
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
                            if (insulationType == null)
                            {
                                logger.LogWarning($"Тип изоляции для элемента {insulation.Id} не найден", docName);
                                continue;
                            }

                            InsulationLiningBase insBase = insulation as InsulationLiningBase;
                            if (insBase == null)
                            {
                                logger.LogWarning($"Элемент {insulation.Id} не является InsulationLiningBase", docName);
                                continue;
                            }

                            Element insulationHost = doc.GetElement(insBase.HostElementId);
                            if (insulationHost == null)
                            {
                                logger.LogWarning($"Изоляция {insulation.Id} потеряла основу (host)", docName);
                                continue;
                            }

                            bool isHostFitting = insulationHost.Category.Name.IndexOf("Fittings", StringComparison.OrdinalIgnoreCase) >= 0 || insulationHost.Category.Name.Contains("Cоединительные");
                            string insulationUnit = RevitUtils.GetSharedParameterValue(insulation, dictionaryGUID.ADSKUnit) ?? "";
                            double thickness = insulation.get_Parameter(BuiltInParameter.RBS_INSULATION_THICKNESS_FOR_DUCT)?.AsDouble() ?? 0;

                            // Генерация новых значений
                            string typeComments = insulationType.get_Parameter(BuiltInParameter.ALL_MODEL_TYPE_COMMENTS)?.AsValueString() ?? "";
                            string newName = $"{typeComments} толшиной {UnitUtils.ConvertFromInternalUnits(thickness, UnitTypeId.Millimeters)} мм";
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

                            var countParamIns = insulation.get_Parameter(dictionaryGUID.ADSKCount);
                            if (countParamIns != null && countParamIns.StorageType == StorageType.Double)
                            {
                                var u = countParamIns.GetUnitTypeId();
                                double toSet;
                                if (u == UnitTypeId.Meters)
                                    toSet = UnitUtils.ConvertToInternalUnits(count * koef, UnitTypeId.Meters);
                                else if (u == UnitTypeId.SquareMeters)
                                    toSet = UnitUtils.ConvertToInternalUnits(count * koef, UnitTypeId.SquareMeters);
                                else
                                    toSet = count * koef;
                                RevitUtils.SetParameterValue(insulation, dictionaryGUID.ADSKCount, toSet);
                            }

                        }
                        catch (Exception ex)
                        {
                            testLog += $"Ошибка при обработке изоляции воздуховодов {ex}\n";
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
                    logger.LogInfo($"Начало обработки гибких воздуховодов. Найдено {ductsFlex.Count} элементов.", docName);
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
                            string flexSize = flexDuct.get_Parameter(BuiltInParameter.RBS_CURVE_DIAMETER_PARAM)?.AsValueString() ?? "";
                            Element flexType = doc.GetElement(flexDuct.GetTypeId());
                            if (flexType == null)
                            {
                                logger.LogWarning($"Тип гибкого воздуховода для элемента {flexDuct.Id} не найден", docName);
                                continue;
                            }

                            // Генерация новых значений
                            string typeComments = flexType.get_Parameter(BuiltInParameter.ALL_MODEL_TYPE_COMMENTS)?.AsValueString() ?? "";
                            string newName = $"{typeComments} ⌀{flexSize}";
                            string sign = $"⌀{flexSize}";
                            double count = 1;
                            string Unit = flexDuct.get_Parameter(dictionaryGUID.ADSKUnit)?.AsValueString() ?? "";
                            ForgeTypeId countUnitts = UnitTypeId.Meters;
                            if (Unit == "м")
                            {
                                countUnitts = UnitTypeId.Meters;
                                Parameter countParam = flexDuct.get_Parameter(BuiltInParameter.CURVE_ELEM_LENGTH);
                                if (countParam != null)
                                    count = UnitUtils.ConvertFromInternalUnits(countParam.AsDouble(), countParam.GetUnitTypeId()) * koef;
                            }

                            // Обработка параметров
                            RevitUtils.SetParameterValue(flexDuct, dictionaryGUID.ADSKName, newName);
                            RevitUtils.SetParameterValue(flexDuct, dictionaryGUID.ADSKSign, sign);

                            var countParamFlex = flexDuct.get_Parameter(dictionaryGUID.ADSKCount);
                            if (countParamFlex != null && countParamFlex.StorageType == StorageType.Double)
                            {
                                var u = countParamFlex.GetUnitTypeId();
                                double toSet = (u == UnitTypeId.Meters)
                                    ? UnitUtils.ConvertToInternalUnits(count, UnitTypeId.Meters)
                                    : count;
                                RevitUtils.SetParameterValue(flexDuct, dictionaryGUID.ADSKCount, toSet);
                            }
                        }
                        catch (Exception ex)
                        {
                            testLog += $"Ошибка при обработке гибких воздуховодов {ex}\n";
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
                    logger.LogInfo($"Начало обработки трубопроводов. Найдено {pipes.Count} элементов.", docName);
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
                            string pipeMark = RevitUtils.GetSharedParameterValue(pipe, dictionaryGUID.ADSKMark) ?? "";
                            Element pipeType = doc.GetElement(pipe.GetTypeId());
                            if (pipeType == null)
                            {
                                logger.LogWarning($"Тип трубы для элемента {pipe.Id} не найден", docName);
                                continue;
                            }

                            string pipeName = pipeType.get_Parameter(BuiltInParameter.ALL_MODEL_TYPE_COMMENTS)?.AsValueString() ?? "";

                            Parameter outSideDimParam = pipe.get_Parameter(BuiltInParameter.RBS_PIPE_OUTER_DIAMETER);
                            Parameter inSideDimParam = pipe.get_Parameter(BuiltInParameter.RBS_PIPE_INNER_DIAM_PARAM);

                            double odInt = outSideDimParam?.AsDouble() ?? 0; // внутр. ед. (фут)
                            double idInt = inSideDimParam?.AsDouble() ?? 0;  // внутр. ед. (фут)
                            double lenInt = pipe.get_Parameter(BuiltInParameter.CURVE_ELEM_LENGTH)?.AsDouble() ?? 0;

                            // === Толщина стенки: единый расчёт и округление до 2 знаков ===
                            double wallInt = (odInt - idInt) / 2.0; // внутр. ед.
                            double wallMmExact = UnitUtils.ConvertFromInternalUnits(wallInt, UnitTypeId.Millimeters);
                            double wallMm2 = Math.Round(wallMmExact, 2, MidpointRounding.AwayFromZero);
                            string wallStr = wallMm2.ToString("0.##", CultureInfo.InvariantCulture);

                            // Доп. величины
                            double odMm = UnitUtils.ConvertFromInternalUnits(odInt, UnitTypeId.Millimeters);
                            double dnMm = UnitUtils.ConvertFromInternalUnits(
                                pipe.get_Parameter(BuiltInParameter.RBS_PIPE_DIAMETER_PARAM).AsDouble(),
                                UnitTypeId.Millimeters);
                            double lenMeters = UnitUtils.ConvertFromInternalUnits(lenInt, UnitTypeId.Meters) * koef;

                            string odStr = odMm.ToString("0.##", CultureInfo.InvariantCulture);
                            string dnStr = dnMm.ToString("0.##", CultureInfo.InvariantCulture);
                            const string marker = "⌀";

                            bool isODxE =
                                pipeMark == "ГОСТ 8732-78"         // бесшовные стальные → OD×e
                                    || pipeMark == "ГОСТ 10704-91"        // электросварные → OD×e
                                    || pipeMark == "ЕN 12735-2"           // медные ACR → OD×e
                                    || pipeMark == "ГОСТ Р 52134-2003"    // термопласты → OD×e (часто с SDR/PN)
                                    || pipeMark == "ГОСТ 18599-2001"      // ПЭ → OD×e (часто с SDR)
                                    || pipeMark == "ГОСТ 32414-2013"      // ПП → OD×e
                                    || pipeMark == "ГОСТ 32415-2013"      // (если используешь — аналогично ПП) → OD×e
                                    || pipeMark == "RAUTITAN flex"        // бренд (PP-R/PE-Xa) — практично OD×e
                                    || pipeMark == "RAUTITAN pink";

                            bool isDN =
                                pipeMark == "ГОСТ 3262-75"         // ВГП → DN
                                    || pipeMark == "ГОСТ Р 52318-2005"   // EN 10255 аналог → DN
                                    || pipeMark == "DIN EN 877";


                            string newSign;
                            if (isODxE)
                            {
                                newSign = $"{marker}{odStr}x{wallStr}";
                            }
                            else if (isDN)
                            {
                                newSign = $"{marker}{dnStr}";
                            }
                            else
                            {
                                newSign = $"{marker}{odStr}x{wallStr}";
                            }

                            string newName = $"{pipeName} {newSign}";

                            // Запись параметров
                            RevitUtils.SetParameterValue(pipe, dictionaryGUID.ADSKName, newName);
                            RevitUtils.SetParameterValue(pipe, dictionaryGUID.ADSKSign, newSign);

                            // Числовой параметр толщины — пишем ровно ту же (округлённую) величину (мм → внутр. ед.)
                            double wallIntRounded = UnitUtils.ConvertToInternalUnits(wallMm2, UnitTypeId.Millimeters);
                            RevitUtils.SetParameterValue(pipe, dictionaryGUID.ADSKThicknes, wallIntRounded);

                            // Количество
                            Parameter countParam = pipe.get_Parameter(dictionaryGUID.ADSKCount);
                            if (countParam != null && countParam.StorageType == StorageType.Double)
                            {
                                ForgeTypeId u = countParam.GetUnitTypeId();
                                double toSet = (u == UnitTypeId.Meters)
                                    ? UnitUtils.ConvertToInternalUnits(lenMeters, UnitTypeId.Meters)
                                    : lenMeters; // безразмерный
                                RevitUtils.SetParameterValue(pipe, dictionaryGUID.ADSKCount, toSet);
                            }
                    }
                    catch (Exception ex)
                    {
                        logger.LogError($"Ошибка при обработке трубопроводов {pipe.Id} {ex}", docName);
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
                    logger.LogInfo($"Начало обработки гибких трубопроводов. Найдено {pipesFlex.Count} элементов.", docName);
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
                            if (flexType == null)
                            {
                                logger.LogWarning($"Тип гибкого трубопровода для элемента {flex.Id} не найден", docName);
                                continue;
                            }

                            string flesTypeComment = flexType.get_Parameter(BuiltInParameter.ALL_MODEL_TYPE_COMMENTS)?.AsValueString() ?? "";
                            string flexUnit = flexType.get_Parameter(dictionaryGUID.ADSKUnit)?.AsValueString() ?? "";
                            string marker = "⌀";
                            double diameter = UnitUtils.ConvertFromInternalUnits(flex.get_Parameter(BuiltInParameter.RBS_PIPE_DIAMETER_PARAM)?.AsDouble() ?? 0, UnitTypeId.Millimeters);
                            double count = 0;

                            // Генерация новых значений
                            string newName = $"{flesTypeComment} {marker}{diameter}";
                            if (flexUnit == "м" || flexUnit == "m")
                            {
                                count = UnitUtils.ConvertFromInternalUnits(flex.get_Parameter(BuiltInParameter.CURVE_ELEM_LENGTH)?.AsDouble() ?? 0, UnitTypeId.Millimeters) / 1000 * koef;
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
                            testLog += $"Ошибка при обработке гибких трубопроводов {flex.Id} {ex}\n";
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
                    logger.LogInfo($"Начало обработки изоляции трубопроводов. Найдено {pipeInsulation.Count} элементов.", docName);
                    tr.Start();
                    foreach (Element insulation in pipeInsulation)
                    {
                        try
                        {
                            if (insulation == null) continue;
                            if (RevitUtils.CheckElement(insulation)) continue;

                            // Получение существующих параметров 
                            InsulationLiningBase insLinBase = insulation as InsulationLiningBase;
                            if (insLinBase == null)
                            {
                                logger.LogWarning($"Элемент {insulation.Id} не является изоляцией", docName);
                                continue;
                            }

                            Element host = doc.GetElement(insLinBase.HostElementId);
                            if (host == null)
                            {
                                logger.LogWarning($"Изоляция {insulation.Id} потеряла основу (host)", docName);
                                continue;
                            }

                            Element insType = doc.GetElement(insulation.GetTypeId());
                            if (insType == null)
                            {
                                logger.LogWarning($"Тип изоляции для элемента {insulation.Id} не найден", docName);
                                continue;
                            }

                            bool isCategotyPypeAcc = host.Category.Id.IntegerValue == ((int)BuiltInCategory.OST_PipeAccessory);
                            string insUnit = RevitUtils.GetSharedParameterValue(insulation, dictionaryGUID.ADSKUnit) ?? "";
                            string insMark = RevitUtils.GetSharedParameterValue(insulation, dictionaryGUID.ADSKMark) ?? "";
                            string insTypeComment = insType.get_Parameter(BuiltInParameter.ALL_MODEL_TYPE_COMMENTS)?.AsString() ?? "";
                            string hostFabric = RevitUtils.GetSharedParameterValue(host, dictionaryGUID.ADSKFabricName) ?? "";

                            double hostOutsideDiamMm = UnitUtils.ConvertFromInternalUnits(
                                host.get_Parameter(BuiltInParameter.RBS_PIPE_OUTER_DIAMETER).AsDouble(), UnitTypeId.Millimeters);

                            double thicknessMm = UnitUtils.ConvertFromInternalUnits(
                                insulation.get_Parameter(BuiltInParameter.RBS_INSULATION_THICKNESS).AsDouble(),
                                UnitTypeId.Millimeters);

                            // Базовое имя
                            string newName = $"{insTypeComment} толщиной {thicknessMm}";

                            // Особые линейки (K-Flex ST / Energocell HT / PE Compact)
                            if (insMark.IndexOf("k-flex st", StringComparison.OrdinalIgnoreCase) >= 0
                                || insMark.IndexOf("Energocell HT", StringComparison.OrdinalIgnoreCase) >= 0
                                || insTypeComment.IndexOf("k-flex st", StringComparison.OrdinalIgnoreCase) >= 0
                                || insTypeComment.IndexOf("Energocell HT", StringComparison.OrdinalIgnoreCase) >= 0)
                            {
                                newName = $"{insTypeComment} {thicknessMm}{CalculateRightInsulationName(hostOutsideDiamMm, hostFabric)} для трубопровода {TransformFabric(hostFabric)}";
                            }

                            if (insMark.IndexOf("PE Compact", StringComparison.OrdinalIgnoreCase) >= 0
                                || insTypeComment.IndexOf("PE Compact", StringComparison.OrdinalIgnoreCase) >= 0)
                            {
                                newName = $"{insTypeComment} {thicknessMm}{CalculateRightInsulationName(hostOutsideDiamMm, hostFabric)} из вспененного полиэтилена с наружным слоем из полимерной армирующей пленки для трубопровода {TransformFabric(hostFabric)}";
                            }

                            // === Расчёт количества по единицам измерения ===
                            double count = 0.0;

                            if (isCategotyPypeAcc)
                            {
                                // аксессуары считаем штуками
                                count = 1.0;
                            }
                            else if (insUnit == "м" || insUnit == "m")
                            {
                                // длина
                                double lenInt = insulation.get_Parameter(BuiltInParameter.CURVE_ELEM_LENGTH).AsDouble();
                                count = UnitUtils.ConvertFromInternalUnits(lenInt, UnitTypeId.Meters);
                            }
                            else if (insUnit == "м²" || insUnit == "m²")
                            {
                                // площадь: для фиттингов — площадь хоста, иначе — своя поверхность
                                if (host.Category.Name.IndexOf("Fittings", StringComparison.OrdinalIgnoreCase) >= 0
                                    || host.Category.Name.Contains("Cоединительные"))
                                {
                                    var pHostArea = host.get_Parameter(dictionaryGUID.ADSKSizeArea);
                                    if (pHostArea != null && pHostArea.StorageType == StorageType.Double)
                                        count = UnitUtils.ConvertFromInternalUnits(pHostArea.AsDouble(), UnitTypeId.SquareMeters);
                                }
                                else
                                {
                                    double areaInt = insulation.get_Parameter(BuiltInParameter.RBS_CURVE_SURFACE_AREA).AsDouble();
                                    count = UnitUtils.ConvertFromInternalUnits(areaInt, UnitTypeId.SquareMeters);
                                }
                            }
                            //else if (insUnit == "м³" || insUnit == "m³")
                            //{
                            //    // объём: используем общий параметр объёма изоляции, если он есть
                            //    var pVol = insulation.get_Parameter(dictionaryGUID.ADSKSizeVolume); // ← подставь ваш GUID объёма
                            //    if (pVol != null && pVol.StorageType == StorageType.Double)
                            //    {
                            //        count = UnitUtils.ConvertFromInternalUnits(pVol.AsDouble(), UnitTypeId.CubicMeters);
                            //    }
                            //    else
                            //    {
                            //        count = 0.0; // нет данных — безопасно не писать количество
                            //    }
                            //}
                            else if (insUnit != null && insUnit.Contains("шт"))
                            {
                                count = 1.0;
                            }
                            else
                            {
                                // запасной вариант — длина
                                double lenInt = insulation.get_Parameter(BuiltInParameter.CURVE_ELEM_LENGTH).AsDouble();
                                count = UnitUtils.ConvertFromInternalUnits(lenInt, UnitTypeId.Meters);
                            }

                            // Применяем коэффициент
                            double countWithK = count * koef;

                            // Запись имени
                            RevitUtils.SetParameterValue(insulation, dictionaryGUID.ADSKName, newName);

                            // Запись ADSK_Количество с учётом единиц самого параметра
                            var countParamIns = insulation.get_Parameter(dictionaryGUID.ADSKCount);
                            if (countParamIns != null && countParamIns.StorageType == StorageType.Double)
                            {
                                var u = countParamIns.GetUnitTypeId();
                                double toSet =
                                    (u == UnitTypeId.Meters) ? UnitUtils.ConvertToInternalUnits(countWithK, UnitTypeId.Meters) :
                                    (u == UnitTypeId.SquareMeters) ? UnitUtils.ConvertToInternalUnits(countWithK, UnitTypeId.SquareMeters) :
                                    (u == UnitTypeId.CubicMeters) ? UnitUtils.ConvertToInternalUnits(countWithK, UnitTypeId.CubicMeters) :
                                                                     countWithK;

                                RevitUtils.SetParameterValue(insulation, dictionaryGUID.ADSKCount, toSet);
                            }

                            // Толщина изоляции в ADSK_Толщина стенки (мм → футы)
                            RevitUtils.SetParameterValue(insulation, dictionaryGUID.ADSKThicknes, UnitUtils.ConvertToInternalUnits(thicknessMm, UnitTypeId.Millimeters));

                            // Прокидываем ADSK_Комплект с хоста, если есть у обоих
                            var hostSet = host.get_Parameter(dictionaryGUID.ADSKKomp);
                            if (hostSet != null && hostSet.StorageType == StorageType.String)
                            {
                                string kit = hostSet.AsString();
                                if (!string.IsNullOrEmpty(kit))
                                    RevitUtils.SetParameterValue(insulation, dictionaryGUID.ADSKKomp, kit);
                            }
                        }
                        catch (Exception ex)
                        {
                            testLog += $"Ошибка при обработке изоляции трубопроводов {ex}\n";
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
                logger.LogWarning($"Ошибки при заполнении параметров: {testLog}", docName);
                TaskDialog.Show("Готово", "Параметры заполнены, но возникли ошибки (см. лог)");
                return Result.Succeeded;
            }
            else
            {
                //TaskDialog.Show("Готово", "Параметры для спецификации заполнены!");
                logger.LogInfo("Параметры для спецификации заполнены успешно", docName);
                TaskDialog.Show("Успех", "Параметры для спецификации заполнены без ошибок");
                return Result.Succeeded;
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
            Category cat = elem?.Category;
            if (cat == null) return false;

            BuiltInCategory bic = (BuiltInCategory)cat.Id.IntegerValue;
            return _ductPipeFittings.Contains(bic);
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
            if (elem == null)
            {
                ductSize = string.Empty;
                return 0;
            }

            if (isFitting)
            {
                // RBS_CALCULATED_SIZE — строка; AsString предпочтительнее, но AsValueString тоже допустим.
                string sizeString =
                    elem.get_Parameter(BuiltInParameter.RBS_CALCULATED_SIZE)?.AsString()
                    ?? elem.get_Parameter(BuiltInParameter.RBS_CALCULATED_SIZE)?.AsValueString()
                    ?? string.Empty;

                return ParseFittingSize(sizeString, isRectangular, ref ductSize);
            }

            // Duct / MEPCurve: читаем в футах и переводим в мм
            if (isRectangular)
            {
                Parameter pW = elem.get_Parameter(BuiltInParameter.RBS_CURVE_WIDTH_PARAM);
                Parameter pH = elem.get_Parameter(BuiltInParameter.RBS_CURVE_HEIGHT_PARAM);

                double wFt = pW?.AsDouble() ?? 0.0;
                double hFt = pH?.AsDouble() ?? 0.0;

                int w = ToMmInt(wFt);
                int h = ToMmInt(hFt);

                if (w <= 0 && h <= 0)
                {
                    ductSize = string.Empty;
                    return 0;
                }

                int a = Math.Min(w, h);
                int b = Math.Max(w, h);
                ductSize = (a > 0 && b > 0) ? $"{a}x{b}" : (b > 0 ? $"{b}" : string.Empty);
                return b;
            }
            else
            {
                Parameter pD = elem.get_Parameter(BuiltInParameter.RBS_CURVE_DIAMETER_PARAM);
                double dFt = pD?.AsDouble() ?? 0.0;
                int d = ToMmInt(dFt);
                ductSize = d > 0 ? $"⌀{d}" : string.Empty;
                return d;
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

        private int ParseFittingSize(string sizeString, bool isRectangular, ref string ductSize)
        {
            ductSize = string.Empty;
            if (string.IsNullOrWhiteSpace(sizeString))
                return 0;

            // Нормализация: убираем пробелы, приводим x/×/х к 'x', выбрасываем единицы измерения
            string s = sizeString
                .Replace(" ", string.Empty)
                .Replace('×', 'x')
                .Replace('х', 'x')
                .Replace("мм", string.Empty)
                .Replace("mm", string.Empty);

            // 1) Попробуем явно выцепить первую пару AxB
            Match rectPair = Regex.Match(s, @"(\d+)\s*[xX]\s*(\d+)");
            // 2) Вытащим все числа из строки (полезно при цепочках вида "500x300-400x300" или "⌀200-⌀160")
            var allNums = Regex.Matches(s, @"\d+").Cast<Match>().Select(m => int.Parse(m.Value)).Where(n => n > 0).ToList();

            if (allNums.Count == 0)
                return 0;

            if (isRectangular)
            {
                if (rectPair.Success)
                {
                    int a = int.Parse(rectPair.Groups[1].Value);
                    int b = int.Parse(rectPair.Groups[2].Value);
                    int min = Math.Min(a, b);
                    int max = Math.Max(a, b);
                    ductSize = $"{min}x{max}";
                    return max;
                }
                // fallback: берём две наибольшие величины и формируем AxB
                allNums.Sort(); allNums.Reverse();
                int max1 = allNums[0];
                int max2 = allNums.Count > 1 ? allNums[1] : max1;
                int minSide = Math.Min(max1, max2);
                int maxSide = Math.Max(max1, max2);
                ductSize = $"{minSide}x{maxSide}";
                return maxSide;
            }
            else
            {
                // Круглые: ищем явные диаметры; если нет метки, просто берём максимум
                // Поддержим варианты "⌀200", "Ø200", "D200" (реже): если есть — всё равно allNums.Max() корректно сработает
                int dia = allNums.Max();
                ductSize = $"⌀{dia}";
                return dia;
            }
        }

        private const int RoundMm = 0;

        private static int ToMmInt(double feet)
        {
            double mm = UnitUtils.ConvertFromInternalUnits(feet, UnitTypeId.Millimeters);
            return (int)Math.Round(mm, RoundMm);
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
                case "Оцинкованная сталь": return "из оцинкованной стали";
                case "Неоцинкованная сталь": return "из неоцинкованной стали";
                case "Нержавеющая сталь": return "из нержавеющей стали";
                case "Сталь": return "из стали";
                case "Чугун": return "чугунного";
                case "Полиэтилен": return "из полиэтилена";
                case "Полипропилен": return "из полипропилена";
                case "НПВХ": return "из НПВХ";
                case "Медь": return "из меди";
                default:return "Не заполнен ADSK_Материал наименование у основы";
            }
        }

        /// <summary>
        /// Вспомогательный метод для подбора части наименования для трубопроводов марки PE Compact, K-Fles ST и Energoflex
        /// </summary>
        /// <param name="outSide"></param>
        /// <param name="hostFabric"></param>
        /// <returns></returns>
        private string CalculateRightInsulationName(double outsideDiamMm, string hostFabric)
        {
            double d = outsideDiamMm;
            string insulDiam = "Не удалось подобрать диаметр трубопровода под типоразмер изоляцию";

            if (d < 18) insulDiam = "18";
            else if (d < 22) insulDiam = "22";
            else if (d < 28) insulDiam = "28";
            else if (d < 35) insulDiam = "35";
            else if (d < 42) insulDiam = "42";
            else if (d < 48) insulDiam = "48";
            else if (d < 54) insulDiam = "54";
            else if (d < 60) insulDiam = "60";
            else if (d < 76) insulDiam = "76";
            else if (d < 89) insulDiam = "89";
            else if (d < 102) insulDiam = "102";
            else if (d < 108) insulDiam = "108";
            else if (d < 114) insulDiam = "114";
            else if (d < 125) insulDiam = "125";
            else if (d < 133) insulDiam = "133";
            else if (d < 140) insulDiam = "140";
            else if (d < 160) insulDiam = "160";

            return $"x{insulDiam}";
        }


        private static readonly HashSet<BuiltInCategory> _ductPipeFittings = new HashSet<BuiltInCategory>
        {
            BuiltInCategory.OST_DuctFitting,
            BuiltInCategory.OST_PipeFitting
        };
    }
}