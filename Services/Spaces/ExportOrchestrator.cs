using ATP_Common_Plugin.Commands.Calculation.SpacesEnvelopeExport;
using ATP_Common_Plugin.Models.Spaces;
using ATP_Common_Plugin.Utils.Excel;
using ATP_Common_Plugin.Utils.Geometry;
using Autodesk.Revit.DB;
using System;
using System.Collections.Generic;
using Outline = Autodesk.Revit.DB.Outline;
using Transform = Autodesk.Revit.DB.Transform;

namespace ATP_Common_Plugin.Services.Spaces
{
    public sealed class ExportOrchestrator
    {
        private readonly SpaceCollectorService _collector;
        private readonly SpaceBoundaryService _boundary;
        private readonly BoundaryExternalityService _externality;
        private readonly BoundaryClipService _clip;
        private readonly FaceMeasureService _measure;
        private readonly OrientationService _orient;
        private readonly SpacesSpatialIndex3D _index;
        private readonly ExcelExportService _excel;
        private readonly ILoggerService _logger;
        private readonly OpeningsOnHostService _openings;
        private readonly InteriorFilterService _interior;
        private readonly LayerTraceService _layers;

        public ExportOrchestrator(
            SpaceCollectorService collector,
            SpaceBoundaryService boundary,
            BoundaryExternalityService externality,
            BoundaryClipService clip,
            FaceMeasureService measure,
            OrientationService orient,
            SpacesSpatialIndex3D index,
            ExcelExportService excel,
            ILoggerService logger,
            OpeningsOnHostService openings,
            InteriorFilterService interior,
            LayerTraceService layers)
        {
            _collector = collector;
            _boundary = boundary;
            _externality = externality;
            _clip = clip;
            _measure = measure;
            _orient = orient;
            _index = index;
            _excel = excel;
            _logger = logger;
            _openings = openings;
            _interior = interior;
            _layers = layers;
        }

        public IList<SpaceInfo> Run(Document doc, ExportSpacesEnvelopeOptions options)
        {
            var spaces = new List<SpaceInfo>();
            var rawSpaces = _collector.CollectSpaces(doc);
            _index.Build(rawSpaces as IEnumerable<SpatialElement> ?? new List<SpatialElement>());

            int spaceIdx = 0;
            foreach (var sp in _collector.CollectSpaces(doc))
            {
                spaceIdx++;
                bool diag = spaceIdx <= 3;              // ← логируем только первые 3

                int subfacesTotal = 0;                  // ← количество subfaces всего
                int exteriorPassed = 0;                 // ← сколько прошло isExt
                double exteriorAreaSumM2 = 0.0;         // ← суммарная площадь наружных клип-гране

                var info = new SpaceInfo
                {
                    Name = sp.Name,
                    Number = sp.Number,
                    Area_M2 = Utils.Geometry.UnitUtilsEx.SquareFeetToSquareMeters(sp.Area)
                };

                var results = _boundary.GetResults(doc, sp);

                // Грани объёма Space
                Solid spaceSolid = results.GetGeometry();
                if (spaceSolid != null && spaceSolid.Faces != null)
                {
                    foreach (Face spaceFace in spaceSolid.Faces)
                    {
                        // Клип-подграни хост-элементов, соприкасающиеся именно с ЭТОЙ гранью пространства
                        var subfaces = results.GetBoundaryFaceInfo(spaceFace);
                        if (subfaces == null || subfaces.Count == 0) continue;

                        if (diag) subfacesTotal += subfaces.Count;

                        foreach (SpatialElementBoundarySubface subface in subfaces)
                        {
                            // Грань хоста (уже клипнутая по текущему Space)
                            Face spaceClipFace = subface.GetSubface();
                            if (spaceClipFace == null) continue;

                            // тГрань хоста (может быть null для «свободных» границ)
                            Face hostFace = subface.GetBoundingElementFace();

                            // Хост-элемент (может отсутствовать для free boundary)
                            Element host = null;
                            Document hostDoc = doc;
                            Transform linkToHost = Transform.Identity;

                            // Пытаемся получить хост через SpatialBoundaryElement (надёжно и для линков)
                            var sbe = subface.SpatialBoundaryElement;
                            if (sbe != null)
                            {
                                // 1) Элемент в текущем документе
                                ElementId hostId = sbe.HostElementId;
                                if (hostId != ElementId.InvalidElementId)
                                {
                                    host = doc.GetElement(hostId);
                                }
                                else
                                {
                                    // 2) Элемент в связанном документе
                                    ElementId linkedElemId = sbe.LinkedElementId;   // id элемента внутри linked-doc
                                    ElementId linkInstId = sbe.LinkInstanceId;    // id RevitLinkInstance в главном doc

                                    if (linkedElemId != ElementId.InvalidElementId &&
                                        linkInstId != ElementId.InvalidElementId)
                                    {
                                        var linkInst = doc.GetElement(linkInstId) as RevitLinkInstance;
                                        if (linkInst != null)
                                        {
                                            var linkDoc = linkInst.GetLinkDocument();
                                            if (linkDoc != null)
                                            {
                                                hostDoc = linkDoc;
                                                linkToHost = linkInst.GetTotalTransform() ?? Transform.Identity;
                                                host = linkDoc.GetElement(linkedElemId);
                                            }
                                        }
                                    }
                                }
                            }

                            // Репрезентативная точка/нормаль хоста
                            BoundingBoxUV bb = spaceClipFace.GetBoundingBox();
                            UV mid = new UV(0.5 * (bb.Min.U + bb.Max.U), 0.5 * (bb.Min.V + bb.Max.V));
                            XYZ p = spaceClipFace.Evaluate(mid);
                            XYZ nCandidate = spaceClipFace.ComputeNormal(mid).Normalize();

                            // Ориентируем нормаль относительно текущего Space: nin — внутрь, nout — наружу
                            XYZ nin, nout;
                            Utils.Geometry.TransformUtils.SelectSpaceOrientedNormals(sp, p, nCandidate, out nin, out nout);

                            // Геометрический тест «наружности» по наружной нормали на клип-грани
                            bool isExt = _externality.IsExteriorFace(doc, sp, spaceClipFace, p, nout, _index);
                            if (!isExt) continue;

                            if (diag) exteriorPassed++;

                            // Клип-контуры (если сервис отдаёт петли — по subface; иначе можно работать по самой face)
                            var loops = _clip.GetClippedLoopsFromBoundaryFace(subface);

                            // Измерения A/B/Area по ВНУТРЕННЕЙ клип-грани
                            double a, b, area; bool isHoriz;
                            _measure.MeasureFace(spaceClipFace, loops, out a, out b, out area, out isHoriz);

                            if (a <= 0 || b <= 0 || area <= 0) continue;

                            if (diag) exteriorAreaSumM2 += area;

                            // Ориентация по True North (для горизонталей — NA)
                            var ori = _orient.GetOrientation(doc, nout, isHoriz);

                            string typeName = string.Empty;
                            if (host != null)
                            {
                                var et = hostDoc.GetElement(host.GetTypeId());
                                if (et != null)
                                    typeName = et.Name;
                                else
                                    typeName = host.Name; // фоллбек, если тип недоступен
                            }

                            // Фильтр внутренних стен (только для вертикалей и только для стен)
                            bool isInterior = false;
                            if (!isHoriz && host is Wall wallHost1)
                            {
                                // linkToHost/hostDoc уже получены выше для этого subface
                                // nout — наружу от помещения в координатах mainDoc
                                try
                                {
                                    isInterior = _interior.IsInteriorWall(
                                        doc, sp, wallHost1, linkToHost, nout, _index, 1.5 /*м*/); ;
                                }
                                catch (Exception ex)
                                {
                                    _logger?.LogError($"Interior check failed for Wall {wallHost1.Id}: {ex.Message}");
                                    isInterior = false; // fail-open, лучше включить чем пропустить фасад
                                }
                            }
                            if (isInterior)
                            {
                                // внутренняя перегородка — пропускаем
                                continue;
                            }

                            if (host is Wall wallHost)
                            {
                                XYZ wallOutNormalHostDoc = nout; // nout уже "наружу" от Space
                                var openings = new List<BoundaryInfo>();

                                try
                                {
                                    // 1) Окна/двери на самой стене
                                    var collected = _openings.Collect(hostDoc, linkToHost, sp, wallHost, wallOutNormalHostDoc);
                                    if (collected != null && collected.Count > 0)
                                        openings.AddRange(collected);
                                }
                                catch (System.Exception ex)
                                {
                                    _logger?.LogError($"Openings.Collect(host) failed for wall {wallHost.Id}: {ex}");
                                }

                                try
                                {
                                    // 2) Если ничего не нашли — пробуем копланарную "за" отделкой
                                    if (openings.Count == 0)
                                    {
                                        double maxOffsetFt = UnitUtilsEx.MetersToFeet(1.0); // до 1 м "за отделкой"
                                        var behindWalls = _openings.FindCoplanarWallsBehind(hostDoc, wallHost, wallOutNormalHostDoc, maxOffsetFt);

                                        if (behindWalls != null)
                                        {
                                            foreach (Wall w2 in behindWalls)
                                            {
                                                try
                                                {
                                                    var openings2 = _openings.Collect(hostDoc, linkToHost, sp, w2, wallOutNormalHostDoc);
                                                    if (openings2 != null && openings2.Count > 0)
                                                        openings.AddRange(openings2);
                                                }
                                                catch (System.Exception ex2)
                                                {
                                                    _logger?.LogError($"Openings.Collect(behind {w2.Id}) failed: {ex2}");
                                                }
                                            }
                                        }
                                    }
                                }
                                catch (System.Exception ex)
                                {
                                    _logger?.LogError($"FindCoplanarWallsBehind failed for wall {wallHost.Id}: {ex}");
                                }

                                // если собрали что-то — добавим
                                if (openings.Count > 0)
                                {
                                    foreach (var openinig in openings)
                                        info.Boundaries.Add(openinig);
                                }
                            }

                            bool wroteAnyLayer = false;
                            if (!isHoriz && host is Wall startWall)
                            {
                                try
                                {
                                    // 1) Собираем цепочку параллельных стен «за отделкой»
                                    var wallsChain = new List<Wall>();
                                    wallsChain.Add(startWall);

                                    double maxOffsetFt = UnitUtilsEx.MetersToFeet(1.2); // до ~1.2 м наружу
                                    var nxy = new XYZ(nout.X, nout.Y, 0.0);
                                    double ln = nxy.GetLength(); if (ln > 1e-9) nxy = nxy / ln; else nxy = XYZ.BasisX;

                                    var behind = FindCoplanarWallsBehind(hostDoc, startWall, nxy, maxOffsetFt);
                                    if (behind != null && behind.Count > 0)
                                    {
                                        // Отсортируем по расстоянию вдоль nxy (ближе → дальше)
                                        XYZ hostMid = GetWallMidPoint(startWall);
                                        var behindList = behind as List<Wall> ?? new List<Wall>(behind);
                                        behindList.Sort((w1, w2) =>
                                        {
                                            double d1 = (GetWallMidPoint(w1) - hostMid).DotProduct(nxy);
                                            double d2 = (GetWallMidPoint(w2) - hostMid).DotProduct(nxy);
                                            return d1.CompareTo(d2);
                                        });
                                        wallsChain.AddRange(behindList);
                                    }

                                    // 2) Для каждой стены в цепочке берём ВНЕШНЮЮ плоскую грань по nout
                                    foreach (var w in wallsChain)
                                    {
                                        var outer = TryGetOuterPlanarFace(w, nxy); // nxy в координатах hostDoc
                                        if (outer == null) continue;

                                        double la, lb, larea; bool lHoriz;
                                        _measure.MeasureFace(outer, null, out la, out lb, out larea, out lHoriz);
                                        if (la <= 0 || lb <= 0 || larea <= 0) continue;

                                        var lori = _orient.GetOrientation(doc, nout, false);

                                        string ltypeName = string.Empty;
                                        var etL = hostDoc.GetElement(w.GetTypeId()) as ElementType;
                                        if (etL != null) ltypeName = etL.Name;

                                        info.Boundaries.Add(new BoundaryInfo
                                        {
                                            HostElementId = w.Id,
                                            Category = (w.Category != null) ? w.Category.Name : string.Empty,
                                            Family = w.Name,
                                            Type = w.GetType().Name,
                                            RevitTypeName = ltypeName,
                                            A_Height_M = System.Math.Round(la, Models.Settings.DigitsLen),
                                            B_Width_M = System.Math.Round(lb, Models.Settings.DigitsLen),
                                            Area_M2 = System.Math.Round(larea, Models.Settings.DigitsArea),
                                            Orientation = lori
                                        });

                                        wroteAnyLayer = true;
                                    }
                                }
                                catch (Exception ex)
                                {
                                    _logger?.LogError($"Wall-chain layering failed: {ex.Message}");
                                    wroteAnyLayer = false;
                                }
                            }

                            // Если слоёв не нашли — пишем базовую клип- грань, как раньше (fallback)
                            if (!wroteAnyLayer)
                            {
                                info.Boundaries.Add(new BoundaryInfo
                                {
                                    HostElementId = host != null ? host.Id : ElementId.InvalidElementId,
                                    Category = (host != null && host.Category != null) ? host.Category.Name : string.Empty,
                                    Family = host != null ? host.Name : string.Empty,
                                    Type = host != null ? host.GetType().Name : string.Empty,
                                    RevitTypeName = typeName,
                                    A_Height_M = System.Math.Round(a, Models.Settings.DigitsLen),
                                    B_Width_M = System.Math.Round(b, Models.Settings.DigitsLen),
                                    Area_M2 = System.Math.Round(area, Models.Settings.DigitsArea),
                                    Orientation = ori
                                });
                            }

                        }
                    }
                }

                // ← финальный лог по первому тройку помещений
                if (diag && _logger != null)
                {
                    _logger.LogInfo(
                        $"[Diag] Space \"{info.Name}\" ({info.Number}): " +
                        $"subfaces={subfacesTotal}, exterior={exteriorPassed}, " +
                        $"exteriorAreaSum={exteriorAreaSumM2:0.###} m², " +
                        $"boundariesWritten={info.Boundaries.Count}");
                }

                spaces.Add(info);
            }

            _excel.Export(spaces);

            // общий итог
            _logger?.LogInfo($"[Diag] Spaces processed: {spaces.Count}");

            return spaces;
        }

        // Найти копланарные/параллельные стены позади hostWall по заданной нормали (вперёд от помещения)
        // maxOffsetFt — максимальный зазор (в футах), например 1.0 м = 3.28084 ft
        public IList<Wall> FindCoplanarWallsBehind(
            Document hostDoc,
            Wall hostWall,
            XYZ noutHost,
            double maxOffsetFt)
        {
            var result = new List<Wall>();
            if (hostDoc == null || hostWall == null || noutHost == null || noutHost.GetLength() < 1e-9) return result;

            // Направление наружу по XY
            double ln = noutHost.GetLength();
            XYZ nxy = (ln > 1e-9) ? (new XYZ(noutHost.X, noutHost.Y, 0.0) / ln) : new XYZ(1, 0, 0);

            if (nxy.IsZeroLength()) return result;
            nxy = nxy.Normalize();

            // ББ стены + вытянем в сторону nxy
            BoundingBoxXYZ bb = hostWall.get_BoundingBox(null);
            if (bb == null) return result;
            Outline searchOl = new Outline(
                bb.Min - nxy.Multiply(maxOffsetFt),
                bb.Max + nxy.Multiply(maxOffsetFt)
            );

            var collector = new FilteredElementCollector(hostDoc)
                .OfClass(typeof(Wall))
                .WhereElementIsNotElementType()
                .ToElements();

            // Геометрия «похожести»:
            // 1) параллельность (вектора ориентации почти совпадают),
            // 2) положительная проекция смещения на nxy (стена действительно "за" отделкой),
            // 3) расстояние вдоль nxy не больше maxOffsetFt.
            XYZ hostOri = ((Wall)hostWall).Orientation;
            double lho = hostOri.GetLength();
            hostOri = (lho > 1e-9) ? hostOri / lho : new XYZ(0, 0, 0);

            // Возьмём реперную точку hostWall
            XYZ hostMid = GetWallMidPoint(hostWall);

            foreach (Element e in collector)
            {
                var w = e as Wall;
                if (w == null || w.Id == hostWall.Id) continue;

                // Быстрый фильтр по bbox
                var wbb = w.get_BoundingBox(null);
                if (wbb == null) continue;
                if (!IntersectsOutline(searchOl, wbb)) continue;

                // Параллельность
                XYZ wOri = w.Orientation;
                double lwo = wOri.GetLength();
                if (lwo < 1e-9) continue;
                wOri = wOri / lwo;
                double dotParallel = System.Math.Abs(hostOri.DotProduct(wOri));
                if (dotParallel < 0.98) continue; // не параллельные стены

                // Смещение вдоль nxy
                XYZ wMid = GetWallMidPoint(w);
                double eps = UnitUtilsEx.MetersToFeet(0.001); // ~1 мм допуск
                double d = (wMid - hostMid).DotProduct(nxy);
                if (d < -eps) continue;                       // реально «внутрь» помещения
                if (d > maxOffsetFt + eps) continue;         // слишком далеко

                result.Add(w);
            }

            return result;
        }

        private static XYZ GetWallMidPoint(Wall w)
        {
            XYZ mid;
            var loc = w.Location as LocationCurve;
            if (loc != null)
            {
                Curve c = loc.Curve;
                double p = 0.5;
                mid = c.Evaluate(p, true);
            }
            else
            {
                var bb = w.get_BoundingBox(null);
                mid = (bb.Min + bb.Max) * 0.5;
            }
            return new XYZ(mid.X, mid.Y, (w.get_BoundingBox(null).Min.Z + w.get_BoundingBox(null).Max.Z) * 0.5);
        }

        private static bool IntersectsOutline(Outline ol, BoundingBoxXYZ bb)
        {
            if (bb == null) return false;
            return !(bb.Max.X < ol.MinimumPoint.X || bb.Min.X > ol.MaximumPoint.X
                  || bb.Max.Y < ol.MinimumPoint.Y || bb.Min.Y > ol.MaximumPoint.Y
                  || bb.Max.Z < ol.MinimumPoint.Z || bb.Min.Z > ol.MaximumPoint.Z);
        }

        private static PlanarFace TryGetOuterPlanarFace(Wall w, XYZ nxyHost)
        {
            if (w == null) return null;
            double ln = nxyHost.GetLength();
            if (ln < 1e-9) return null;
            nxyHost = nxyHost / ln;

            var opt = new Options
            {
                ComputeReferences = false,
                IncludeNonVisibleObjects = true,
                DetailLevel = ViewDetailLevel.Fine
            };
            var ge = w.get_Geometry(opt);
            if (ge == null) return null;

            PlanarFace bestPos = null;
            double bestPosDot = 0.6; // мягче порог для «правильного» направления

            PlanarFace bestAbs = null;
            double bestAbsDot = 0.0;

            foreach (var go in ge)
            {
                var solid = go as Solid;
                if (solid == null || solid.Faces == null) continue;

                foreach (Face f in solid.Faces)
                {
                    var pf = f as PlanarFace;
                    if (pf == null) continue;

                    XYZ fn = pf.FaceNormal;
                    XYZ fnxy = new XYZ(fn.X, fn.Y, 0.0);
                    double lf = fnxy.GetLength();
                    if (lf < 1e-9) continue;
                    fnxy = fnxy / lf;

                    double dot = fnxy.DotProduct(nxyHost);
                    double ad = System.Math.Abs(dot);

                    // Кандидат в «правильном» направлении
                    if (dot > bestPosDot)
                    {
                        bestPosDot = dot;
                        bestPos = pf;
                    }

                    // Fallback: максимальная параллельность по модулю
                    if (ad > bestAbsDot)
                    {
                        bestAbsDot = ad;
                        bestAbs = pf;
                    }
                }
            }

            return bestPos ?? bestAbs;
        }
    }
}