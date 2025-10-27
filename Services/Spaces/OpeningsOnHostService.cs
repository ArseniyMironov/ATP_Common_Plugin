using ATP_Common_Plugin.Models.Spaces;
using ATP_Common_Plugin.Utils;
using ATP_Common_Plugin.Utils.Geometry;
using Autodesk.Revit.DB;
using System;
using System.Collections.Generic;

namespace ATP_Common_Plugin.Services.Spaces
{
    /// <summary>
    /// Collects Windows/Doors hosted on a given wall-host that belong to the current Space.
    /// A/B are estimated from instance bbox (height = Z extent, width = extent along wall horizontal).
    /// Orientation is taken from the parent wall's outward normal (nout).
    /// </summary>
    public sealed class OpeningsOnHostService
    {
        private readonly ILoggerService _logger;

        public OpeningsOnHostService(ILoggerService logger) { _logger = logger; }

        public IList<BoundaryInfo> Collect(
                                           Document hostDoc,          // doc, где живёт стена/окна/двери (может быть linkDoc)
                                           Transform linkToHost,      // трансформ из hostDoc → главный doc (Identity для не-линка)
                                           SpatialElement space,      // space из главного doc
                                           Element host,              // стена в hostDoc
                                           XYZ wallOutNormalHostDoc)  // нормаль «наружу» в координатах главного doc (см. ниже)
        {
            var result = new List<BoundaryInfo>();
            if (wallOutNormalHostDoc == null) return result;
            if (hostDoc == null || space == null || host == null) return result;
            if (!(host is Wall)) return result;

            // Горизонталь вдоль стены (в координатах ГЛАВНОГО документа)
            XYZ noutHost = new XYZ(wallOutNormalHostDoc.X, wallOutNormalHostDoc.Y, 0.0);
            double len = noutHost.GetLength();
            if (len < 1e-9) return result;
            noutHost = noutHost / len;
            XYZ hHost = noutHost.CrossProduct(XYZ.BasisZ);
            if (hHost.GetLength() < 1e-9) hHost = XYZ.BasisX;
            else hHost = hHost.Normalize();

            // Собираем окна/двери в документе хоста (hostDoc)
            var fams = new FilteredElementCollector(hostDoc)
                .WhereElementIsNotElementType()
                .OfClass(typeof(FamilyInstance))
                .ToElements();

            foreach (Element e in fams)
            {
                var fi = e as FamilyInstance;
                if (fi == null) continue;
                var cat = fi.Category;
                if (cat == null) continue;
                var bic = (BuiltInCategory)cat.Id.IntegerValue;
                if (bic != BuiltInCategory.OST_Windows && bic != BuiltInCategory.OST_Doors) continue;

                if (fi.Host == null || fi.Host.Id != host.Id) continue;

                // Центр экземпляра (в координатах hostDoc)
                XYZ p0_local = GetInstanceCenter(fi);
                if (p0_local == null) continue;

                // Трансформируем в координаты главного doc, чтобы корректно проверять принадлежность Space
                XYZ p0_host = linkToHost.OfPoint(p0_local);

                // Маленький шаг внутрь помещения: двигаемся ПРОТИВ наружной нормали
                double eps = UnitUtilsEx.MetersToFeet(0.15);
                XYZ pInside = p0_host - noutHost.Multiply(eps);
                if (!Utils.Spaces.SpacePointTests.IsPointInSpaceFast(space, pInside))
                    continue; // окно/дверь не на нашем Space

                // Размеры по bbox экземпляра (в локальном doc), но проекция на ось h считаем уже в host-координатах
                BoundingBoxXYZ bb = fi.get_BoundingBox(null);
                if (bb == null) continue;

                double a_m = UnitUtilsEx.FeetToMeters(System.Math.Max(0.0, bb.Max.Z - bb.Min.Z)); // высота

                // 8 углов bbox → в host-координаты → проекция на hHost
                var corners_local = GetBBoxCorners(bb);
                double minH = double.PositiveInfinity, maxH = double.NegativeInfinity;
                foreach (var cLocal in corners_local)
                {
                    XYZ cHost = linkToHost.OfPoint(cLocal);
                    double proj = cHost.DotProduct(hHost);
                    if (proj < minH) minH = proj;
                    if (proj > maxH) maxH = proj;
                }
                double b_ft = System.Math.Max(0.0, maxH - minH);
                double b_m = UnitUtilsEx.FeetToMeters(b_ft);

                if (a_m < Models.Settings.MinExtent || b_m < Models.Settings.MinExtent) continue;

                double area_m2 = System.Math.Round(a_m * b_m, Models.Settings.DigitsArea);

                // Ориентация берём по noutHost
                var orient = Orientation4FromNormal(space.Document, noutHost);

                // Метки
                string typeName = string.Empty;
                var et = hostDoc.GetElement(fi.GetTypeId()) as ElementType;
                if (et != null) typeName = et.Name;

                result.Add(new BoundaryInfo
                {
                    HostElementId = fi.Id,
                    Category = cat.Name,
                    Family = fi.Name,
                    Type = fi.GetType().Name,
                    RevitTypeName = typeName,
                    A_Height_M = System.Math.Round(a_m, Models.Settings.DigitsLen),
                    B_Width_M = System.Math.Round(b_m, Models.Settings.DigitsLen),
                    Area_M2 = area_m2,
                    Orientation = orient
                });
            }

            return result;
        }

        private static XYZ GetInstanceCenter(FamilyInstance fi)
        {
            var lp = fi.Location as LocationPoint;
            if (lp != null) return lp.Point;

            BoundingBoxXYZ bb = fi.get_BoundingBox(null);
            if (bb == null) return null;
            return (bb.Min + bb.Max) * 0.5;
        }

        private static IList<XYZ> GetBBoxCorners(BoundingBoxXYZ bb)
        {
            var list = new List<XYZ>(8);
            var min = bb.Min; var max = bb.Max;
            list.Add(new XYZ(min.X, min.Y, min.Z));
            list.Add(new XYZ(max.X, min.Y, min.Z));
            list.Add(new XYZ(min.X, max.Y, min.Z));
            list.Add(new XYZ(max.X, max.Y, min.Z));
            list.Add(new XYZ(min.X, min.Y, max.Z));
            list.Add(new XYZ(max.X, min.Y, max.Z));
            list.Add(new XYZ(min.X, max.Y, max.Z));
            list.Add(new XYZ(max.X, max.Y, max.Z));
            return list;
        }

        private static Orientation4 Orientation4FromNormal(Autodesk.Revit.DB.Document doc, XYZ outward)
        {
            if (System.Math.Abs(outward.Z) > 0.9) return Orientation4.NA;
            XYZ nxy = new XYZ(outward.X, outward.Y, 0.0);
            if (nxy.GetLength() < 1e-9) return Orientation4.NA;
            double ang = TrueNorthUtils.GetAngleToTrueNorth(doc, nxy);
            return TransformUtils.QuantizeTo4(ang);
        }

        internal IEnumerable<object> FindCoplanarWallsBehind(Document hostDoc, Wall hostWall, XYZ wallOutNormalHostDoc, double maxOffsetFt)
        {
            var result = new List<Wall>();
            if (hostDoc == null || hostWall == null) return result;

            // Направление наружу по XY
            XYZ nxy = new XYZ(wallOutNormalHostDoc.X, wallOutNormalHostDoc.Y, 0.0);
            double lenDir = nxy.GetLength();
            if (lenDir < 1e-9) return result;
            nxy = nxy / lenDir;

            // ББ hostWall и поисковый объём (расширим вдоль nxy)
            BoundingBoxXYZ bb = hostWall.get_BoundingBox(null);
            if (bb == null) return result;
            Outline searchOl = new Outline(
                bb.Min - nxy.Multiply(maxOffsetFt),
                bb.Max + nxy.Multiply(maxOffsetFt)
            );

            // Ориентация базовой стены (перпендикуляр к стене)
            XYZ hostOri = hostWall.Orientation;
            double lenHO = hostOri.GetLength();
            if (lenHO >= 1e-9) hostOri = hostOri / lenHO; else hostOri = new XYZ(0, 0, 0);

            // Реперная точка базовой стены
            XYZ hostMid = GetWallMidPoint(hostWall);

            var collector = new FilteredElementCollector(hostDoc)
                .OfClass(typeof(Wall))
                .WhereElementIsNotElementType()
                .ToElements();

            foreach (Element e in collector)
            {
                var w = e as Wall;
                if (w == null || w.Id == hostWall.Id) continue;

                // Быстрый фильтр по bbox
                var wbb = w.get_BoundingBox(null);
                if (wbb == null) continue;
                if (!IntersectsOutline(searchOl, wbb)) continue;

                // Параллельность (почти та же ориентация)
                XYZ wOri = w.Orientation;
                double lenWO = wOri.GetLength();
                if (lenWO < 1e-9) continue;
                wOri = wOri / lenWO;
                double dotParallel = System.Math.Abs(hostOri.DotProduct(wOri));
                if (dotParallel < 0.98) continue;

                // Смещение вдоль nxy: стена действительно "за" отделкой и не дальше maxOffsetFt
                XYZ wMid = GetWallMidPoint(w);
                double d = (wMid - hostMid).DotProduct(nxy);
                if (d <= 0) continue;
                if (d > maxOffsetFt) continue;

                result.Add(w);
            }

            return result;
        }

        private static XYZ GetWallMidPoint(Wall w)
        {
            var loc = w.Location as LocationCurve;
            if (loc != null && loc.Curve != null)
            {
                return loc.Curve.Evaluate(0.5, true);
            }
            var bb = w.get_BoundingBox(null);
            if (bb != null) return (bb.Min + bb.Max) * 0.5;
            return XYZ.Zero;
        }

        private static bool IntersectsOutline(Outline ol, BoundingBoxXYZ bb)
        {
            if (bb == null) return false;
            return !(bb.Max.X < ol.MinimumPoint.X || bb.Min.X > ol.MaximumPoint.X
                  || bb.Max.Y < ol.MinimumPoint.Y || bb.Min.Y > ol.MaximumPoint.Y
                  || bb.Max.Z < ol.MinimumPoint.Z || bb.Min.Z > ol.MaximumPoint.Z);
        }
    }
}
