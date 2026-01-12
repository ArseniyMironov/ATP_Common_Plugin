using Autodesk.Revit.DB;

namespace ATP_Common_Plugin.Services.Spaces
{
    /// <summary>
    /// Геометрическая эвристика "внутренняя стена" по описанному правилу:
    /// из центра элемента шаг на d в сторону, противоположную центру помещения,
    /// затем +90°, затем 180° от последнего (с шагом 2d). Любое попадание в ДРУГОЕ Space → внутренняя.
    /// </summary>
    public sealed class InteriorFilterService
    {
        private readonly ILoggerService _logger;

        public InteriorFilterService(ILoggerService logger) { _logger = logger; }

        public bool IsInteriorWall(Document mainDoc, SpatialElement currentSpace, Wall wallHost, Transform linkToHost, XYZ noutMain, SpacesSpatialIndex3D index, double stepMeters = 1.5)
        {
            if (mainDoc == null || currentSpace == null || wallHost == null || index == null) return false;

            // Центр стены в mainDoc
            Transform t = linkToHost ?? Transform.Identity;
            XYZ centerLocal = GetWallCenter(wallHost);
            if (centerLocal == null) return false;
            XYZ center = t.OfPoint(centerLocal);

            // Z строго внутри текущего Space (середина bbox Space), чтобы не ловить чужие уровни
            BoundingBoxXYZ sbb = currentSpace.get_BoundingBox(null);
            if (sbb == null) return false;
            double zMid = 0.5 * (sbb.Min.Z + sbb.Max.Z);

            // Движемся НАРУЖУ строго по nout (надёжнее, чем "от центра к центру Space")
            XYZ nxy = new XYZ(noutMain.X, noutMain.Y, 0.0);
            double ln = nxy.GetLength();
            if (ln < 1e-9)
            {
                // Fallback: если nout сломался, попробуем «от центра к центру Space»
                XYZ sc = new XYZ(0.5 * (sbb.Min.X + sbb.Max.X), 0.5 * (sbb.Min.Y + sbb.Max.Y), zMid);
                XYZ vToSpace = new XYZ(sc.X - center.X, sc.Y - center.Y, 0.0);
                double lv = vToSpace.GetLength();
                if (lv < 1e-9) return false;
                nxy = (-vToSpace) / lv; // «наружу» — от помещения
            }
            else
            {
                nxy = nxy / ln;
            }

            double d = Utils.Geometry.UnitUtilsEx.MetersToFeet(stepMeters);

            // Точка строго снаружи от границы стены по нормали помещения
            XYZ p1 = new XYZ(center.X + nxy.X * d, center.Y + nxy.Y * d, zMid);

            // Если «снаружи» сразу попали в ДРУГОЕ помещение — это внутренняя перегородка (межкомнатная)
            var other = index.FindSpaceContainingPoint(mainDoc, p1, currentSpace);
            if (other != null) return true;

            // Иначе считаем наружной.
            return false;
        }

        private static XYZ GetWallCenter(Wall w)
        {
            var lc = w.Location as LocationCurve;
            if (lc != null && lc.Curve != null)
            {
                return lc.Curve.Evaluate(0.5, true);
            }
            var bb = w.get_BoundingBox(null);
            if (bb != null) return (bb.Min + bb.Max) * 0.5;
            return null;
        }
    }
}