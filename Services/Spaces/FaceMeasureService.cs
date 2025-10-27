using ATP_Common_Plugin.Utils.Geometry;
using Autodesk.Revit.DB;
using System;

namespace ATP_Common_Plugin.Services.Spaces
{
    /// <summary>
    /// Measures A/B extents and area for clipped face polygons.
    /// </summary>
    public sealed class FaceMeasureService
    {
        private readonly ILoggerService _logger;

        public FaceMeasureService(ILoggerService logger) { _logger = logger; }
        public void MeasureFace(
            Face face,
            CurveLoop[] clippedLoops,
            out double aMeters,
            out double bMeters,
            out double areaM2,
            out bool isHorizontal)
        {
            aMeters = 0.0;
            bMeters = 0.0;
            areaM2 = 0.0;
            isHorizontal = false;

            if (face == null) return;

            // 1) Нормаль и плоскость аппроксимации (для неплоских тоже берём локальную плоскость по нормали в центре)
            BoundingBoxUV bb = face.GetBoundingBox();
            UV mid = new UV(0.5 * (bb.Min.U + bb.Max.U), 0.5 * (bb.Min.V + bb.Max.V));
            XYZ n = face.ComputeNormal(mid).Normalize();

            // 2) Вертикаль в плоскости грани: проекция глобальной Z на плоскость
            XYZ z = XYZ.BasisZ;
            XYZ v = z - (z.DotProduct(n)) * n; // remove normal component
            double vLen = v.GetLength();

            if (vLen < 1e-6)
            {
                // Горизонтальная грань: ориентацию потом ставим NA, оси берём орт-норм в плоскости
                isHorizontal = true;
                // Для горизонтали просто возьмём два ортогональных базиса в плоскости:
                // подберём h и v из произвольных векторов
                XYZ any = Math.Abs(n.X) < 0.9 ? XYZ.BasisX : XYZ.BasisY;
                XYZ h = (any - any.DotProduct(n) * n).Normalize(); // в плоскости
                v = n.CrossProduct(h).Normalize();                 // второй базис в плоскости
                ComputeExtents(face, h, v, out aMeters, out bMeters);
            }
            else
            {
                // 3) Нормальный случай: v — "высота" (по вертикали в плоскости), h — поперёк
                v = v / vLen;
                XYZ h = n.CrossProduct(v).Normalize();

                ComputeExtents(face, h, v, out aMeters, out bMeters);
            }

            // 4) Площадь — берём площадь клип-грани (в футах) и переводим в м²
            areaM2 = UnitUtilsEx.SquareFeetToSquareMeters(face.Area);

            // 5) Отсечка микрограней
            if (aMeters < Models.Settings.MinExtent || bMeters < Models.Settings.MinExtent)
            {
                aMeters = 0.0; bMeters = 0.0; areaM2 = 0.0; // дадим понять вызывающему, что грань слишком мала
            }
        }

        // Вспомогательная: собираем точки по EdgeLoops и считаем экстенты в локальных осях (h — ширина, v — высота)
        private static void ComputeExtents(Face face, XYZ h, XYZ v, out double aMeters, out double bMeters)
        {
            double minH = double.PositiveInfinity, maxH = double.NegativeInfinity;
            double minV = double.PositiveInfinity, maxV = double.NegativeInfinity;

            // Тесселлируем рёбра клип-грани
            EdgeArrayArray loops = face.EdgeLoops;
            for (int i = 0; i < loops.Size; i++)
            {
                EdgeArray loop = loops.get_Item(i);
                for (int j = 0; j < loop.Size; j++)
                {
                    Edge e = loop.get_Item(j);
                    var pts = e.AsCurve().Tessellate();
                    foreach (XYZ p in pts)
                    {
                        // Проекция точки на базис (h,v) в плоскости грани
                        // Смещение от произвольного начала: возьмём любую точку на грани через Evaluate(mid)
                        // (для притирки можно кешировать origin)
                        // Здесь origin = проекция самой точки (в любом случае разности идут по dot)
                        double ph = p.DotProduct(h);
                        double pv = p.DotProduct(v);
                        if (ph < minH) minH = ph;
                        if (ph > maxH) maxH = ph;
                        if (pv < minV) minV = pv;
                        if (pv > maxV) maxV = pv;
                    }
                }
            }

            double extentH_ft = (maxH - minH);
            double extentV_ft = (maxV - minV);

            // Перевод фут→метр. DotProduct даёт значения в футах, как координаты.
            bMeters = UnitUtilsEx.FeetToMeters(System.Math.Max(0.0, extentH_ft));
            aMeters = UnitUtilsEx.FeetToMeters(System.Math.Max(0.0, extentV_ft));
        }
    }
}