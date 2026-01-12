using Autodesk.Revit.DB;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ATP_Common_Plugin.Services.Spaces
{
    /// <summary>
    /// Находит "слои" вертикальных ограждений наружу от грани помещения:
    /// строит узкий призматический solid по нормали nout до MaxDepth и ищет элементы,
    /// чьи внешние плоские грани параллельны грани помещения и смотрят в сторону nout.
    /// Возвращает упорядоченный список от ближайшего к помещению слоя к дальнему.
    /// </summary>
    public sealed class LayerTraceService
    {
        private readonly ILoggerService _logger;
        public LayerTraceService(ILoggerService logger) { _logger = logger; }

        // Параметры по умолчанию (метры → футы)
        private const double DefaultMaxDepthM = 1.2;        // максимальная глубина поиска слоёв наружу
        private const double DefaultProbeWidthM = 0.25;      // ширина щупа (квадратного) в плоскости грани

        public IList<LayerHit> TraceVerticalLayers(
            Document hostDoc,                // документ, где будем искать элементы (main или link)
            Transform linkToHost,            // main → hostDoc (Identity для main)
            Face spaceClipFaceMain,          // клип-грань помещения (в координатах main)
            XYZ noutMain,                    // нормаль «наружу» (в координатах main)
            double? maxDepthM = null,
            double? probeWidthM = null)
        {
            var hits = new List<LayerHit>();
            if (hostDoc == null || spaceClipFaceMain == null || noutMain == null) return hits;

            // 1) Направления в плоскости грани (в main)
            BoundingBoxUV bb = spaceClipFaceMain.GetBoundingBox();
            UV mid = new UV(0.5 * (bb.Min.U + bb.Max.U), 0.5 * (bb.Min.V + bb.Max.V));
            XYZ n = spaceClipFaceMain.ComputeNormal(mid).Normalize();

            // Вертикальный базис в плоскости: v — проекция глобальной Z
            XYZ z = XYZ.BasisZ;
            XYZ v = z - (z.DotProduct(n)) * n;
            double vLen = v.GetLength();
            if (vLen < 1e-9)
            {
                // горизонтальная грань — слои не ищем
                return hits;
            }
            v = v / vLen;
            XYZ h = n.CrossProduct(v).Normalize();

            // 2) Призма-щуп (в main), затем преобразуем в hostDoc
            double maxDepthFt = Utils.Geometry.UnitUtilsEx.MetersToFeet(maxDepthM ?? DefaultMaxDepthM);
            double halfWft = Utils.Geometry.UnitUtilsEx.MetersToFeet((probeWidthM ?? DefaultProbeWidthM) * 0.5);

            XYZ p0Main = spaceClipFaceMain.Evaluate(mid);
            XYZ noutUnit = noutMain;
            double nLen = noutUnit.GetLength();
            if (nLen < 1e-9) return hits;
            noutUnit = noutUnit / nLen;

            // Квадратный профиль в плоскости (centered в p0Main)
            var profile = new List<CurveLoop>(1);
            var loop = new CurveLoop();

            XYZ pA = p0Main + h * (+halfWft) + v * (+halfWft);
            XYZ pB = p0Main + h * (-halfWft) + v * (+halfWft);
            XYZ pC = p0Main + h * (-halfWft) + v * (-halfWft);
            XYZ pD = p0Main + h * (+halfWft) + v * (-halfWft);

            loop.Append(Line.CreateBound(pA, pB));
            loop.Append(Line.CreateBound(pB, pC));
            loop.Append(Line.CreateBound(pC, pD));
            loop.Append(Line.CreateBound(pD, pA));
            profile.Add(loop);

            // Экструзия по nout
            Solid probeMain = GeometryCreationUtilities.CreateExtrusionGeometry(profile, noutUnit, maxDepthFt);

            // В координаты hostDoc
            Transform mainToHost = (linkToHost == null || linkToHost.IsIdentity) ? Transform.Identity : linkToHost.Inverse;
            Solid probeHost = SolidUtils.CreateTransformed(probeMain, mainToHost);

            // 3) Быстрый коллектор + точный фильтр пересечения solid’ом
            var intersectFilter = new ElementIntersectsSolidFilter(probeHost);

            var candidates = new FilteredElementCollector(hostDoc)
                .WhereElementIsNotElementType()
                .WherePasses(intersectFilter)
                .ToElements();

            if (candidates == null || candidates.Count == 0) return hits;

            // 4) Выбор нужной плоской грани для каждого элемента
            // Преобразуем nout в координаты hostDoc для сравнения нормалей
            XYZ noutHost = mainToHost.OfVector(noutUnit);

            var opt = new Options
            {
                ComputeReferences = true,
                IncludeNonVisibleObjects = true,
                DetailLevel = ViewDetailLevel.Fine
            };

            // Реперная точка для сортировки по расстоянию
            XYZ p0Host = mainToHost.OfPoint(p0Main);

            foreach (var e in candidates)
            {
                // Берём только вертикальные ограждения: стены и панели
                var cat = e.Category;
                if (cat == null) continue;
                BuiltInCategory bic = (BuiltInCategory)cat.Id.IntegerValue;
                if (bic != BuiltInCategory.OST_Walls && bic != BuiltInCategory.OST_CurtainWallPanels) continue;

                try
                {
                    GeometryElement ge = e.get_Geometry(opt);
                    if (ge == null) continue;

                    LayerHit best = null;
                    double bestD = 1.0;

                    foreach (GeometryObject go in ge)
                    {
                        var solid = go as Solid;
                        if (solid == null || solid.Faces == null) continue;

                        foreach (Face f in solid.Faces)
                        {
                            var pf = f as PlanarFace;
                            if (pf == null) continue;

                            // Параллельность и направление наружу
                            XYZ fn = pf.FaceNormal; // в координатах hostDoc
                            double dot = fn.DotProduct(noutHost);
                            if (dot < 0.95) continue; // не та сторона/не параллельно

                            // Расстояние от репера до плоскости вдоль noutHost — выбираем самую дальнюю "наружную"
                            double signed = fn.DotProduct(p0Host - pf.Origin);
                            double dist = System.Math.Abs(signed); // плоскости могут иметь разные origin; берём модуль
                            // Но важно идти по направлению nout: dot>0 уже гарантирован, используем dist как метрику дальности
                            if (dist > bestD)
                            {
                                bestD = dist;
                                best = new LayerHit
                                {
                                    Element = e,
                                    OuterFace = pf,
                                    HostDoc = hostDoc,
                                    LinkToHost = linkToHost
                                };
                            }
                        }
                    }

                    if (best != null)
                        hits.Add(best);
                }
                catch (Exception ex)
                {
                    _logger?.LogError($"LayerTrace: element {e.Id} failed: {ex.Message}");
                }
            }

            // 5) Сортируем по расстоянию от p0 вдоль nout (от ближнего к дальнему)
            hits.Sort((a, b) =>
            {
                var pfA = a.OuterFace as PlanarFace;
                var pfB = b.OuterFace as PlanarFace;
                if (pfA == null || pfB == null) return 0;

                double da = Math.Abs(noutHost.DotProduct(p0Host - pfA.Origin));
                double db = Math.Abs(noutHost.DotProduct(p0Host = pfB.Origin));
                return da.CompareTo(db);
            });

            return hits;

        }
    }

    public sealed class LayerHit
    {
        public Element Element;          // найденный элемент-слой
        public Face OuterFace;           // его внешняя грань (в hostDoc)
        public Document HostDoc;         // док, где живёт Element
        public Transform LinkToHost;     // main → hostDoc (Identity для main)
    }
}
