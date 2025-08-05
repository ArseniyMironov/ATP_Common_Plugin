using ATP_Common_Plugin.Services;
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
    class MarkOpennings : IExternalCommand
    {
        private const double LABEL_WIDTH_MM = 30 / 304.8; // примерная ширина марки
        private const double LABEL_HEIGHT_MM = 25 / 304.8; // примерная высота марки
        private const int MAX_SPIRAL_STEPS_MM = 50;
        private const double STEP_MM = 15 / 304.8; // 150 мм шаг спирали

        public Result Execute(ExternalCommandData commandData, ref string message, ElementSet elements)
        {
            UIDocument uidoc = commandData.Application.ActiveUIDocument;
            Document doc = uidoc.Document;
            string docName = doc.Title;
            View view = doc.ActiveView;
            var logger = ATP_App.GetService<ILoggerService>();

            using (Transaction tr = new Transaction(doc, "Удаление старых марок"))
            {
                logger.LogInfo("Удаление старых марок отверстий", docName);
                tr.Start();
                var oldTags = new FilteredElementCollector(doc, view.Id)
                    .OfCategory(BuiltInCategory.OST_SprinklerTags)
                    .WhereElementIsNotElementType()
                    .Cast<IndependentTag>()
                    .ToList();

                foreach (var tag in oldTags)
                    doc.Delete(tag.Id);

                tr.Commit();
                logger.LogInfo("Cтарые мароки удалены", docName);
            }

            var sprinklers = new FilteredElementCollector(doc, view.Id)
                .OfCategory(BuiltInCategory.OST_Sprinklers)
                .WhereElementIsNotElementType()
                .Where(e => e is FamilyInstance fi && (fi.Symbol.Family.IsInPlace == false && fi.Host == null && fi.SuperComponent == null))
                .ToList();

            double viewScale = view.Scale;
            double LABEL_WIDTH = LABEL_WIDTH_MM * viewScale;
            double LABEWL_HEIGHT = LABEL_HEIGHT_MM * viewScale;
            double MAX_SPIRAL_STEPS = MAX_SPIRAL_STEPS_MM * viewScale;
            double STEP = STEP_MM * viewScale;

            int created = 0;

            using (Transaction tr = new Transaction(doc, "Маркировка"))
            {
                logger.LogInfo("Начало маркировки отверстий", docName);
                tr.Start();

                HashSet<Outline> occupiedZones = new HashSet<Outline>();
                foreach (var sprinkler in sprinklers)
                {
                    // Добавим габариты спринклера в список занятых зон
                    BoundingBoxXYZ bbox = sprinkler.get_BoundingBox(view);
                    if (bbox != null)
                    {
                        Outline sprinklerOutline = new Outline(bbox.Min, bbox.Max);
                        occupiedZones.Add(sprinklerOutline);
                        //DrawOutline(doc, sprinklerOutline, view); // (опционально) отрисуем спринклер
                    }

                    LocationPoint location = sprinkler.Location as LocationPoint;
                    if (location == null) continue;

                    XYZ origin = location.Point;

                    for (int i = 0; i < MAX_SPIRAL_STEPS; i++)
                    {
                        XYZ offset = SpiralOffset(i, STEP);
                        XYZ tagPos = origin + offset;

                        Outline tagOutline = GetLabelOutline(tagPos, LABEWL_HEIGHT, LABEL_WIDTH);
                        //DrawOutline(doc, tagOutline, view); // Отрисовка занятой зоны

                        if (!IntersectsWithOccupied(occupiedZones, tagOutline))
                        {
                            IndependentTag tag = IndependentTag.Create(doc,
                                view.Id,
                                new Reference(sprinkler),
                                false,
                                TagMode.TM_ADDBY_CATEGORY,
                                TagOrientation.Horizontal,
                                tagPos
                            );

                            tag.HasLeader = true;
                            tag.LeaderEndCondition = LeaderEndCondition.Free;

                            occupiedZones.Add(tagOutline);
                            created++;
                            break;
                        }
                    }
                }

                tr.Commit();
            }

            logger.LogInfo($"Создано {created} маркирок отверстий", docName);
            logger.LogInfo("Конец маркировки отверстий", docName);
            return Result.Succeeded;
        }

        private XYZ SpiralOffset(int step, double STEP)
        {
            if (step == 0)
                return XYZ.Zero;

            int layer = (int)Math.Floor((Math.Sqrt(step) + 1) / 2);
            int legLength = layer * 2;
            int legStart = (2 * layer - 1) * (2 * layer - 1);
            int leg = (step - legStart) / legLength;
            int posInLeg = (step - (2 * layer - 1) * (2 * layer - 1)) % legLength;

            int dx = 0, dy = 0;
            switch (leg)
            {
                case 0: dx = posInLeg - layer; dy = -layer; break; // вниз
                case 1: dx = layer; dy = -layer + posInLeg; break; // вправо
                case 2: dx = layer - posInLeg; dy = layer; break;  // вверх
                case 3: dx = -layer; dy = layer - posInLeg; break; // влево
            }

            return new XYZ(dx * STEP, dy * STEP, 0);
        }

        private bool IntersectsWithOccupied(HashSet<Outline> occupied, Outline candidate)
        {
            foreach (var zone in occupied)
            {
                if (zone.Intersects(candidate, 0.001))
                    return true;
            }

            return false;
        }

        private Outline GetLabelOutline(XYZ center, double LABEL_HEIGHT, double LABEL_WIDTH)
        {
            XYZ min = new XYZ(center.X - LABEL_WIDTH / 2, center.Y - LABEL_HEIGHT * 0.3, center.Z - 0.1);
            XYZ max = new XYZ(center.X + LABEL_WIDTH / 2, center.Y + LABEL_HEIGHT * 0.6, center.Z + 0.1);
            return new Outline(min, max);
        }

        private void DrawOutline(Document doc, Outline outline, View view)
        {
            XYZ p1 = outline.MinimumPoint;
            XYZ p2 = new XYZ(outline.MaximumPoint.X, outline.MinimumPoint.Y, outline.MinimumPoint.Z);
            XYZ p3 = outline.MaximumPoint;
            XYZ p4 = new XYZ(outline.MinimumPoint.X, outline.MaximumPoint.Y, outline.MinimumPoint.Z);

            CreateDetailLine(doc, view, p1, p2);
            CreateDetailLine(doc, view, p2, p3);
            CreateDetailLine(doc, view, p3, p4);
            CreateDetailLine(doc, view, p4, p1);
        }

        private void CreateDetailLine(Document doc, View view, XYZ p1, XYZ p2)
        {
            // Получаем плоскость вида
            Plane viewPlane = Plane.CreateByNormalAndOrigin(view.ViewDirection, view.Origin);

            // Проецируем точки в плоскость вида
            XYZ projectedP1 = ProjectPointToPlane(p1, viewPlane);
            XYZ projectedP2 = ProjectPointToPlane(p2, viewPlane);

            Line line = Line.CreateBound(projectedP1, projectedP2);
            doc.Create.NewDetailCurve(view, line);
        }
        private XYZ ProjectPointToPlane(XYZ point, Plane plane)
        {
            XYZ origin = plane.Origin;
            XYZ normal = plane.Normal;
            XYZ vec = point - origin;
            double distance = vec.DotProduct(normal);
            return point - distance * normal;
        }
    }
}