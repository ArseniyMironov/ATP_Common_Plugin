using ATP_Common_Plugin.Models.Spaces;
using ATP_Common_Plugin.Services;
using ATP_Common_Plugin.Services.Spaces;
using ATP_Common_Plugin.Utils.Excel;
using Autodesk.Revit.Attributes;
using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using System;
using System.Collections.Generic;

namespace ATP_Common_Plugin.Commands.Calculation.SpacesEnvelopeExport
{
    [Regeneration(RegenerationOption.Manual)]
    [Transaction(TransactionMode.Manual)]
    /// <summary>
    /// Entry point command (Ribbon button). Orchestrates collection, geometry, and export.
    /// </summary>
    public sealed class ExportSpacesEnvelopeCommand : IExternalCommand
    {
        public Result Execute(ExternalCommandData commandData, ref string message, ElementSet elements)
        {
            UIApplication uiapp = commandData.Application;
            UIDocument uidoc = uiapp.ActiveUIDocument;
            Document doc = uidoc.Document;
            var logger = ATP_App.GetService<ILoggerService>();

            if (uidoc == null || uidoc.Document == null)
            {
                logger.LogError("Нет открытого документа");
                return Result.Cancelled;
            }

            try
            {
                var options = new ExportSpacesEnvelopeOptions(); // TODO: fill from UI later

                var orchestrator = new ExportOrchestrator(
                    new SpaceCollectorService(logger),
                    new SpaceBoundaryService(logger),
                    new BoundaryExternalityService(logger),
                    new BoundaryClipService(logger),
                    new FaceMeasureService(logger),
                    new OrientationService(logger),
                    new SpacesSpatialIndex3D(logger),
                    new ExcelExportService(logger),
                    logger,
                    new OpeningsOnHostService(logger),
                    new InteriorFilterService(logger),
                    new LayerTraceService(logger));

                IList<SpaceInfo> result = orchestrator.Run(doc, options);
                int totalBoundaries = 0;
                foreach (var s in result)
                    if (s.Boundaries != null) totalBoundaries += s.Boundaries.Count;

                logger.LogInfo($"[Diag] Export done. Spaces={result.Count}, Boundaries total={totalBoundaries}");
                logger.LogInfo("Данные пространств успешно экспортированы");
                return Result.Succeeded;
            }
            catch (OperationCanceledException)
            {
                return Result.Cancelled;
            }
            catch (Exception ex)
            {
                logger.LogError(ex.Message);
                return Result.Failed;
            }
        }
    }
}