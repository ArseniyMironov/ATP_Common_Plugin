using Autodesk.Revit.DB;

namespace ATP_Common_Plugin.Services.Spaces
{
    /// <summary>
    /// Wraps SpatialElementGeometryCalculator usage.
    /// </summary>
    public sealed class SpaceBoundaryService
    {
        private readonly ILoggerService _logger;

        public SpaceBoundaryService(ILoggerService logger) { _logger = logger; }
        public SpatialElementGeometryResults GetResults(Document doc, SpatialElement space)
        {
            var opts = new SpatialElementBoundaryOptions
            {
                SpatialElementBoundaryLocation = SpatialElementBoundaryLocation.Finish,
                StoreFreeBoundaryFaces = true
            };
            var calc = new SpatialElementGeometryCalculator(doc, opts);
            return calc.CalculateSpatialElementGeometry(space);
        }
    }
}