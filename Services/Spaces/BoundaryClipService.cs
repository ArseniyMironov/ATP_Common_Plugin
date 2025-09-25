using Autodesk.Revit.DB;

namespace ATP_Common_Plugin.Services.Spaces
{
    /// <summary>
    /// Ensures we measure only the clipped portion of the host face that actually touches the Space.
    /// MVP: rely on BoundaryFace loops from calculator; extend later if clipping is needed.
    /// </summary>
    public sealed class BoundaryClipService
    {
        private readonly ILoggerService _logger;

        public BoundaryClipService(ILoggerService logger) { _logger = logger; }
        public CurveLoop[] GetClippedLoopsFromBoundaryFace(SpatialElementBoundarySubface subface)
        {
            // TODO: extract/convert boundary face loops to planarity frame.
            // For skeleton: return empty array to compile; real impl will fill loops.
            return new CurveLoop[0];
        }
    }
}