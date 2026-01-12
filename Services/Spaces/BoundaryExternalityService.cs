using ATP_Common_Plugin.Utils.Geometry;
using ATP_Common_Plugin.Utils.Spaces;
using Autodesk.Revit.DB;

namespace ATP_Common_Plugin.Services.Spaces
{
    /// <summary>
    /// Geometric test: interior vs exterior using point-in-space checks.
    /// </summary>
    public sealed class BoundaryExternalityService
    {
        private readonly ILoggerService _logger;

        public BoundaryExternalityService(ILoggerService logger) { _logger = logger; }
        public bool IsExteriorFace(
            Document doc,
            SpatialElement currentSpace,
            Face face,
            XYZ pointOnFace,
            XYZ faceNormal,
            SpacesSpatialIndex3D index)
        {
            // Move a bit inside and outside along the normal
            double eps = UnitUtilsEx.MetersToFeet(Models.Settings.Epsilon);
            XYZ pin = pointOnFace - faceNormal.Multiply(eps);
            XYZ pout = pointOnFace + faceNormal.Multiply(eps);

            bool inCurrent = SpacePointTests.IsPointInSpaceFast(currentSpace, pin);
            if (!inCurrent) return false; // safety: should be inside current space

            // if outside point is NOT in any space => exterior
            var other = index.FindSpaceContainingPoint(doc, pout, currentSpace);
            return other == null;
        }
    }
}