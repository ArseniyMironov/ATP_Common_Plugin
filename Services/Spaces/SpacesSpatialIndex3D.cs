using Autodesk.Revit.DB;
using System.Collections.Generic;

namespace ATP_Common_Plugin.Services.Spaces
{
    /// <summary>
    /// Simple spatial index over Spaces' bounding boxes to speed up "point in space" queries.
    /// </summary>
    public sealed class SpacesSpatialIndex3D
    {
        private readonly ILoggerService _logger;
        private readonly List<SpatialElement> _spaces = new List<SpatialElement>();

        public SpacesSpatialIndex3D(ILoggerService logger) { _logger = logger; }
        public void Build(IEnumerable<SpatialElement> spaces)
        {
            _spaces.Clear();
            foreach (var s in spaces) _spaces.Add(s);
        }

        public SpatialElement FindSpaceContainingPoint(Document doc, XYZ point, SpatialElement exclude)
        {
            foreach (var s in _spaces)
            {
                if (s == null || (exclude != null && s.Id == exclude.Id)) continue;
                if (Utils.Spaces.SpacePointTests.IsPointInSpaceFast(s, point)) return s;
            }
            return null;
        }
    }
}