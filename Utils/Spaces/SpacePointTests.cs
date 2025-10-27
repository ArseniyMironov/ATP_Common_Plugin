using Autodesk.Revit.DB;

namespace ATP_Common_Plugin.Utils.Spaces
{
    public static class SpacePointTests
    {
        public static bool IsPointInSpaceFast(SpatialElement space, XYZ p)
        {
            if (space == null || p == null) return false;
            // MVP: Revit API has limited direct point-in-space; use BoundingBox filter + tolerance if needed later.
            BoundingBoxXYZ bb = space.get_BoundingBox(null);
            if (bb == null) return false;
            return p.X >= bb.Min.X && p.X <= bb.Max.X
                && p.Y >= bb.Min.Y && p.Y <= bb.Max.Y
                && p.Z >= bb.Min.Z && p.Z <= bb.Max.Z;
        }
    }
}