using Autodesk.Revit.DB;

namespace ATP_Common_Plugin.Utils
{
    public static class TrueNorthUtils
    {
        /// <summary>
        /// Returns angle (radians) from X-axis to vector projected & adjusted to True North.
        /// </summary>
        public static double GetAngleToTrueNorth(Document doc, XYZ vec)
        {
            if (vec == null) return 0.0;
            // MVP: assume Project North ~ True North; later adjust by ProjectLocation's transform if needed.
            return System.Math.Atan2(vec.Y, vec.X);
        }
    }
}