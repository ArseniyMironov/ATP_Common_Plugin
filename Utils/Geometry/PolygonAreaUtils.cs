using Autodesk.Revit.DB;
using System.Collections.Generic;

namespace ATP_Common_Plugin.Utils.Geometry
{
    public static class PolygonAreaUtils
    {
        /// <summary>
        /// Shoelace formula in 2D (local h-v plane). Expects closed polygon (first point != last).
        /// </summary>
        public static double Area2D(IList<XYZ> points2D)
        {
            if (points2D == null || points2D.Count < 3) return 0.0;
            double s = 0.0;
            for (int i = 0; i < points2D.Count; i++)
            {
                var a = points2D[i];
                var b = points2D[(i + 1) % points2D.Count];
                s += (a.X * b.Y - b.X * a.Y);
            }
            return System.Math.Abs(s) * 0.5;
        }
    }
}