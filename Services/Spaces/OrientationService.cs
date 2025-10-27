using ATP_Common_Plugin.Models.Spaces;
using ATP_Common_Plugin.Utils;
using ATP_Common_Plugin.Utils.Geometry;
using Autodesk.Revit.DB;

namespace ATP_Common_Plugin.Services.Spaces
{
    public sealed class OrientationService
    {
        private readonly ILoggerService _logger;

        public OrientationService(ILoggerService logger) { _logger = logger; }
        public Orientation4 GetOrientation(Autodesk.Revit.DB.Document doc, XYZ outwardNormal, bool isHorizontal)
        {
            if (isHorizontal) return Orientation4.NA;

            // project to XY, adjust by True North
            XYZ nxy = new XYZ(outwardNormal.X, outwardNormal.Y, 0.0);
            if (nxy.IsZeroLength()) return Orientation4.NA;

            double ang = TrueNorthUtils.GetAngleToTrueNorth(doc, nxy); // radians 0..2π

            // Quantize to N/E/S/W (every 90°, +/-45°)
            return TransformUtils.QuantizeTo4(ang);
        }
    }
}