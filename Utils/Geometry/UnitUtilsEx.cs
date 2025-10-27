namespace ATP_Common_Plugin.Utils.Geometry
{
    public static class UnitUtilsEx
    {
        private const double FtToM = 0.3048;

        public static double FeetToMeters(double feet)
        {
            return feet * FtToM;
        }

        public static double MetersToFeet(double meters) => meters / FtToM;

        public static double SquareFeetToSquareMeters(double ft2)
        {
            return ft2 * FtToM * FtToM;
        }
    }
}