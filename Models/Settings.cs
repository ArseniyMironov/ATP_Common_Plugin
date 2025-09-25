namespace ATP_Common_Plugin.Models
{
    public static class Settings
    {
        // geometry thresholds (meters)
        public const double Epsilon = 0.05;
        public const double MinExtent = 0.05;

        // rounding
        public const int DigitsLen = 3;
        public const int DigitsArea = 3;

        // excel
        public const string SheetName = "Spaces_Envelope";
        public const string HeaderRoom = "Space Name";
        public const string HeaderArea = "Area, m²";
        public const string HeaderBoundary = "Boundary";
        public const string HeaderA = "A (Height), m";
        public const string HeaderB = "B (Width), m";
        public const string HeaderS = "Area, m²";
        public const string HeaderOri = "Orientation";
    }
}