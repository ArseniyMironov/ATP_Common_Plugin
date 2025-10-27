namespace ATP_Common_Plugin.Utils.Spaces
{
    public static class BoundaryFilters
    {
        public static bool IsTiny(double aMeters, double bMeters)
        {
            return aMeters < Models.Settings.MinExtent || bMeters < Models.Settings.MinExtent;
        }
    }
}