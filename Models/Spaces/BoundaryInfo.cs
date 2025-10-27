using Autodesk.Revit.DB;

namespace ATP_Common_Plugin.Models.Spaces
{
    public sealed class BoundaryInfo
    {
        public ElementId HostElementId { get; set; }
        public string Category { get; set; }
        public string Family { get; set; }
        public string Type { get; set; }

        public string RevitTypeName { get; set; }

        public double A_Height_M { get; set; }
        public double B_Width_M { get; set; }
        public double Area_M2 { get; set; }
        public Orientation4 Orientation { get; set; }
    }
}