using System.Collections.Generic;

namespace ATP_Common_Plugin.Models.Spaces
{
    public sealed class SpaceInfo
    {
        public string Name { get; set; }
        public string Number { get; set; }
        public double Area_M2 { get; set; }

        public List<BoundaryInfo> Boundaries { get; set; } = new List<BoundaryInfo>();
    }
}