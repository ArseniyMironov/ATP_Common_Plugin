using Autodesk.Revit.DB;
using Autodesk.Revit.DB.Mechanical;
using System.Collections.Generic;

namespace ATP_Common_Plugin.Services.Spaces
{
    public sealed class SpaceCollectorService
    {
        private readonly ILoggerService _logger;

        public SpaceCollectorService(ILoggerService logger) { _logger = logger; }
        public IEnumerable<Space> CollectSpaces(Document doc)
        {
            var cat = BuiltInCategory.OST_MEPSpaces;
            var col = new FilteredElementCollector(doc)
                .OfCategory(cat)
                .WhereElementIsNotElementType();

            foreach (var e in col)
            {
                Space sp = e as Space;
                if (sp != null) yield return sp;
            }
        }
    }
}