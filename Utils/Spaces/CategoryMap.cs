using Autodesk.Revit.DB;

namespace ATP_Common_Plugin.Utils.Spaces
{
    public static class CategoryMap
    {
        public static string GetCategoryName(Element e)
        {
            return (e != null && e.Category != null) ? e.Category.Name : string.Empty;
        }
    }
}