using Autodesk.Revit.DB;
using System.Collections.Generic;
using System.Linq;

namespace ATP_Common_Plugin.Utils
{
    public static class selecttionBuiltInInstance
    {
        /// <summary>
        /// Выбирает элементы заданной категории, отфильтровывая элементы из технических рабочих наборов.
        /// </summary>
        public static IList<Element> selectInstanceOfCategory(Autodesk.Revit.DB.Document doc, BuiltInCategory category)
        {
            IList<Element> elements = new FilteredElementCollector(doc)
                .OfCategory(category)
                .WhereElementIsNotElementType()
                .ToList();

            elements = RevitUtils.FilterTempFamilyInstance(doc, elements);
            
            return elements;
        }
    }
}
