using ATP_Common_Plugin.Utils;
using Autodesk.Revit.Attributes;
using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using Microsoft.Office.Interop.Excel;
using System;
using System.Collections.Generic;
using System.Linq;

namespace ATP_Common_Plugin.Commands
{
    [Regeneration(RegenerationOption.Manual)]
    [Transaction(TransactionMode.Manual)]
    class MarkPlumbFix : IExternalCommand
    {
        public Result Execute(ExternalCommandData commandData, ref string message, ElementSet elements)
        {
            UIApplication uiApp = commandData.Application;
            UIDocument uiDoc = uiApp.ActiveUIDocument;
            Document doc = uiDoc.Document;
            //var logger = ATP_App.GetService<ILoggerService>();
            try
            {
                IList<Element> fixt = selecttionBuiltInInstance.selectInstanceOfCategory(doc, BuiltInCategory.OST_PlumbingFixtures);

                IList<Element> filteredFixt = new List<Element>();

                foreach (var elem in fixt)
                {
                    string model = doc.GetElement(elem.GetTypeId()).get_Parameter(BuiltInParameter.ALL_MODEL_MODEL).AsValueString();
                    if (model.Contains("Трап"))
                        filteredFixt.Add(elem);
                }

                Dictionary<string, List<Element>> sysTypeDict = new Dictionary<string, List<Element>>();
                foreach (var el in filteredFixt)
                {
                    string systemType = el.get_Parameter(BuiltInParameter.RBS_PIPING_SYSTEM_TYPE_PARAM).AsValueString();
                    sysTypeDict[systemType].Append(el);
                }

                using (Transaction tr = new Transaction(doc, "Маркировка Plumbing Fixtures"))
                {
                    tr.Start();

                    foreach (var group in sysTypeDict.Keys)
                    {
                        int num = 1;
                        var sortedList = sysTypeDict[group]
                            .OrderBy(x => x.get_Geometry(null).GetBoundingBox().Min.Z)
                            .OrderBy(x => x.get_Geometry(null).GetBoundingBox().Min.X)
                            .OrderBy(x => x.get_Geometry(null).GetBoundingBox().Min.Y);
                        foreach (var elem in sysTypeDict[group])
                        {
                            elem.get_Parameter(BuiltInParameter.ALL_MODEL_INSTANCE_COMMENTS).Set($"ВВ-{num}"); // Заменить на маркировка скрипт
                            num++;
                        }
                    }

                    tr.Commit();
                }
                return Result.Succeeded;
            }
            catch (Exception ex)
            {
                //logger.LogError($"Ошибка! {ex.Message}", docName);
                return Result.Failed;
            }
        }
    }
}
