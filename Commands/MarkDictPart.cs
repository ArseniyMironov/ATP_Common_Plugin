using Autodesk.Revit.Attributes;
using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using System;
using System.Collections.Generic;
using System.IO;

namespace ATP_Common_Plugin.Commands
{
    [Regeneration(RegenerationOption.Manual)]
    [Transaction(TransactionMode.Manual)]
    class MarkDictPart : IExternalCommand
    {
        public Result Execute(ExternalCommandData commandData, ref string message, ElementSet elements)
        {


            return Result.Succeeded;
        }
        private List<string> ReadExcell(string path)
        {
            var result = new List<string> ();

            try
            {
                var fileInfo = new FileInfo(path);

                if (!fileInfo.Exists)
                    throw new FileNotFoundException("Файл не найден", path);

            }
            catch (Exception ex)
            {
                throw new Exception($"Ошибка EPPlus: {ex.Message}");
            }

            return result;
        }
    }
}
