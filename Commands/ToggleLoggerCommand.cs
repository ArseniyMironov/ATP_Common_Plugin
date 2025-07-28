using ATP_Common_Plugin.Services;
using Autodesk.Revit.Attributes;
using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using System;

namespace ATP_Common_Plugin.Commands
{
    [Transaction(TransactionMode.Manual)]
    class ToggleLoggerCommand : IExternalCommand
    {
        public Result Execute(ExternalCommandData commandData, ref string message, ElementSet elements)
        {
            try
            {
                var logger = ATP_App.GetService<ILoggerService>();

                // Диагностическая информация
                string debugInfo = $"IsWindowVisible: {logger.IsWindowVisible}\n" +
                                  $"Window instance: {(logger.IsWindowVisible ? "exists" : "null")}";

                if (logger.IsWindowVisible)
                {
                    logger.HideWindow();
                    //TaskDialog.Show("Debug", $"Hiding window\n{debugInfo}");
                }
                else
                {
                    logger.ShowWindow();
                    //TaskDialog.Show("Debug", $"Showing window\n{debugInfo}");
                }

                return Result.Succeeded;
            }
            catch (Exception ex)
            {
                message = ex.Message;
                return Result.Failed;
            }
        }
    }
}
