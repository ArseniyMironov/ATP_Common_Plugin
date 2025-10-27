using ATP_Common_Plugin.Models.Spaces;
using ATP_Common_Plugin.Services;
using System.Collections.Generic;

namespace ATP_Common_Plugin.Utils.Excel
{
    /// <summary>
    /// Minimal Excel Interop writer. Release COM objects carefully.
    /// </summary>
    public sealed class ExcelExportService
    {
        private readonly ILoggerService _logger;

        public ExcelExportService(ILoggerService logger) { _logger = logger; }

        public void Export(IList<SpaceInfo> spaces)
        {
            if (spaces == null || spaces.Count == 0) return;

            Microsoft.Office.Interop.Excel.Application app = null;
            Microsoft.Office.Interop.Excel.Workbook wb = null;
            Microsoft.Office.Interop.Excel.Worksheet ws = null;

            try
            {
                app = new Microsoft.Office.Interop.Excel.Application();
                wb = app.Workbooks.Add();
                ws = (Microsoft.Office.Interop.Excel.Worksheet)wb.ActiveSheet;
                ws.Name = Models.Settings.SheetName;

                int row = 1;
                // headers
                ws.Cells[row, 1] = Models.Settings.HeaderRoom;
                ws.Cells[row, 2] = Models.Settings.HeaderArea;
                ws.Cells[row, 3] = Models.Settings.HeaderBoundary;
                ws.Cells[row, 4] = Models.Settings.HeaderA;
                ws.Cells[row, 5] = Models.Settings.HeaderB;
                ws.Cells[row, 6] = Models.Settings.HeaderS;
                ws.Cells[row, 7] = Models.Settings.HeaderOri;
                row++;

                foreach (var sp in spaces)
                {
                    // space header row
                    ws.Cells[row, 1] = sp.Name + " (" + sp.Number + ")";
                    ws.Cells[row, 2] = sp.Area_M2;
                    row++;

                    // boundaries
                    int idx = 1;
                    foreach (var b in sp.Boundaries)
                    {
                        string cat = b.Category ?? string.Empty;
                        string tnm = b.RevitTypeName ?? string.Empty; 
                        
                        string label;
                        if (!string.IsNullOrEmpty(cat) || !string.IsNullOrEmpty(tnm))
                        {
                            // "[Category]-[Type]" (без лишних дефисов, если чего-то нет)
                            if (string.IsNullOrEmpty(cat)) label = tnm;
                            else if (string.IsNullOrEmpty(tnm)) label = cat;
                            else label = $"{cat}-{tnm}";
                        }
                        else
                        {
                            // фоллбек, если нет хоста (Free boundary) или не удалось определить
                            label = "FreeBoundary";
                        }

                        ws.Cells[row, 3] = label;
                        idx++;

                        ws.Cells[row, 4] = b.A_Height_M;
                        ws.Cells[row, 5] = b.B_Width_M;
                        ws.Cells[row, 6] = b.Area_M2;
                        ws.Cells[row, 7] = b.Orientation.ToString();
                        row++;
                    }
                }

                app.Visible = true; // optional: show to user
            }
            finally
            {
                // Do not save automatically; user can decide in Excel UI.
                Release(ws);
                Release(wb);
                Release(app);
            }
        }

        private static void Release(object com)
        {
            try
            {
                if (com != null) System.Runtime.InteropServices.Marshal.ReleaseComObject(com);
            }
            catch { }
        }
    }
}