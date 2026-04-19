using System;
using System.Runtime.InteropServices;
using Microsoft.Office.Interop.Excel;
using Range = Microsoft.Office.Interop.Excel.Range;
using Application = Microsoft.Office.Interop.Excel.Application;

namespace ArcTool.Core.Services
{
    /// <summary>
    /// Service quản lý giao tiếp với Excel.
    /// V5.2: Thêm GetActiveSheetName() để hỗ trợ auto-create View theo tên sheet.
    /// COM release order: child → parent. Không ReleaseComObject sau Delete().
    /// </summary>
    public class ExcelInteropService : IDisposable
    {
        private Application _excelApp;
        private Workbook    _workbook;
        private Worksheet   _activeSheet;

        // Scale 35x cho chất lượng ảnh cao
        private const double FIXED_SCALE_FACTOR  = 35.0;
        private const double MAX_EXCEL_DIMENSION = 32000;

        public bool OpenFile(string filePath)
        {
            try
            {
                _excelApp = new Application
                {
                    Visible       = false,
                    DisplayAlerts = false
                };
                _workbook    = _excelApp.Workbooks.Open(filePath);
                _activeSheet = _workbook.ActiveSheet as Worksheet;
                return true;
            }
            catch
            {
                return false;
            }
        }

        /// <summary>
        /// Trả về tên sheet đang active trong file Excel.
        /// Dùng để đặt tên Drafting View tương ứng trong Revit.
        /// </summary>
        public string GetActiveSheetName()
        {
            return _activeSheet?.Name ?? string.Empty;
        }

        public bool ExportPrintAreaAsHighResImage(string outputPath)
        {
            if (_activeSheet == null) return false;
            Range targetRange = null;

            try
            {
                string printArea = _activeSheet.PageSetup.PrintArea;
                targetRange = !string.IsNullOrEmpty(printArea)
                    ? _activeSheet.Range[printArea]
                    : _activeSheet.UsedRange;

                return ExportRangeInternal(targetRange, outputPath);
            }
            catch
            {
                return false;
            }
            finally
            {
                if (targetRange != null) Marshal.ReleaseComObject(targetRange);
            }
        }

        /// <summary>
        /// [CORE] Tuần tự, không Sleep, không retry. Scale 35x cố định.
        /// </summary>
        private bool ExportRangeInternal(Range range, string outputPath)
        {
            ChartObjects chartObjects = null;
            ChartObject  chartObj     = null;
            Chart        chart        = null;

            try
            {
                range.CopyPicture(XlPictureAppearance.xlPrinter, XlCopyPictureFormat.xlPicture);

                double originalWidth  = (double)range.Width;
                double originalHeight = (double)range.Height;

                chartObjects = (ChartObjects)_activeSheet.ChartObjects();
                chartObj     = chartObjects.Add(0, 0, originalWidth, originalHeight);
                chart        = chartObj.Chart;

                chartObj.Activate();
                chart.Paste();

                double newWidth  = Math.Min(originalWidth  * FIXED_SCALE_FACTOR, MAX_EXCEL_DIMENSION);
                double newHeight = Math.Min(originalHeight * FIXED_SCALE_FACTOR, MAX_EXCEL_DIMENSION);

                chartObj.Width  = newWidth;
                chartObj.Height = newHeight;

                chart.Export(outputPath, "PNG", false);
                return true;
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"ExcelInteropService Export Error: {ex.Message}");
                return false;
            }
            finally
            {
                // COM release: child → parent. KHÔNG release chartObj sau Delete().
                ReleaseObject(chart);
                if (chartObj != null)
                {
                    try { chartObj.Delete(); } catch { }
                    // Không gọi ReleaseObject(chartObj) — Delete đã revoke COM
                }
                ReleaseObject(chartObjects);
            }
        }

        public void Dispose()
        {
            // Release theo thứ tự: sheet → workbook → app
            if (_activeSheet != null)
            {
                ReleaseObject(_activeSheet);
                _activeSheet = null;
            }
            if (_workbook != null)
            {
                try { _workbook.Close(false); } catch { }
                ReleaseObject(_workbook);
                _workbook = null;
            }
            if (_excelApp != null)
            {
                try { _excelApp.Quit(); } catch { }
                ReleaseObject(_excelApp);
                _excelApp = null;
            }
            GC.Collect();
            GC.WaitForPendingFinalizers();
        }

        private void ReleaseObject(object obj)
        {
            try
            {
                if (obj != null) Marshal.ReleaseComObject(obj);
            }
            catch { }
            // KHÔNG null obj ở đây — null field gốc ở caller mới có tác dụng
        }
    }
}
