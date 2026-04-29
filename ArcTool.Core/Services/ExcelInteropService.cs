using System;
using System.Collections.Generic;
using System.Runtime.InteropServices;
using Microsoft.Office.Interop.Excel;
using Range = Microsoft.Office.Interop.Excel.Range;
using Application = Microsoft.Office.Interop.Excel.Application;

namespace ArcTool.Core.Services
{
    /// <summary>
    /// Service quản lý giao tiếp với Excel.
    /// V5.3: Thêm GetSheetNames(), GetNamedRanges(), ExportRegion() cho V3.0.
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

        // ══════════════════════════════════════════════════════════════════════
        //  V5.3 — CÁC METHOD MỚI CHO EXCEL TO REVIT V3.0
        // ══════════════════════════════════════════════════════════════════════

        /// <summary>
        /// Lấy tên tất cả các sheet trong file Excel đang mở.
        /// Dùng để populate WorkSheet dropdown trong ExcelToRevitWindow.
        ///
        /// COM: _workbook.Worksheets trả về Sheets wrapper — phải release cả
        /// wrapper lẫn từng Worksheet cá thể để tránh leak.
        /// </summary>
        /// <returns>List tên sheet theo thứ tự xuất hiện trong workbook. Rỗng nếu chưa OpenFile.</returns>
        public List<string> GetSheetNames()
        {
            var names  = new List<string>();
            if (_workbook == null) return names;

            // Sheets là COM wrapper — phải release sau khi dùng xong
            Sheets sheets = null;
            try
            {
                sheets = _workbook.Worksheets;
                foreach (Worksheet ws in sheets)
                {
                    names.Add(ws.Name);
                    // Release từng Worksheet ngay — không tích lũy COM handles
                    Marshal.ReleaseComObject(ws);
                }
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"ExcelInteropService.GetSheetNames Error: {ex.Message}");
            }
            finally
            {
                // Release Sheets wrapper sau khi duyệt xong
                if (sheets != null) Marshal.ReleaseComObject(sheets);
            }

            return names;
        }

        /// <summary>
        /// Lấy tên tất cả Named Ranges thuộc về một sheet cụ thể.
        /// Dùng để populate Region dropdown khi user đã chọn WorkSheet.
        ///
        /// Lọc theo sheet: chỉ trả về Named Range mà RefersToRange.Worksheet.Name == sheetName.
        /// Named Ranges workbook-level hoặc span nhiều sheet sẽ bị bỏ qua (xử lý qua catch).
        ///
        /// COM: _workbook.Names trả về Names wrapper, từng Name và Range đều phải release.
        /// </summary>
        /// <param name="sheetName">Tên sheet (phân biệt hoa/thường theo Excel).</param>
        /// <returns>List tên Named Range thuộc sheet. Rỗng nếu không có hoặc chưa OpenFile.</returns>
        public List<string> GetNamedRanges(string sheetName)
        {
            var result = new List<string>();
            if (_workbook == null || string.IsNullOrWhiteSpace(sheetName)) return result;

            // Names là COM wrapper của toàn bộ Named Ranges trong workbook
            Names allNames = null;
            try
            {
                allNames = _workbook.Names;
                foreach (Name namedRange in allNames)
                {
                    try
                    {
                        // RefersToRange throw COMException nếu Named Range là formula phức tạp
                        // hoặc trỏ đến vùng đã xóa → dùng try-catch riêng cho từng mục
                        Range r = namedRange.RefersToRange;

                        // Chỉ lấy Named Range thuộc đúng sheet được chỉ định
                        // r.Worksheet có thể throw nếu range span nhiều sheet → catch bên dưới
                        if (r?.Worksheet?.Name == sheetName)
                        {
                            result.Add(namedRange.Name);
                        }

                        // Release Range COM ngay sau khi đọc xong
                        if (r != null) Marshal.ReleaseComObject(r);
                    }
                    catch
                    {
                        // Named Range không hợp lệ (formula, deleted range, cross-sheet) → bỏ qua
                        // Không propagate — một range lỗi không nên chặn các range còn lại
                    }
                    finally
                    {
                        // Release Name COM object sau mỗi iteration — child trước parent
                        Marshal.ReleaseComObject(namedRange);
                    }
                }
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"ExcelInteropService.GetNamedRanges Error: {ex.Message}");
            }
            finally
            {
                // Release Names wrapper (parent) sau khi đã release tất cả Name (child)
                if (allNames != null) Marshal.ReleaseComObject(allNames);
            }

            return result;
        }

        /// <summary>
        /// Export một vùng cụ thể trong sheet thành file PNG.
        /// Ưu tiên resolve vùng: regionName (Named Range) → Print Area → UsedRange.
        ///
        /// Chiến lược _activeSheet swap:
        ///   ExportRangeInternal() dùng _activeSheet.ChartObjects() để tạo chart tạm.
        ///   Method này tạm thời set _activeSheet = worksheet đích, gọi ExportRangeInternal(),
        ///   rồi restore về giá trị ban đầu trong finally. Không sửa ExportRangeInternal().
        ///
        /// COM: Worksheet lấy từ _workbook.Worksheets[sheetName] là COM wrapper cục bộ —
        ///   luôn release trong finally, độc lập với _activeSheet field của instance.
        /// </summary>
        /// <param name="sheetName">Tên sheet nguồn.</param>
        /// <param name="regionName">
        ///   Tên Named Range. null hoặc rỗng = bỏ qua bước này, fallback Print Area → UsedRange.
        /// </param>
        /// <param name="outputPath">Đường dẫn file PNG đầu ra (phải kết thúc bằng .png).</param>
        /// <returns>true nếu export thành công, false nếu thất bại ở bất kỳ bước nào.</returns>
        public bool ExportRegion(string sheetName, string regionName, string outputPath)
        {
            if (_workbook == null || string.IsNullOrWhiteSpace(sheetName)) return false;

            // ws là COM wrapper cục bộ — luôn release trong finally
            Worksheet ws          = null;
            Range     targetRange = null;

            // Lưu _activeSheet hiện tại để restore sau khi ExportRangeInternal chạy xong
            // Lý do: ExportRangeInternal dùng _activeSheet.ChartObjects() — cần trỏ đúng sheet
            Worksheet savedActiveSheet = _activeSheet;

            try
            {
                // Lấy worksheet theo tên (1-based hoặc by name — COM Excel hỗ trợ cả hai)
                ws = _workbook.Worksheets[sheetName] as Worksheet;
                if (ws == null)
                {
                    System.Diagnostics.Debug.WriteLine(
                        $"ExcelInteropService.ExportRegion: Sheet '{sheetName}' không tìm thấy.");
                    return false;
                }

                // Swap _activeSheet → worksheet đích trước khi gọi ExportRangeInternal
                _activeSheet = ws;

                // ── RESOLVE VÙNG EXPORT (NamedRange → PrintArea → UsedRange) ──

                // 1. Named Range — chỉ thử nếu regionName có giá trị hợp lệ
                if (!string.IsNullOrWhiteSpace(regionName))
                {
                    try
                    {
                        targetRange = ws.Range[regionName];
                    }
                    catch
                    {
                        // Named Range không tồn tại trên sheet này → fallback
                        System.Diagnostics.Debug.WriteLine(
                            $"ExcelInteropService.ExportRegion: Named Range '{regionName}' không tìm thấy trên sheet '{sheetName}'. Fallback Print Area.");
                    }
                }

                // 2. Print Area — fallback khi không có Named Range
                if (targetRange == null)
                {
                    try
                    {
                        string printArea = ws.PageSetup.PrintArea;
                        if (!string.IsNullOrEmpty(printArea))
                            targetRange = ws.Range[printArea];
                    }
                    catch
                    {
                        // PageSetup không hợp lệ (protected sheet) → fallback tiếp
                        System.Diagnostics.Debug.WriteLine(
                            $"ExcelInteropService.ExportRegion: Không đọc được Print Area của sheet '{sheetName}'. Fallback UsedRange.");
                    }
                }

                // 3. UsedRange — fallback cuối cùng, luôn có giá trị nếu sheet có data
                if (targetRange == null)
                    targetRange = ws.UsedRange;

                return ExportRangeInternal(targetRange, outputPath);
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"ExcelInteropService.ExportRegion Error: {ex.Message}");
                return false;
            }
            finally
            {
                // QUAN TRỌNG: restore _activeSheet TRƯỚC KHI release ws
                // Nếu restore sau, _activeSheet sẽ trỏ đến COM object đã bị revoke
                _activeSheet = savedActiveSheet;

                // Release theo thứ tự: targetRange (child) → ws (parent của range)
                if (targetRange != null) Marshal.ReleaseComObject(targetRange);

                // ws là COM wrapper cục bộ — release độc lập với savedActiveSheet
                // savedActiveSheet sẽ được release trong Dispose() theo lifecycle bình thường
                if (ws != null) Marshal.ReleaseComObject(ws);
            }
        }
    }
}
