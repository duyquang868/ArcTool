using System;
using System.Collections.Generic;
using System.IO;
using System.Runtime.InteropServices;
using Microsoft.Office.Interop.Excel;
using PDFtoImage;
using SkiaSharp;
using Range = Microsoft.Office.Interop.Excel.Range;
using Application = Microsoft.Office.Interop.Excel.Application;

namespace ArcTool.Core.Services
{
    /// <summary>
    /// Service quản lý giao tiếp với Excel.
    /// V6.0: Thay pipeline CopyPicture/Clipboard/Chart bằng ExportAsFixedFormat → PDFtoImage → PNG → SkiaSharp crop.
    /// V5.3: Thêm GetSheetNames(), GetNamedRanges(), ExportRegion() cho V3.0.
    /// V5.2: Thêm GetActiveSheetName() để hỗ trợ auto-create View theo tên sheet.
    /// COM release order: child → parent. Không ReleaseComObject sau Delete().
    /// </summary>
    public class ExcelInteropService : IDisposable
    {
        private static readonly object PdfiumLoadLock = new();
        private static readonly object SkiaLoadLock = new();
        private static IntPtr _pdfiumHandle;
        private static IntPtr _skiaHandle;
        private static bool _pdfiumLoadAttempted;
        private static bool _skiaLoadAttempted;

        private Application _excelApp;
        private Workbook    _workbook;

        public bool OpenFile(string filePath)
        {
            try
            {
                _excelApp = new Application
                {
                    Visible       = false,
                    DisplayAlerts = false
                };
                _workbook = _excelApp.Workbooks.Open(filePath);
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
            Worksheet activeSheet = null;
            try
            {
                activeSheet = _workbook?.ActiveSheet as Worksheet;
                return activeSheet?.Name ?? string.Empty;
            }
            finally
            {
                if (activeSheet != null) Marshal.ReleaseComObject(activeSheet);
            }
        }

        public bool ExportPrintAreaAsHighResImage(string outputPath)
        {
            if (_workbook == null) return false;

            Worksheet activeSheet = null;
            Range targetRange = null;

            try
            {
                activeSheet = _workbook.ActiveSheet as Worksheet;
                if (activeSheet == null) return false;

                string printArea = activeSheet.PageSetup.PrintArea;
                targetRange = !string.IsNullOrEmpty(printArea)
                    ? activeSheet.Range[printArea]
                    : activeSheet.UsedRange;

                return ExportRangeInternal(activeSheet, targetRange, outputPath);
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine(
                    $"ExcelInteropService.ExportPrintAreaAsHighResImage Error: {ex.Message}");
                throw; // Propagate
            }
            finally
            {
                if (targetRange != null) Marshal.ReleaseComObject(targetRange);
                if (activeSheet != null) Marshal.ReleaseComObject(activeSheet);
            }
        }

        private static string GetRuntimeFolder()
        {
            return RuntimeInformation.ProcessArchitecture switch
            {
                Architecture.X64 => "win-x64",
                Architecture.X86 => "win-x86",
                Architecture.Arm64 => "win-arm64",
                _ => null
            };
        }

        private static string[] GetNativeLibraryCandidates(string libraryFileName)
        {
            string assemblyDir = Path.GetDirectoryName(typeof(ExcelInteropService).Assembly.Location);
            if (string.IsNullOrWhiteSpace(assemblyDir))
                return Array.Empty<string>();

            string runtimeFolder = GetRuntimeFolder();
            if (string.IsNullOrWhiteSpace(runtimeFolder))
                return Array.Empty<string>();

            return new[]
            {
                Path.Combine(assemblyDir, libraryFileName),
                Path.Combine(assemblyDir, "native", libraryFileName),
                Path.Combine(assemblyDir, "runtimes", runtimeFolder, "native", libraryFileName)
            };
        }

        private static bool EnsurePdfiumLoaded()
        {
            lock (PdfiumLoadLock)
            {
                if (_pdfiumLoadAttempted)
                    return _pdfiumHandle != IntPtr.Zero;

                _pdfiumLoadAttempted = true;

                if (NativeLibrary.TryLoad("pdfium", out _pdfiumHandle))
                    return true;

                foreach (string candidate in GetNativeLibraryCandidates("pdfium.dll"))
                {
                    if (!File.Exists(candidate))
                        continue;

                    if (NativeLibrary.TryLoad(candidate, out _pdfiumHandle))
                        return true;
                }

                return false;
            }
        }

        private static bool EnsureSkiaSharpLoaded()
        {
            lock (SkiaLoadLock)
            {
                if (_skiaLoadAttempted)
                    return _skiaHandle != IntPtr.Zero;

                _skiaLoadAttempted = true;

                if (NativeLibrary.TryLoad("libSkiaSharp", out _skiaHandle))
                    return true;

                foreach (string candidate in GetNativeLibraryCandidates("libSkiaSharp.dll"))
                {
                    if (!File.Exists(candidate))
                        continue;

                    if (NativeLibrary.TryLoad(candidate, out _skiaHandle))
                        return true;
                }

                return false;
            }
        }

        /// <summary>
        /// [CORE] Export range thành PNG qua pipeline Excel → PDF → PNG → SkiaSharp crop.
        ///
        /// V6.0:
        ///   - Bỏ hoàn toàn CopyPicture/Clipboard/Chart pipeline cũ.
        ///   - Dùng ExportAsFixedFormat(xlTypePDF) để tránh giới hạn virtual DC trong hidden mode.
        ///   - Render PDF bằng PDFtoImage (PDFium) ở 300 DPI.
        ///   - Crop white margins bằng SkiaSharp.
        /// </summary>
        private bool ExportRangeInternal(Worksheet ws, Range range, string outputPath)
        {
            PageSetup pageSetup = null;
            string tempPdf = null;

            static bool IsNonWhitePixel(SKColor color, byte threshold)
                => color.Alpha < threshold
                || color.Red < threshold
                || color.Green < threshold
                || color.Blue < threshold;

            try
            {
                pageSetup = ws.PageSetup;
                pageSetup.PrintArea      = range.Address[false, false];
                pageSetup.Zoom           = false;
                pageSetup.FitToPagesWide = 1;
                pageSetup.FitToPagesTall = 1;
                pageSetup.TopMargin      = 0;
                pageSetup.BottomMargin   = 0;
                pageSetup.LeftMargin     = 0;
                pageSetup.RightMargin    = 0;
                // xlPaperEsheet không có trong printer config của nhiều máy Windows → fallback
                try { pageSetup.PaperSize = XlPaperSize.xlPaperEsheet; }
                catch
                {
                    try { pageSetup.PaperSize = XlPaperSize.xlPaperA3; }
                    catch { /* Giữ paper size hiện tại — FitToPages vẫn handle được */ }
                }

                tempPdf = Path.Combine(Path.GetTempPath(),
                    $"ArcTool_ExcelSync_{Guid.NewGuid():N}.pdf");

                ws.ExportAsFixedFormat(
                    XlFixedFormatType.xlTypePDF,
                    tempPdf,
                    XlFixedFormatQuality.xlQualityStandard,
                    false,
                    false,
                    1,
                    1,
                    false);

                if (!EnsurePdfiumLoaded())
                {
                    System.Diagnostics.Debug.WriteLine(
                        "ExcelInteropService Export Error: pdfium.dll không load được từ add-in folder hoặc runtimes/win-*/native.");
                    return false;
                }

                using var pdfStream = File.OpenRead(tempPdf);
                Conversion.SavePng(
                    outputPath,
                    pdfStream,
                    0,
                    false,
                    null,
                    new RenderOptions { Dpi = 300, WithAnnotations = false });

                if (!EnsureSkiaSharpLoaded())
                {
                    System.Diagnostics.Debug.WriteLine(
                        "ExcelInteropService Export Error: libSkiaSharp.dll không load được từ add-in folder, native, hoặc runtimes/win-*/native.");
                    return true;
                }

                const byte WhiteThreshold = 240;
                using var bitmap = SKBitmap.Decode(outputPath);
                if (bitmap == null || bitmap.Width <= 0 || bitmap.Height <= 0)
                    return true;

                int top = 0;
                while (top < bitmap.Height)
                {
                    bool found = false;
                    for (int x = 0; x < bitmap.Width; x++)
                    {
                        if (IsNonWhitePixel(bitmap.GetPixel(x, top), WhiteThreshold))
                        {
                            found = true;
                            break;
                        }
                    }

                    if (found) break;
                    top++;
                }

                if (top >= bitmap.Height)
                    return true;

                int bottom = bitmap.Height - 1;
                while (bottom >= top)
                {
                    bool found = false;
                    for (int x = 0; x < bitmap.Width; x++)
                    {
                        if (IsNonWhitePixel(bitmap.GetPixel(x, bottom), WhiteThreshold))
                        {
                            found = true;
                            break;
                        }
                    }

                    if (found) break;
                    bottom--;
                }

                int left = 0;
                while (left < bitmap.Width)
                {
                    bool found = false;
                    for (int y = top; y <= bottom; y++)
                    {
                        if (IsNonWhitePixel(bitmap.GetPixel(left, y), WhiteThreshold))
                        {
                            found = true;
                            break;
                        }
                    }

                    if (found) break;
                    left++;
                }

                int right = bitmap.Width - 1;
                while (right >= left)
                {
                    bool found = false;
                    for (int y = top; y <= bottom; y++)
                    {
                        if (IsNonWhitePixel(bitmap.GetPixel(right, y), WhiteThreshold))
                        {
                            found = true;
                            break;
                        }
                    }

                    if (found) break;
                    right--;
                }

                int cropWidth = right - left + 1;
                int cropHeight = bottom - top + 1;
                if (cropWidth <= 0 || cropHeight <= 0)
                    return true;

                var cropRect = new SKRectI(left, top, right + 1, bottom + 1);
                var cropInfo = new SKImageInfo(cropWidth, cropHeight, bitmap.ColorType, bitmap.AlphaType);
                using var cropped = new SKBitmap(cropInfo);
                if (!bitmap.ExtractSubset(cropped, cropRect))
                    return true;

                using var image = SKImage.FromBitmap(cropped);
                using var data = image.Encode(SKEncodedImageFormat.Png, 100);
                if (data == null)
                    return true;

                using var outputStream = File.Open(outputPath, FileMode.Create, FileAccess.Write, FileShare.None);
                data.SaveTo(outputStream);
                return true;
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"ExcelInteropService Export Error: {ex.Message}");
                throw; // Propagate lên caller — không nuốt exception
            }
            finally
            {
                try { if (!string.IsNullOrEmpty(tempPdf) && File.Exists(tempPdf)) File.Delete(tempPdf); } catch { }
            }
        }

        public void Dispose()
        {
            // Release theo thứ tự: workbook → app
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
        //  V5.3 — CÁC METHOD MỚI CHO EXCEL TO REVIT V3.0 (GetSheetNames/GetNamedRanges giữ nguyên)
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
        /// V6.0: ExportRangeInternal() nhận trực tiếp Worksheet đích,
        /// nên không còn _activeSheet swap pattern.
        ///
        /// COM: Worksheet lấy từ _workbook.Worksheets[sheetName] là COM wrapper cục bộ —
        ///   luôn release trong finally. targetRange release trước ws.
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

                return ExportRangeInternal(ws, targetRange, outputPath);
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"ExcelInteropService.ExportRegion Error: {ex.Message}");
                throw; // Propagate — caller cần biết lý do thực sự để show dialog cho user
            }
            finally
            {
                // Release theo thứ tự: targetRange (child) → ws (parent của range)
                if (targetRange != null) Marshal.ReleaseComObject(targetRange);
                if (ws != null) Marshal.ReleaseComObject(ws);
            }
        }
    }
}
