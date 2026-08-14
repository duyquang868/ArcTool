using System;
using System.Collections.Generic;
using System.IO;
using System.Reflection;
using System.Runtime.InteropServices;

namespace ArcTool.Core.Services.Excel
{
    public sealed class WpsWorkbookPdfExporter : ISpreadsheetPdfExporter
    {
        private const string ProgIdKet = "KET.Application";
        private const string ProgIdEt = "ET.Application";
        private const string ProgIdKingsoftEt = "Kingsoft.ET.Application";

        private static readonly string[] ProgIdCandidates = { ProgIdKet, ProgIdEt, ProgIdKingsoftEt };

        private const int FixedFormatTypePdfValue = 0;
        private const int FixedFormatQualityStandardValue = 0;
        private const int PaperSizeEsheetValue = 26;
        private const int PaperSizeA3Value = 8;

        private const BindingFlags PublicInstance = BindingFlags.Public | BindingFlags.Instance;

        private object _app;
        private object _workbook;
        private bool _disposed;

        public SpreadsheetEngine Engine => SpreadsheetEngine.Wps;

        private static object GetProp(object target, string name, params object[] args)
            => target.GetType().InvokeMember(name, BindingFlags.GetProperty | PublicInstance, null, target, args);

        private static bool TryGetProp(object target, string name, out object value)
        {
            value = null;
            if (target == null) return false;

            try
            {
                value = GetProp(target, name);
                return value != null;
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine(
                    $"WpsWorkbookPdfExporter: property '{name}' bind failed: {ex.Message}");
                return false;
            }
        }

        private static void SetProp(object target, string name, object value)
            => target.GetType().InvokeMember(name, BindingFlags.SetProperty | PublicInstance, null, target, new object[] { value });

        private static object CallMethod(object target, string name, params object[] args)
            => target.GetType().InvokeMember(name, BindingFlags.InvokeMethod | PublicInstance, null, target, args);

        private static object[] CreateWorkbookOpenArgs(string filePath, int parameterCount)
        {
            var args = new object[parameterCount];
            args[0] = filePath;
            for (int i = 1; i < args.Length; i++)
                args[i] = Type.Missing;

            return args;
        }

        private static object TryGetWorkbookFromWindow(object activeWindow)
        {
            if (activeWindow == null) return null;

            try
            {
                var workbook = GetProp(activeWindow, "Workbook");
                if (workbook != null)
                    return workbook;
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine(
                    $"WpsWorkbookPdfExporter.Open: ActiveWindow.Workbook unavailable: {ex.Message}");
            }

            return null;
        }

        private static object TryOpenWorkbook(object workbooks, string filePath)
        {
            if (workbooks == null)
                return null;

            object[][] attempts =
            {
                CreateWorkbookOpenArgs(filePath, 15),
                CreateWorkbookOpenArgs(filePath, 13)
            };

            string[] memberNames = { "Open", "_Open" };

            for (int i = 0; i < attempts.Length; i++)
            {
                try
                {
                    var workbook = CallMethod(workbooks, memberNames[i], attempts[i]);
                    if (workbook != null)
                        return workbook;
                }
                catch (Exception ex)
                {
                    System.Diagnostics.Debug.WriteLine(
                        $"WpsWorkbookPdfExporter.Open: Workbooks.{memberNames[i]} failed: {ex.Message}");
                }
            }

            return null;
        }

        private static object AcquireWorkbookSurface(object app, string filePath, out object workbooks)
        {
            workbooks = null;

            if (TryGetProp(app, "ActiveWorkbook", out var activeWorkbook))
                return activeWorkbook;

            if (TryGetProp(app, "ActiveWindow", out var activeWindow))
            {
                try
                {
                    var windowWorkbook = TryGetWorkbookFromWindow(activeWindow);
                    if (windowWorkbook != null)
                        return windowWorkbook;
                }
                finally
                {
                    Marshal.ReleaseComObject(activeWindow);
                }
            }

            if (!TryGetProp(app, "Workbooks", out var workbookCollection))
                return null;

            workbooks = workbookCollection;
            return TryOpenWorkbook(workbookCollection, filePath);
        }

        public bool Open(string filePath)
        {
            object app = null;
            Exception lastError = null;

            foreach (var progId in ProgIdCandidates)
            {
                Type t;
                try { t = Type.GetTypeFromProgID(progId, throwOnError: false); }
                catch (Exception ex) { lastError = ex; continue; }

                if (t == null) continue;

                object candidate;
                try
                {
                    candidate = Activator.CreateInstance(t);
                }
                catch (Exception ex)
                {
                    lastError = ex;
                    System.Diagnostics.Debug.WriteLine(
                        $"WpsWorkbookPdfExporter.Open: '{progId}' resolved but activation failed: {ex.Message}");
                    continue;
                }

                try
                {
                    SetProp(candidate, "Visible", false);
                    SetProp(candidate, "DisplayAlerts", false);
                }
                catch (Exception ex)
                {
                    lastError = ex;
                    System.Diagnostics.Debug.WriteLine(
                        $"WpsWorkbookPdfExporter.Open: '{progId}' failed viability smoke test: {ex.Message}");
                    Marshal.ReleaseComObject(candidate);
                    continue;
                }

                app = candidate;
                break;
            }

            if (app == null)
            {
                System.Diagnostics.Debug.WriteLine(
                    "WpsWorkbookPdfExporter.Open: EngineAbsent" +
                    (lastError != null ? $" (last error: {lastError.Message})" : " (no WPS ProgID registered for this user)"));
                return false;
            }

            _app = app;

            object workbooks = null;
            try
            {
                var workbook = AcquireWorkbookSurface(_app, filePath, out workbooks);
                if (workbook == null)
                    return false;

                _workbook = workbook;
                return true;
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"WpsWorkbookPdfExporter.Open: EngineFoundOpenFailed: {ex.Message}");
                return false;
            }
            finally
            {
                if (workbooks != null) Marshal.ReleaseComObject(workbooks);
            }
        }

        public IReadOnlyList<string> GetSheetNames()
        {
            var result = new List<string>();
            if (_workbook == null) return result;

            object worksheets = null;
            try
            {
                worksheets = GetProp(_workbook, "Worksheets");

                int count;
                try { count = Convert.ToInt32(GetProp(worksheets, "Count")); }
                catch (Exception ex)
                {
                    System.Diagnostics.Debug.WriteLine($"WpsWorkbookPdfExporter.GetSheetNames: cannot enumerate — {ex.Message}");
                    return result;
                }

                for (int i = 1; i <= count; i++)
                {
                    object worksheet = null;
                    try
                    {
                        worksheet = GetProp(worksheets, "Item", i);
                        if (worksheet == null) continue;
                        var name = GetProp(worksheet, "Name");
                        if (name != null) result.Add(name.ToString());
                    }
                    catch (Exception ex)
                    {
                        System.Diagnostics.Debug.WriteLine($"WpsWorkbookPdfExporter.GetSheetNames: skipped sheet #{i} — {ex.Message}");
                    }
                    finally
                    {
                        if (worksheet != null) Marshal.ReleaseComObject(worksheet);
                    }
                }
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"WpsWorkbookPdfExporter.GetSheetNames: {ex.Message}");
            }
            finally
            {
                if (worksheets != null) Marshal.ReleaseComObject(worksheets);
            }

            return result;
        }

        public IReadOnlyList<string> GetNamedRanges(string sheetName)
        {
            var result = new List<string>();
            if (_workbook == null || string.IsNullOrWhiteSpace(sheetName)) return result;

            object names = null;
            try
            {
                names = GetProp(_workbook, "Names");

                int count;
                try { count = Convert.ToInt32(GetProp(names, "Count")); }
                catch (Exception ex)
                {
                    System.Diagnostics.Debug.WriteLine($"WpsWorkbookPdfExporter.GetNamedRanges: cannot enumerate — {ex.Message}");
                    return result;
                }

                for (int i = 1; i <= count; i++)
                {
                    object nameObj = null;
                    object range = null;
                    object rangeWorksheet = null;
                    try
                    {
                        nameObj = GetProp(names, "Item", i);
                        if (nameObj == null) continue;

                        string nmName;
                        try { nmName = GetProp(nameObj, "Name")?.ToString(); }
                        catch { continue; }

                        try { range = GetProp(nameObj, "RefersToRange"); }
                        catch { continue; }
                        if (range == null) continue;

                        try { rangeWorksheet = GetProp(range, "Worksheet"); }
                        catch { continue; }

                        string wsName;
                        try { wsName = GetProp(rangeWorksheet, "Name")?.ToString(); }
                        catch { continue; }

                        if (!string.IsNullOrEmpty(nmName)
                            && string.Equals(wsName, sheetName, StringComparison.OrdinalIgnoreCase))
                        {
                            result.Add(nmName);
                        }
                    }
                    catch (Exception ex)
                    {
                        System.Diagnostics.Debug.WriteLine($"WpsWorkbookPdfExporter.GetNamedRanges: skipped entry #{i} — {ex.Message}");
                    }
                    finally
                    {
                        if (rangeWorksheet != null) Marshal.ReleaseComObject(rangeWorksheet);
                        if (range != null) Marshal.ReleaseComObject(range);
                        if (nameObj != null) Marshal.ReleaseComObject(nameObj);
                    }
                }
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"WpsWorkbookPdfExporter.GetNamedRanges: {ex.Message}");
            }
            finally
            {
                if (names != null) Marshal.ReleaseComObject(names);
            }

            return result;
        }

        public bool ExportRegionToPdf(string sheetName, string regionName, string outputPdfPath)
        {
            if (_workbook == null || string.IsNullOrWhiteSpace(sheetName)) return false;

            object worksheet = null;
            object targetRange = null;
            object pageSetup;

            try
            {
                object worksheets = null;
                try
                {
                    worksheets = GetProp(_workbook, "Worksheets");
                    try
                    {
                        worksheet = GetProp(worksheets, "Item", sheetName);
                    }
                    catch (Exception ex)
                    {
                        System.Diagnostics.Debug.WriteLine(
                            $"WpsWorkbookPdfExporter.ExportRegionToPdf: sheet '{sheetName}' not found: {ex.Message}");
                        return false;
                    }
                }
                finally
                {
                    if (worksheets != null) Marshal.ReleaseComObject(worksheets);
                }

                if (worksheet == null)
                    return false;

                try
                {
                    pageSetup = GetProp(worksheet, "PageSetup");
                }
                catch (Exception ex)
                {
                    System.Diagnostics.Debug.WriteLine(
                        $"WpsWorkbookPdfExporter.ExportRegionToPdf: PageSetup unavailable: {ex.Message}");
                    return false;
                }

                if (!string.IsNullOrWhiteSpace(regionName))
                {
                    try { targetRange = GetProp(worksheet, "Range", regionName); }
                    catch (Exception ex)
                    {
                        System.Diagnostics.Debug.WriteLine(
                            $"WpsWorkbookPdfExporter.ExportRegionToPdf: named range '{regionName}' not found on '{sheetName}', falling back: {ex.Message}");
                    }
                }

                if (targetRange == null)
                {
                    try
                    {
                        var printArea = GetProp(pageSetup, "PrintArea") as string;
                        if (!string.IsNullOrEmpty(printArea))
                            targetRange = GetProp(worksheet, "Range", printArea);
                    }
                    catch (Exception ex)
                    {
                        System.Diagnostics.Debug.WriteLine(
                            $"WpsWorkbookPdfExporter.ExportRegionToPdf: print area unreadable on '{sheetName}', falling back to used range: {ex.Message}");
                    }
                }

                if (targetRange == null)
                {
                    try
                    {
                        targetRange = GetProp(worksheet, "UsedRange");
                    }
                    catch (Exception ex)
                    {
                        System.Diagnostics.Debug.WriteLine(
                            $"WpsWorkbookPdfExporter.ExportRegionToPdf: no usable range on '{sheetName}': {ex.Message}");
                        return false;
                    }
                }

                if (targetRange == null) return false;

                string address;
                try
                {
                    address = GetProp(targetRange, "Address", false, false) as string;
                }
                catch
                {
                    try
                    {
                        address = GetProp(targetRange, "Address") as string;
                    }
                    catch (Exception ex)
                    {
                        System.Diagnostics.Debug.WriteLine(
                            $"WpsWorkbookPdfExporter.ExportRegionToPdf: Range.Address unavailable: {ex.Message}");
                        return false;
                    }
                }

                if (string.IsNullOrEmpty(address))
                    return false;

                try
                {
                    SetProp(pageSetup, "PrintArea", address);
                    SetProp(pageSetup, "Zoom", false);
                    SetProp(pageSetup, "FitToPagesWide", 1);
                    SetProp(pageSetup, "FitToPagesTall", 1);
                    SetProp(pageSetup, "TopMargin", 0.0);
                    SetProp(pageSetup, "BottomMargin", 0.0);
                    SetProp(pageSetup, "LeftMargin", 0.0);
                    SetProp(pageSetup, "RightMargin", 0.0);
                }
                catch (Exception ex)
                {
                    System.Diagnostics.Debug.WriteLine(
                        $"WpsWorkbookPdfExporter.ExportRegionToPdf: page setup normalization failed: {ex.Message}");
                    return false;
                }

                try { SetProp(pageSetup, "PaperSize", PaperSizeEsheetValue); }
                catch
                {
                    try { SetProp(pageSetup, "PaperSize", PaperSizeA3Value); }
                    catch { }
                }

                try
                {
                    CallMethod(worksheet, "ExportAsFixedFormat",
                        FixedFormatTypePdfValue, outputPdfPath, FixedFormatQualityStandardValue, false, false, 1, 1, false);
                }
                catch (Exception ex)
                {
                    System.Diagnostics.Debug.WriteLine(
                        $"WpsWorkbookPdfExporter.ExportRegionToPdf: PDF export unsupported on this WPS build: {ex.Message}");
                    return false;
                }

                if (!File.Exists(outputPdfPath))
                    return false;

                return true;
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"WpsWorkbookPdfExporter.ExportRegionToPdf: {ex.Message}");
                return false;
            }
            finally
            {
                if (targetRange != null) Marshal.ReleaseComObject(targetRange);
                if (worksheet != null) Marshal.ReleaseComObject(worksheet);
            }
        }

        public void Dispose()
        {
            if (_disposed) return;
            _disposed = true;

            if (_workbook != null)
            {
                try { CallMethod(_workbook, "Close", false); }
                catch (Exception ex)
                {
                    System.Diagnostics.Debug.WriteLine($"WpsWorkbookPdfExporter.Dispose: Close failed (ignored): {ex.Message}");
                }
                finally
                {
                    try { Marshal.ReleaseComObject(_workbook); } catch { }
                    _workbook = null;
                }
            }

            if (_app != null)
            {
                try { CallMethod(_app, "Quit"); }
                catch (Exception ex)
                {
                    System.Diagnostics.Debug.WriteLine($"WpsWorkbookPdfExporter.Dispose: Quit failed (ignored): {ex.Message}");
                }
                finally
                {
                    try { Marshal.ReleaseComObject(_app); } catch { }
                    _app = null;
                }
            }
        }
    }
}
