using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using ClosedXML.Excel;

namespace ArcTool.Core.Services.Excel
{
    internal sealed class WorkbookRegionSnapshotService : IDisposable
    {
        private XLWorkbook _workbook;
        private string _sourceFilePath;
        private bool _disposed;

        public bool Open(string filePath)
        {
            if (_disposed)
                throw new ObjectDisposedException(nameof(WorkbookRegionSnapshotService));

            DisposeWorkbook();

            if (string.IsNullOrWhiteSpace(filePath) || !File.Exists(filePath))
                return false;

            try
            {
                _workbook = new XLWorkbook(filePath);
                _sourceFilePath = filePath;
                return true;
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine(
                    $"WorkbookRegionSnapshotService.Open: failed to open '{filePath}': {ex.Message}");
                DisposeWorkbook();
                return false;
            }
        }

        public IReadOnlyList<string> GetSheetNames()
        {
            if (_workbook == null)
                return Array.Empty<string>();

            try
            {
                return _workbook.Worksheets
                    .Select(ws => ws.Name)
                    .ToList();
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine(
                    $"WorkbookRegionSnapshotService.GetSheetNames: failed for '{_sourceFilePath}': {ex.Message}");
                return Array.Empty<string>();
            }
        }

        public IReadOnlyList<string> GetNamedRanges(string sheetName)
        {
            if (_workbook == null || string.IsNullOrWhiteSpace(sheetName))
                return Array.Empty<string>();

            try
            {
                IXLWorksheet ws = FindWorksheet(sheetName);
                if (ws == null)
                    return Array.Empty<string>();

                return ws.DefinedNames
                    .Select(name => name.Name)
                    .Where(name => !string.IsNullOrWhiteSpace(name))
                    .Distinct(StringComparer.OrdinalIgnoreCase)
                    .OrderBy(name => name, StringComparer.OrdinalIgnoreCase)
                    .ToList();
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine(
                    $"WorkbookRegionSnapshotService.GetNamedRanges: failed for sheet='{sheetName}' in '{_sourceFilePath}': {ex.Message}");
                return Array.Empty<string>();
            }
        }

        public bool CreateRegionWorkbook(string sheetName, string regionName, string outputWorkbookPath)
        {
            if (_workbook == null || string.IsNullOrWhiteSpace(sheetName) || string.IsNullOrWhiteSpace(outputWorkbookPath))
                return false;

            try
            {
                IXLWorksheet sourceSheet = FindWorksheet(sheetName);
                if (sourceSheet == null)
                {
                    System.Diagnostics.Debug.WriteLine(
                        $"WorkbookRegionSnapshotService.CreateRegionWorkbook: sheet '{sheetName}' not found in '{_sourceFilePath}'.");
                    return false;
                }

                if (!TryResolveSourceRange(sourceSheet, regionName, out IXLRange sourceRange, out string resolvedRegionType))
                {
                    System.Diagnostics.Debug.WriteLine(
                        $"WorkbookRegionSnapshotService.CreateRegionWorkbook: no exportable range for sheet='{sheetName}', region='{regionName ?? "(null)"}'.");
                    return false;
                }

                string outputDir = Path.GetDirectoryName(outputWorkbookPath);
                if (!string.IsNullOrWhiteSpace(outputDir))
                    Directory.CreateDirectory(outputDir);

                using (var tempWorkbook = new XLWorkbook())
                {
                    IXLWorksheet targetSheet = tempWorkbook.Worksheets.Add(SanitizeWorksheetName(sourceSheet.Name));
                    CopyRangeToTemporarySheet(sourceRange, targetSheet);
                    NormalizePageSetup(targetSheet);

                    string targetAddress = targetSheet.Range(
                        1,
                        1,
                        sourceRange.RowCount(),
                        sourceRange.ColumnCount()).RangeAddress.ToString();
                    targetSheet.PageSetup.PrintAreas.Clear();
                    targetSheet.PageSetup.PrintAreas.Add(targetAddress);

                    tempWorkbook.SaveAs(outputWorkbookPath);
                }

                System.Diagnostics.Debug.WriteLine(
                    $"WorkbookRegionSnapshotService.CreateRegionWorkbook: created temp workbook for {resolvedRegionType} at '{outputWorkbookPath}'.");
                return true;
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine(
                    $"WorkbookRegionSnapshotService.CreateRegionWorkbook: failed for sheet='{sheetName}', region='{regionName ?? "(null)"}': {ex.Message}");
                return false;
            }
        }

        public void Dispose()
        {
            if (_disposed) return;
            _disposed = true;
            DisposeWorkbook();
        }

        private void DisposeWorkbook()
        {
            if (_workbook == null)
                return;

            try
            {
                _workbook.Dispose();
            }
            catch
            {
            }
            finally
            {
                _workbook = null;
                _sourceFilePath = null;
            }
        }

        private IXLWorksheet FindWorksheet(string sheetName)
        {
            return _workbook.Worksheets.FirstOrDefault(
                ws => string.Equals(ws.Name, sheetName, StringComparison.OrdinalIgnoreCase));
        }

        private bool TryResolveSourceRange(
            IXLWorksheet sourceSheet,
            string regionName,
            out IXLRange sourceRange,
            out string resolvedRegionType)
        {
            sourceRange = null;
            resolvedRegionType = null;

            if (!string.IsNullOrWhiteSpace(regionName)
                && TryGetNamedRange(sourceSheet, regionName, out sourceRange))
            {
                resolvedRegionType = "NamedRange";
                return true;
            }

            if (TryGetPrintArea(sourceSheet, out sourceRange))
            {
                resolvedRegionType = "PrintArea";
                return true;
            }

            if (TryGetUsedRange(sourceSheet, out sourceRange))
            {
                resolvedRegionType = "UsedRange";
                return true;
            }

            return false;
        }

        private static bool TryGetNamedRange(IXLWorksheet sourceSheet, string regionName, out IXLRange range)
        {
            range = null;

            try
            {
                var definedName = sourceSheet.DefinedNames.FirstOrDefault(
                    name => string.Equals(name.Name, regionName, StringComparison.OrdinalIgnoreCase));
                if (definedName == null)
                    return false;

                range = definedName.Ranges.FirstOrDefault();
                return range != null;
            }
            catch
            {
                return false;
            }
        }

        private static bool TryGetPrintArea(IXLWorksheet sourceSheet, out IXLRange range)
        {
            range = null;

            try
            {
                range = sourceSheet.PageSetup.PrintAreas.FirstOrDefault();
                return range != null;
            }
            catch
            {
                return false;
            }
        }

        private static bool TryGetUsedRange(IXLWorksheet sourceSheet, out IXLRange range)
        {
            range = null;

            try
            {
                range = sourceSheet.RangeUsed();
                return range != null;
            }
            catch
            {
                return false;
            }
        }

        private static void CopyRangeToTemporarySheet(IXLRange sourceRange, IXLWorksheet targetSheet)
        {
            sourceRange.CopyTo(targetSheet.Cell(1, 1));

            for (int offset = 0; offset < sourceRange.ColumnCount(); offset++)
            {
                targetSheet.Column(offset + 1).Width = sourceRange.Worksheet.Column(sourceRange.RangeAddress.FirstAddress.ColumnNumber + offset).Width;
            }

            for (int offset = 0; offset < sourceRange.RowCount(); offset++)
            {
                targetSheet.Row(offset + 1).Height = sourceRange.Worksheet.Row(sourceRange.RangeAddress.FirstAddress.RowNumber + offset).Height;
            }
        }

        private static void NormalizePageSetup(IXLWorksheet targetSheet)
        {
            targetSheet.PageSetup.FitToPages(1, 1);
            targetSheet.PageSetup.CenterHorizontally = false;
            targetSheet.PageSetup.CenterVertically = false;
            targetSheet.PageSetup.Margins.Top = 0;
            targetSheet.PageSetup.Margins.Bottom = 0;
            targetSheet.PageSetup.Margins.Left = 0;
            targetSheet.PageSetup.Margins.Right = 0;
            targetSheet.PageSetup.Margins.Header = 0;
            targetSheet.PageSetup.Margins.Footer = 0;
        }

        private static string SanitizeWorksheetName(string sheetName)
        {
            if (string.IsNullOrWhiteSpace(sheetName))
                return "Sheet1";

            string sanitized = sheetName;
            foreach (char invalidChar in Path.GetInvalidFileNameChars())
                sanitized = sanitized.Replace(invalidChar.ToString(), string.Empty);

            sanitized = sanitized
                .Replace("[", string.Empty)
                .Replace("]", string.Empty)
                .Replace(":", string.Empty)
                .Replace("*", string.Empty)
                .Replace("?", string.Empty)
                .Replace("/", string.Empty)
                .Replace("\\", string.Empty);

            if (sanitized.Length > 31)
                sanitized = sanitized.Substring(0, 31);

            return string.IsNullOrWhiteSpace(sanitized) ? "Sheet1" : sanitized;
        }
    }
}
