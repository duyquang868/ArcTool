using System;
using System.Collections.Generic;
using System.IO;

namespace ArcTool.Core.Services.Excel
{
    public class SpreadsheetImageExportService : IDisposable
    {
        private ISpreadsheetPdfExporter _exporter;
        private bool _disposed;

        public bool OpenFile(string filePath)
        {
            if (_disposed)
                throw new ObjectDisposedException(nameof(SpreadsheetImageExportService));

            DisposeExporter();

            ISpreadsheetPdfExporter[] candidates =
            {
                new MsExcelWorkbookPdfExporter(),
                new WpsWorkbookPdfExporter()
            };

            foreach (ISpreadsheetPdfExporter candidate in candidates)
            {
                try
                {
                    if (candidate.Open(filePath))
                    {
                        _exporter = candidate;
                        return true;
                    }
                }
                catch (Exception ex)
                {
                    System.Diagnostics.Debug.WriteLine(
                        $"SpreadsheetImageExportService.OpenFile: {candidate.Engine} threw while opening '{filePath}': {ex.Message}");
                }

                try
                {
                    candidate.Dispose();
                }
                catch
                {
                }
            }

            System.Diagnostics.Debug.WriteLine(
                $"SpreadsheetImageExportService.OpenFile: no spreadsheet engine could open '{filePath}'.");
            return false;
        }

        public List<string> GetSheetNames()
        {
            return _exporter == null
                ? new List<string>()
                : new List<string>(_exporter.GetSheetNames());
        }

        public List<string> GetNamedRanges(string sheetName)
        {
            return _exporter == null
                ? new List<string>()
                : new List<string>(_exporter.GetNamedRanges(sheetName));
        }

        public bool ExportRegion(string sheetName, string regionName, string outputPngPath)
        {
            if (_exporter == null)
                return false;

            string tempPdfPath = Path.Combine(
                Path.GetTempPath(),
                $"ArcTool_ExcelSync_{Guid.NewGuid():N}.pdf");

            try
            {
                if (!_exporter.ExportRegionToPdf(sheetName, regionName, tempPdfPath))
                {
                    System.Diagnostics.Debug.WriteLine(
                        $"SpreadsheetImageExportService.ExportRegion: PDF export failed for engine '{_exporter.Engine}', sheet='{sheetName}', region='{regionName ?? "(null)"}'.");
                    return false;
                }

                bool rasterSucceeded = PdfRasterImageService.RenderPdfToCroppedPng(tempPdfPath, outputPngPath);
                if (!rasterSucceeded)
                {
                    System.Diagnostics.Debug.WriteLine(
                        $"SpreadsheetImageExportService.ExportRegion: raster stage failed for '{tempPdfPath}'.");
                }

                return rasterSucceeded;
            }
            finally
            {
                TryDeleteTempFile(tempPdfPath, "temp PDF");
            }
        }

        public void Dispose()
        {
            if (_disposed) return;
            _disposed = true;
            DisposeExporter();
        }

        private void DisposeExporter()
        {
            if (_exporter == null) return;

            try
            {
                _exporter.Dispose();
            }
            catch
            {
            }
            finally
            {
                _exporter = null;
            }
        }

        private static void TryDeleteTempFile(string filePath, string fileRole)
        {
            if (string.IsNullOrWhiteSpace(filePath))
                return;

            try
            {
                if (File.Exists(filePath))
                    File.Delete(filePath);
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine(
                    $"SpreadsheetImageExportService.ExportRegion: could not delete {fileRole} '{filePath}': {ex.Message}");
            }
        }
    }
}
