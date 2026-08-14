using System;
using System.Collections.Generic;
using System.Runtime.InteropServices;
using Microsoft.Office.Interop.Excel;

namespace ArcTool.Core.Services.Excel
{
    public class MsExcelWorkbookPdfExporter : ISpreadsheetPdfExporter
    {
        private Microsoft.Office.Interop.Excel.Application _excelApp;
        private Workbook _workbook;

        public SpreadsheetEngine Engine => SpreadsheetEngine.MsExcel;

        public bool Open(string filePath)
        {
            try
            {
                _excelApp = new Microsoft.Office.Interop.Excel.Application
                {
                    Visible = false,
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

        public IReadOnlyList<string> GetSheetNames()
        {
            var names = new List<string>();
            if (_workbook == null) return names;

            Sheets sheets = null;
            try
            {
                sheets = _workbook.Worksheets;
                foreach (Worksheet ws in sheets)
                {
                    names.Add(ws.Name);
                    Marshal.ReleaseComObject(ws);
                }
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"MsExcelWorkbookPdfExporter.GetSheetNames Error: {ex.Message}");
            }
            finally
            {
                if (sheets != null) Marshal.ReleaseComObject(sheets);
            }

            return names;
        }

        public IReadOnlyList<string> GetNamedRanges(string sheetName)
        {
            var result = new List<string>();
            if (_workbook == null || string.IsNullOrWhiteSpace(sheetName)) return result;

            Names allNames = null;
            try
            {
                allNames = _workbook.Names;
                foreach (Name namedRange in allNames)
                {
                    try
                    {
                        Microsoft.Office.Interop.Excel.Range r = namedRange.RefersToRange;
                        if (r?.Worksheet?.Name == sheetName)
                        {
                            result.Add(namedRange.Name);
                        }

                        if (r != null) Marshal.ReleaseComObject(r);
                    }
                    catch
                    {
                    }
                    finally
                    {
                        Marshal.ReleaseComObject(namedRange);
                    }
                }
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"MsExcelWorkbookPdfExporter.GetNamedRanges Error: {ex.Message}");
            }
            finally
            {
                if (allNames != null) Marshal.ReleaseComObject(allNames);
            }

            return result;
        }

        public bool ExportRegionToPdf(string sheetName, string regionName, string outputPdfPath)
        {
            if (_workbook == null || string.IsNullOrWhiteSpace(sheetName)) return false;

            Worksheet ws = null;
            Microsoft.Office.Interop.Excel.Range targetRange = null;

            try
            {
                ws = _workbook.Worksheets[sheetName] as Worksheet;
                if (ws == null)
                {
                    System.Diagnostics.Debug.WriteLine(
                        $"MsExcelWorkbookPdfExporter.ExportRegionToPdf: Sheet '{sheetName}' not found.");
                    return false;
                }

                if (!string.IsNullOrWhiteSpace(regionName))
                {
                    try
                    {
                        targetRange = ws.Range[regionName];
                    }
                    catch
                    {
                        System.Diagnostics.Debug.WriteLine(
                            $"MsExcelWorkbookPdfExporter.ExportRegionToPdf: Named Range '{regionName}' not found on sheet '{sheetName}'. Falling back to Print Area.");
                    }
                }

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
                        System.Diagnostics.Debug.WriteLine(
                            $"MsExcelWorkbookPdfExporter.ExportRegionToPdf: Could not read Print Area of sheet '{sheetName}'. Falling back to UsedRange.");
                    }
                }

                if (targetRange == null)
                    targetRange = ws.UsedRange;

                return ExportRangeToPdf(ws, targetRange, outputPdfPath);
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"MsExcelWorkbookPdfExporter.ExportRegionToPdf Error: {ex.Message}");
                return false;
            }
            finally
            {
                if (targetRange != null) Marshal.ReleaseComObject(targetRange);
                if (ws != null) Marshal.ReleaseComObject(ws);
            }
        }

        private bool ExportRangeToPdf(Worksheet ws, Microsoft.Office.Interop.Excel.Range range, string outputPdfPath)
        {
            try
            {
                PageSetup pageSetup = ws.PageSetup;
                pageSetup.PrintArea = range.Address[false, false];
                pageSetup.Zoom = false;
                pageSetup.FitToPagesWide = 1;
                pageSetup.FitToPagesTall = 1;
                pageSetup.TopMargin = 0;
                pageSetup.BottomMargin = 0;
                pageSetup.LeftMargin = 0;
                pageSetup.RightMargin = 0;

                try { pageSetup.PaperSize = XlPaperSize.xlPaperEsheet; }
                catch
                {
                    try { pageSetup.PaperSize = XlPaperSize.xlPaperA3; }
                    catch { }
                }

                ws.ExportAsFixedFormat(
                    XlFixedFormatType.xlTypePDF,
                    outputPdfPath,
                    XlFixedFormatQuality.xlQualityStandard,
                    false,
                    false,
                    1,
                    1,
                    false);

                return true;
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"MsExcelWorkbookPdfExporter.ExportRangeToPdf Error: {ex.Message}");
                return false;
            }
        }

        public void Dispose()
        {
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
        }
    }
}
