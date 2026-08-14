using System;
using System.IO;
using System.Linq;
using ArcTool.Core.Services.Excel;

namespace ArcTool.Core.Tools
{
    internal static class ExcelExportSmokeHarness
    {
        public static int Main(string[] args)
        {
            if (args == null || args.Length < 1)
            {
                Console.Error.WriteLine("Usage: ExcelExportSmokeHarness <xlsx-path> [sheet-name] [region-name] [output-png-path]");
                return 2;
            }

            string workbookPath = args[0];
            string requestedSheetName = args.Length > 1 ? args[1] : null;
            string regionName = args.Length > 2 ? args[2] : null;
            string outputPngPath = args.Length > 3
                ? args[3]
                : Path.Combine(Path.GetTempPath(), $"ArcTool_ExcelSmoke_{Guid.NewGuid():N}.png");

            Console.WriteLine($"Workbook: {workbookPath}");
            Console.WriteLine($"Requested sheet: {requestedSheetName ?? "(auto-first-sheet)"}");
            Console.WriteLine($"Requested region: {regionName ?? "(PrintArea/UsedRange fallback)"}");
            Console.WriteLine($"Output PNG: {outputPngPath}");

            try
            {
                using (var svc = new SpreadsheetImageExportService())
                {
                    if (!svc.OpenFile(workbookPath))
                    {
                        Console.Error.WriteLine("OpenFile returned false.");
                        return 1;
                    }

                    var sheetNames = svc.GetSheetNames();
                    Console.WriteLine("Sheets: " + string.Join(" | ", sheetNames));
                    if (sheetNames.Count == 0)
                    {
                        Console.Error.WriteLine("No sheets were discovered.");
                        return 1;
                    }

                    string sheetName = requestedSheetName;
                    if (string.IsNullOrWhiteSpace(sheetName))
                        sheetName = sheetNames[0];

                    Console.WriteLine($"Using sheet: {sheetName}");
                    var namedRanges = svc.GetNamedRanges(sheetName);
                    Console.WriteLine("Named ranges: " + (namedRanges.Count == 0 ? "(none)" : string.Join(" | ", namedRanges)));

                    bool exportSucceeded = svc.ExportRegion(sheetName, regionName, outputPngPath);
                    Console.WriteLine($"ExportRegion returned: {exportSucceeded}");
                    Console.WriteLine($"PNG exists: {File.Exists(outputPngPath)}");
                    if (File.Exists(outputPngPath))
                    {
                        var fileInfo = new FileInfo(outputPngPath);
                        Console.WriteLine($"PNG size: {fileInfo.Length} bytes");
                    }

                    return exportSucceeded && File.Exists(outputPngPath) ? 0 : 1;
                }
            }
            catch (Exception ex)
            {
                Console.Error.WriteLine(ex.ToString());
                return 1;
            }
        }
    }
}
