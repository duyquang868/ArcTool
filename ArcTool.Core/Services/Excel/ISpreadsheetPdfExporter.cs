using System;
using System.Collections.Generic;

namespace ArcTool.Core.Services.Excel
{
    public enum SpreadsheetEngine
    {
        MsExcel,
        Wps
    }

    public interface ISpreadsheetPdfExporter : IDisposable
    {
        SpreadsheetEngine Engine { get; }
        bool Open(string filePath);
        IReadOnlyList<string> GetSheetNames();
        IReadOnlyList<string> GetNamedRanges(string sheetName);
        bool ExportRegionToPdf(string sheetName, string regionName, string outputPdfPath);
    }
}
