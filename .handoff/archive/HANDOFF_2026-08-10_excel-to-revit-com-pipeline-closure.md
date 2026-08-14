# ArcTool — HANDOFF ARCHIVE
**Archived:** 2026-08-10
**Phase:** Excel to Revit — restore direct COM export pipeline with MS Excel primary and WPS fallback
**Status:** CLOSED

---

## What this phase delivered

Implemented and verified the new live Excel-to-Revit export baseline:

- restored direct COM export on the **source workbook**
- removed the **ClosedXML temp-workbook** branch from the live pipeline
- set engine preference to **MS Excel first**, **WPS fallback second**
- kept the shared **PDF -> PNG -> crop** stage unchanged
- kept the **PNG -> Revit import/update** flow in `ExcelSyncEngine` unchanged

### Source changes completed
- added `ArcTool.Core/Services/Excel/ISpreadsheetPdfExporter.cs`
- added `ArcTool.Core/Services/Excel/MsExcelWorkbookPdfExporter.cs`
- added `ArcTool.Core/Services/Excel/WpsWorkbookPdfExporter.cs`
- refactored `ArcTool.Core/Services/Excel/SpreadsheetImageExportService.cs`
- removed `ArcTool.Core/Services/Excel/Pdf24WorkbookPdfConverter.cs` from the live codebase

### Runtime architecture after closure
```text
source workbook
  -> SpreadsheetImageExportService
     -> MsExcelWorkbookPdfExporter (preferred)
     -> WpsWorkbookPdfExporter (fallback)
     -> temp PDF
     -> PdfRasterImageService
     -> PNG
  -> ExcelSyncEngine imports PNG into Revit (unchanged)
```

---

## Verification completed

Build gate passed:
- `ArcTool.Core -> ...\ArcTool.Core.dll`

Observed warning remained unrelated to this phase:
- `ArcTool.Core/Services/QuickDimensionReadOnlyXmlLogService.cs(77,32)` `CS8600`

Static verification completed:
- `SpreadsheetImageExportService` no longer creates a temp workbook
- live path no longer references `Pdf24WorkbookPdfConverter`
- `WorkbookRegionSnapshotService` still exists in source history but is no longer in the live export path
- region resolution remains `NamedRange -> PrintArea -> UsedRange`
- `ExcelSyncEngine` still owns only the PNG import/update side

---

## Durable records updated in this phase

- `Memory/project_excel_to_revit_wps_provider_split.md`
- `Memory/MEMORY.md`
- `.Dossier/Detailed Technical Dossier - Excel to Revit.md`

No ADR write.
No Revit runtime smoke.
No re-index.

---

## Next-session starting point

This phase is closed. The next Excel-to-Revit phase, if reopened, should start from the new baseline above rather than from the retired ClosedXML/PDF24 experiment.

Use the live seam:
- `ArcTool.Core/Services/Excel/SpreadsheetImageExportService.cs`

Primary engine file:
- `ArcTool.Core/Services/Excel/MsExcelWorkbookPdfExporter.cs`

Fallback engine file:
- `ArcTool.Core/Services/Excel/WpsWorkbookPdfExporter.cs`

Shared raster stage:
- `ArcTool.Core/Services/Excel/PdfRasterImageService.cs`

Revit-side import remains in:
- `ArcTool.Core/Services/ExcelSyncEngine.cs`
