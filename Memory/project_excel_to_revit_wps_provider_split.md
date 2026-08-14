---
name: project_excel_to_revit_wps_provider_split
description: Excel to Revit only works with MS Excel installed; root cause is the PDF export call, and the locked fix is a separate WPS provider file converging at the PDF artifact.
metadata:
  type: project
---

Discovered 2026-08-08: the CLOSED **Excel to Revit** feature only works when Microsoft Excel is installed. It fails on machines that have WPS Spreadsheet but no MS Excel.

Updated 2026-08-10: the live pipeline now uses `SpreadsheetImageExportService` under `ArcTool.Core/Services/Excel/` with **MS Excel first and WPS fallback second**. The temporary ClosedXML workbook-shaping branch was accepted as an intermediate experiment only and is no longer in the live path.

**Root cause** (confirmed by reading `ArcTool.Core/Services/ExcelInteropService.cs` in full): the feature is a *render* pipeline, not a data-import one. The blocking calls are `new Application()` in `OpenFile` (line 32) and `ws.ExportAsFixedFormat(XlFixedFormatType.xlTypePDF, ...)` in `ExportRangeInternal` (line 221). Everything after the PDF (`PDFtoImage`/PDFium 300 DPI → PNG → SkiaSharp white-margin crop, threshold 240) is already engine-agnostic. ClosedXML / Open XML SDK / NPOI cannot fix this — they parse `.xlsx`, they do not render.

**Blast radius** — 3 direct consumers construct `ExcelInteropService` themselves, so implementation is work-package scale per `CLAUDE.md`:
- `ArcTool.Core/Services/ExcelSyncEngine.cs:160` (`OpenFile` 162, `ExportRegion` 171)
- `ArcTool.Core/UI/ExcelToRevitWindow.xaml.cs:423` (`OpenFile` 425)
- `ArcTool.Core/UI/ExcelToRevitWindow.xaml.cs:468` (`OpenFile` 470)

**User's locked architecture constraint** (verbatim intent): WPS COM/automation must live in its **own separate logic file**, must **not touch** the MS Excel logic file; the two branches converge **only at the point of producing PDF**.

**Why:** keeping the branches physically separate prevents a WPS regression from destabilizing the production MS Excel path, and keeps the Interop dependency confined to one file.

**How to apply:** implement under `ArcTool.Core/Services/Excel/`:
- `ISpreadsheetPdfExporter` — abstraction whose convergence point is the **PDF file path**, never a workbook or COM object
- `MsExcelWorkbookPdfExporter.cs` — MS Excel branch, keeps `Microsoft.Office.Interop.Excel`, direct workbook open + direct region export, primary engine
- `WpsWorkbookPdfExporter.cs` — late-bound only (`Type.GetTypeFromProgID` + `Activator.CreateInstance`), **no** Interop reference, numeric constants instead of `XlPaperSize` / `XlFixedFormatType` / `XlFixedFormatQuality`; ProgID fallback `KET.Application` → `ET.Application` → `Kingsoft.ET.Application`; normal visible WPS UI during export is accepted product behavior
- `PdfRasterImageService.cs` — shared raster half extracted from `ExcelInteropService.cs:129-355` (`EnsurePdfiumLoaded`, `EnsureSkiaSharpLoaded`, `Conversion.SavePng`, Skia crop)
- `SpreadsheetImageExportService.cs` — coordinator that opens MS Excel first, falls back to WPS only if needed, and owns the temp PDF lifetime only

Live invariant after 2026-08-10: region resolution remains `NamedRange -> PrintArea -> UsedRange`, but it now runs directly on the source workbook through the active COM exporter. Revit-side import in `ExcelSyncEngine` (`ImageTypeOptions` → `ImageType.Create` → `ImageInstance.Create`, Smart Scale, two-transaction flow) stays unchanged.

**Accepted trade-off** stated to the user: this swaps one hard app dependency for a choice of two — after the work the feature needs *either* MS Excel *or* a WPS install with working COM automation; it still cannot run with neither.

**Open validation risks:** WPS COM registration reportedly broken in some builds (12.1.0.22529 / 12.1.0.22525); PDF fidelity vs Excel unverified (page setup, print area, font substitution, merged cells, scaling, margins, line weight); named-range scope parity (`RefersToRange`, workbook- vs worksheet-scope, protected sheets) unverified. Any Revit-side smoke test requires an explicit user request.

Related: [[feedback_multi_agent_work_package_workflow]], [[feedback_revit_runtime_operator_control_and_journal_analysis]]. Deep record: `.Dossier/Detailed Technical Dossier - Excel to Revit.md` section 19.
