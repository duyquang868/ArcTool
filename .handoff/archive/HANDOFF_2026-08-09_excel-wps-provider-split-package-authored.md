# ArcTool — HANDOFF TO NEXT SESSION
**Updated:** 2026-08-09  
**Status:** ACTIVE — Excel to Revit / WPS PDF-export provider split: research CLOSED, implementation NOT started; continue in a new chat

> Previous Quick Dimension phase-4 handoff is archived at `.handoff/archive/HANDOFF_2026-08-09_qd-t38-gate-carryforward.md` and its carry-forward state is still valid (see "Parallel open track" below).

---

## Goal and user request

Primary request for this phase (research only):
- confirm context on the **Excel to Revit** feature (marked CLOSED in `CLAUDE.md`)
- diagnose a newly discovered production defect: the feature works **only when Microsoft Excel is installed**, and does **not** work with **WPS Spreadsheet**
- research whether an API path exists so the feature can run on a machine with **WPS but no MS Excel**
- design the fix under an explicit architecture constraint from the user

Locked user directive (verbatim intent):
> "tôi muốn WPS COM/automation sẻ có một file logic riêng cho nó đừng đụng đến file logic của MS excel nhé, 2 thứ này chỉ cùng đi đến một điểm đó là xuất ra pdf thôi"

Meaning, as a hard constraint:
- WPS COM/automation lives in **its own separate logic file**
- it must **not modify** the MS Excel logic file
- the two branches converge **only at the PDF artifact**

---

## Current phase

Phase unit for this chat: **root-cause diagnosis + provider-split architecture design for the Excel→PDF step. No code.**

Completed in this phase:
- read `ExcelInteropService.cs` in full and confirmed the real dependency bottleneck
- corrected an early wrong diagnosis (see "Corrections" below)
- mapped the full blast radius (3 direct consumers)
- validated WPS COM automation feasibility from documentation/community evidence
- locked the target layering that satisfies the separation constraint

Not done in this phase: no source edit, no build, no Revit launch, no Revit MCP call, no smoke test, no re-index.

---

## Root cause (confirmed by reading current code)

The feature is **not** a data-import pipeline. It is a **render** pipeline:

```text
.xlsx
  -> MS Excel COM automation      <-- HARD DEPENDENCY
  -> ExportAsFixedFormat(PDF)     <-- THE ACTUAL BOTTLENECK
  -> PDFtoImage (PDFium, 300 DPI) <-- engine-agnostic, already fine
  -> PNG
  -> SkiaSharp white-margin crop  <-- engine-agnostic, already fine
  -> Revit ImageType / ImageInstance
```

Exact blocking call sites in [ExcelInteropService.cs](ArcTool.Core/Services/ExcelInteropService.cs):
- line 32 — `OpenFile`: `_excelApp = new Application { Visible = false, DisplayAlerts = false };` then `_excelApp.Workbooks.Open(filePath)`
- line 221 — `ExportRangeInternal`: `ws.ExportAsFixedFormat(XlFixedFormatType.xlTypePDF, tempPdf, XlFixedFormatQuality.xlQualityStandard, false, false, 1, 1, false)`
- file-level hard binding at lines 5–9: `using Microsoft.Office.Interop.Excel;` plus `Range` / `Application` aliases
- project-level hard binding: `<COMReference Include="Microsoft.Office.Interop.Excel">` in [ArcTool.Core.csproj:20](ArcTool.Core/ArcTool.Core.csproj:20)

Everything after the PDF is already independent of which spreadsheet app produced it. **Only the PDF-producing half needs a second provider.**

---

## Corrections carried forward (do not repeat these mistakes)

1. **ClosedXML / Open XML SDK / NPOI cannot fix this.** They read `.xlsx` without Excel but are **not rendering engines**. They cannot produce the PDF/PNG this feature needs. My first answer wrongly proposed them; the user corrected it and the correction was right.
2. **`CopyPicture` / clipboard / chart hacks are already removed** (V6.0). Do not reintroduce them; the dossier records why they failed (hidden Excel + virtual device context).
3. **Do not assume `Excel.Application` maps to WPS.** It does not reliably.

---

## Blast radius — every consumer that constructs the service directly

- [ExcelSyncEngine.cs:160](ArcTool.Core/Services/ExcelSyncEngine.cs:160) — `using (var svc = new ExcelInteropService())`, then `svc.OpenFile(...)` at 162 and `svc.ExportRegion(...)` at 171
- [ExcelToRevitWindow.xaml.cs:423](ArcTool.Core/UI/ExcelToRevitWindow.xaml.cs:423) — metadata read (`OpenFile` at 425)
- [ExcelToRevitWindow.xaml.cs:468](ArcTool.Core/UI/ExcelToRevitWindow.xaml.cs:468) — metadata read (`OpenFile` at 470)

Consequence: provider selection must reach the **UI metadata reads too** (`GetSheetNames`, `GetNamedRanges`), not only the export path. Touching 3+ source files ⇒ this qualifies for the multi-agent work package workflow under `CLAUDE.md`.

---

## Locked architecture (approved in discussion, not yet implemented)

New folder `ArcTool.Core/Services/Excel/`:

| File | Role | Constraint |
|---|---|---|
| `ISpreadsheetPdfExporter.cs` | abstraction; convergence point is the **PDF artifact**, never a workbook/COM object | no `Xl*` types in the signature |
| `MsExcelWorkbookPdfExporter.cs` | MS Excel branch | keeps `Microsoft.Office.Interop.Excel`; behavior unchanged |
| `WpsWorkbookPdfExporter.cs` | WPS branch | **late binding only**, no Interop reference, numeric constants instead of `XlPaperSize` / `XlFixedFormatType` / `XlFixedFormatQuality` |
| `PdfRasterImageService.cs` | shared raster half extracted from [ExcelInteropService.cs:129-355](ArcTool.Core/Services/ExcelInteropService.cs:129) (`EnsurePdfiumLoaded`, `EnsureSkiaSharpLoaded`, `Conversion.SavePng`, SkiaSharp crop, `WhiteThreshold = 240`) | engine-agnostic |
| `SpreadsheetImageExportService.cs` | coordinator: pick provider → PDF → raster → PNG | single entry point for all 3 call sites |

Revit-side import in `ExcelSyncEngine` stays **unchanged**.

### WPS provider specifics
- instantiate via `Type.GetTypeFromProgID(...)` + `Activator.CreateInstance(...)`, drive with `dynamic`/reflection
- ProgID fallback chain: **`KET.Application`** (modern/preferred) → **`ET.Application`** (legacy) → **`Kingsoft.ET.Application`**
- replicate the same 10 steps as the MS branch: create app → `Visible=false` → `DisplayAlerts=false` → open workbook → get worksheet by name → resolve region (**NamedRange → PrintArea → UsedRange**) → normalize page setup (PrintArea, `Zoom=false`, FitToPages 1×1, margins 0, prefer E-sheet then A3) → `ExportAsFixedFormat(PDF)` → release COM child → parent → quit
- WPS exposes `Workbooks.Open` and `Workbook.ExportAsFixedFormat`, which is why this is feasible

---

## Accepted trade-off (already stated to the user)

This swaps one app dependency for a choice of two. After the work, the feature needs **either** MS Excel **or** a WPS install with working COM automation. It still will not run on a machine with neither, because the rendering fidelity comes from the spreadsheet application itself.

---

## Open validation risks (unresolved)

1. **WPS COM registration is broken in some builds** — reported for `12.1.0.22529` / `12.1.0.22525`. Provider must fail cleanly with a user-facing message, not crash.
2. **PDF fidelity differs from Excel** — page setup, print area, font substitution, merged cells, scaling, margins, line weight. Needs a side-by-side comparison on a real fixture.
3. **Named-range scope semantics** — `RefersToRange`, workbook- vs worksheet-scope, protected sheets. Current MS code filters `r?.Worksheet?.Name == sheetName`; WPS parity is unverified.
4. **Typed enum removal** — every `Xl*` enum in the WPS branch must become a numeric constant.
5. Any Revit-side verification requires **explicit user instruction** (operator-controlled runtime).

---

## Done / unfinished / blocked

Done:
- root cause confirmed at the exact call sites
- blast radius mapped
- WPS feasibility validated from documentation/community evidence
- architecture locked and consistent with the user's separation constraint

Unfinished (next phase):
- build the work package per `.claude/skills/arctool-work-package/` from `.claude/workpackages/_TEMPLATE/`
- implement the 5 files above
- rewire the 3 call sites
- update `.Dossier/Detailed Technical Dossier - Excel to Revit.md` sections 9/10/13 once implemented

Blocked:
- fidelity and ProgID-availability validation need a machine with WPS installed plus explicit operator runtime approval

---

## Verification run

Verified by reading source: `ExcelInteropService.cs` (585 lines, full), `ArcTool.Core.csproj` (full), `ExcelSyncEngine.cs` (first ~200 lines), `ExcelToRevitWindow.xaml.cs` (first ~220 lines), `ExcelMapping.cs` (excerpt).

Not run: no build, no tests, no Revit runtime, no Revit MCP, no re-index. Reason: research/design phase only, and the user has not authorized a runtime action.

Tool notes: `get_architecture` requires project id `D-Quang mini-OneDrive - MSFT-Plugin Revit-ArcTool` (2421 nodes / 3834 edges). Two WebFetch attempts on WPS docs failed (403); findings came from WebSearch summaries instead. Broad greps must exclude `.claude/worktrees/agent-*` copies or they blow the 250-file limit.

---

## Next-session starting point

Start a **NEW chat** with the prompt in "New-chat prompt" below.

Immediate carry-forward context:
- the bottleneck is `ExportAsFixedFormat`, not file parsing — do not re-litigate ClosedXML
- MS Excel logic file must not be modified in behavior; WPS gets its own file
- convergence point is the PDF artifact
- start with plan + work package, not code

---

## Parallel open track (unchanged, different phase)

Quick Dimension phase-4 hardening is still open and independent:
- `R8_C07` is closed as a concrete negative mid-run verdict; `T3.7` is the durable source
- `C01`/`C02`/`C03` **were tested** by the user — the gap is publication, not evidence
- next QD action is to publish `T3.5_result.md`, then `T3.6_result.md`, then re-evaluate `T3.8`
- full detail: `.handoff/archive/HANDOFF_2026-08-09_qd-t38-gate-carryforward.md`

Do not mix this track with the Excel/WPS phase in the same chat.

---

## Invariants to preserve

1. One chat = one phase; this research/design phase is closed and implementation starts in a new chat.
2. Revit runtime is operator-controlled: no Revit launch, `.rvt` open, MCP call, or smoke test without an explicit request.
3. The MS Excel exporter file and the WPS exporter file stay separate; they meet only at the PDF.
4. The PDF → PDFtoImage → PNG → SkiaSharp crop → Revit import half is shared and must not be duplicated per provider.
5. Excel to Revit dossier invariants remain in force: COM release child → parent, `StoredWidth`/`StoredHeight` in millimetres, two-transaction image create/resize, local-time drift checks, `_suppressRowEvents` discipline.
6. Revit API docs (https://www.revitapidocs.com/2026/) must be checked before the code-change phase.
7. Worker/subagent dispatch model in ArcTool work packages: `"sonnet"`.

---

## New-chat prompt

```
ArcTool — Excel to Revit: tách provider xuất PDF.

Vấn đề: feature phụ thuộc MS Excel COM (ArcTool.Core/Services/ExcelInteropService.cs),
máy chỉ có WPS thì không chạy. Điểm nghẽn: new Application() (line 32) và
ws.ExportAsFixedFormat(xlTypePDF, ...) (line 221).

Yêu cầu: WPS COM/automation nằm ở FILE LOGIC RIÊNG, không đụng file logic MS Excel.
Hai nhánh chỉ gặp nhau ở điểm xuất ra PDF; pipeline PDF → PDFtoImage → PNG → SkiaSharp crop
→ Revit ImageType/ImageInstance dùng chung.

Kiến trúc đã chốt:
- ISpreadsheetPdfExporter
- ArcTool.Core/Services/Excel/MsExcelWorkbookPdfExporter.cs (giữ Interop)
- ArcTool.Core/Services/Excel/WpsWorkbookPdfExporter.cs (late binding, KHÔNG reference Interop,
  ProgID fallback: KET.Application → ET.Application → Kingsoft.ET.Application, dùng hằng số,
  không dùng enum Xl*)
- ArcTool.Core/Services/Excel/PdfRasterImageService.cs (PDFtoImage + SkiaSharp, tách từ
  ExcelInteropService.cs:129-355)
- ArcTool.Core/Services/Excel/SpreadsheetImageExportService.cs (chọn provider)

Call sites cần rewire: ExcelSyncEngine.cs:160, ExcelToRevitWindow.xaml.cs:423 và :468.

Đọc trước: .handoff/HANDOFF_TO_NEXT_SESSION.md và
Memory/project_excel_to_revit_wps_provider_split.md

Vì chạm 3+ source file → lập work package theo .claude/skills/arctool-work-package/,
worker dispatch model: "sonnet". Không chạy Revit/MCP/smoke test nếu tôi không yêu cầu rõ.

Bắt đầu bằng plan + work package, chưa sửa code.
```

---

## Reference files

- Archived QD handoff (parallel track): `.handoff/archive/HANDOFF_2026-08-09_qd-t38-gate-carryforward.md`
- Feature dossier: `.Dossier/Detailed Technical Dossier - Excel to Revit.md`
- Durable memory record for this phase: `Memory/project_excel_to_revit_wps_provider_split.md`
- Work package skill/scaffold: `.claude/skills/arctool-work-package/`, `.claude/workpackages/_TEMPLATE/`
- Root operating document: `CLAUDE.md`
