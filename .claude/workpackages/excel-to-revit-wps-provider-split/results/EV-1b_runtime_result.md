# EV-1b — Runtime re-check after T5.1b patch

- status: BLOCKED
- date: 2026-08-09
- inputs_read: `.claude/workpackages/excel-to-revit-wps-provider-split/evidence/EV-1b_probe.ps1`; `.claude/workpackages/excel-to-revit-wps-provider-split/evidence/EV-1b_output.txt`; `ArcTool.Core/Services/Excel/WpsWorkbookPdfExporter.cs`; `.claude/workpackages/excel-to-revit-wps-provider-split/results/T5.1b_result.md`; `.claude/workpackages/excel-to-revit-wps-provider-split/results/T3.7_rerun_after_T5.1b_result.md`
- write_scope_touched: `.claude/workpackages/excel-to-revit-wps-provider-split/results/EV-1b_runtime_result.md`

## Findings
- EV-1b was run locally without Revit against the same sample workbook `C:\Users\ADMIN\Desktop\PA4\BULONG.xlsx` using a scratch copy. WPS Spreadsheet still resolves through `KET.Application` on this machine (`12.1.0.28032`), and application activation still works.
- The post-patch `Workbooks.Open` ladder introduced by `T5.1b` did **not** fix the runtime defect on this WPS build. All four patched call shapes failed with the same bind-time COM error `DISP_E_TYPEMISMATCH` (`0x80020005`):
  1. `Open(filePath)`
  2. `Open(filePath, Type.Missing)`
  3. `Open(filePath, Type.Missing, Type.Missing)`
  4. `Open(filePath, Type.Missing, Type.Missing, Type.Missing, Type.Missing)`
- Diagnostic shapes outside the current patch also failed identically:
  - full 15-arg trailing `Type.Missing`
  - trailing `$null` variants
  - `Missing.Value`
  - explicit `UpdateLinks=0`
  - explicit `UpdateLinks=0, ReadOnly=false`
- Because workbook open still never succeeded, every downstream WPS member remains unverified in runtime: `Worksheets`, `Names`, `UsedRange`, `PageSetup`, `Range.Address`, `PaperSize`, and `ExportAsFixedFormat` were all skipped.
- This isolates the remaining defect more tightly than EV-1/T5.1b: the failure is not merely “one positional arg vs a few explicit optional placeholders.” On this WPS automation surface, the current `InvokeMember(..., "Open", ...)` binding strategy itself is still incompatible with the discovered `KET.Application` `Workbooks` COM object.
- Non-open members stayed healthy in EV-1b exactly as before: `CreateInstance`, `Visible(set)`, `DisplayAlerts(set)`, `Application.Version(get)=12.0`, `Workbooks`, and `Application.Quit` all succeeded.

## Decision
`T5.1b` is build-clean but runtime-insufficient. EV-1 remains open and now requires a second follow-up patch that changes the WPS `Workbooks.Open` invocation strategy again. Package status is therefore **BLOCKED on a new WPS-only fix task**, not PASS.

## Open questions
- Which late-bound invocation shape is actually accepted by this `KET.Application` `Workbooks.Open` dispatch surface is still unknown.
- It is still unknown whether downstream members (`Worksheets`, `Names`, `PageSetup`, `ExportAsFixedFormat`) are compatible once workbook open is solved.

## Handoff for downstream tasks
- Open a new WPS-only fix task with exclusive `write_scope` `ArcTool.Core/Services/Excel/WpsWorkbookPdfExporter.cs`.
- The next patch should target the invocation mechanism itself, not just add more trailing argument counts to the existing `InvokeMember` call.
- Authoritative evidence path: `.claude/workpackages/excel-to-revit-wps-provider-split/evidence/EV-1b_output.txt`.
