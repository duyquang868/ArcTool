# EXCEL TO REVIT — WPS PROVIDER SPLIT — TASK MANIFEST

Execution order, dependency graph, and exclusive write scopes.

Worker model: `model: "sonnet"` on every dispatch. No exceptions recorded for this package.

---

## Phase 1 — Preflight and API lock

| Task | Purpose | Depends on | Exclusive write scope |
|---|---|---|---|
| `T1.1` | Lock the `ISpreadsheetPdfExporter` session shape and the exact MS-Excel COM member list the split must preserve; cite Microsoft Learn for each member | — | result only |
| `T1.2` | Lock the WPS late-binding strategy: ProgID order, `InvokeMember` shapes, numeric constants for `xlTypePDF` / `xlQualityStandard` / paper sizes, failure semantics | `T1.1` | result only |
| `T1.3` | Preflight GO / NO-GO: is the split implementable with no Interop leak into the WPS file and no behavior change on MS Excel? | `T1.2` | result only |

## Phase 2 — Shared stage and abstraction

| Task | Purpose | Depends on | Exclusive write scope |
|---|---|---|---|
| `T2.1` | Author `ISpreadsheetPdfExporter.cs` (interface + `SpreadsheetEngine` enum + XML docs) | `T1.3` | `ArcTool.Core/Services/Excel/ISpreadsheetPdfExporter.cs` |
| `T2.2` | Author `PdfRasterImageService.cs` — port lines 100-177 and 231-348 verbatim in behavior, threshold 240, 300 DPI, native-loader fallbacks | `T1.3` | `ArcTool.Core/Services/Excel/PdfRasterImageService.cs` |
| `T2.3` | Build gate after the two engine-neutral files land | `T2.1`, `T2.2` | result only |

## Phase 3 — Providers, coordinator, rewire

| Task | Purpose | Depends on | Exclusive write scope |
|---|---|---|---|
| `T3.1` | Author `MsExcelWorkbookPdfExporter.cs` — port `OpenFile`, `GetSheetNames`, `GetNamedRanges`, region resolution, PageSetup + `ExportAsFixedFormat`, `Dispose`; drop the two dead methods | `T2.3` | `ArcTool.Core/Services/Excel/MsExcelWorkbookPdfExporter.cs` |
| `T3.2` | Author `WpsWorkbookPdfExporter.cs` — late binding only, no Interop, ProgID fallback chain, numeric constants | `T2.3` | `ArcTool.Core/Services/Excel/WpsWorkbookPdfExporter.cs` |
| `T3.3` | Author `SpreadsheetImageExportService.cs` — auto-detect (MS first), session factory, PDF→PNG handoff, temp-PDF ownership, diagnostic message when neither engine is present | `T3.1`, `T3.2` | `ArcTool.Core/Services/Excel/SpreadsheetImageExportService.cs` |
| `T3.4` | Backup `ExcelInteropService.cs` to `Services/_backup/ExcelInteropService.cs.bak`, then delete the original | `T3.3` | `ArcTool.Core/Services/_backup/ExcelInteropService.cs.bak`, `ArcTool.Core/Services/ExcelInteropService.cs` |
| `T3.5` | Rewire `ExcelSyncEngine.cs:160` to the coordinator; touch nothing else in that file | `T3.4` | `ArcTool.Core/Services/ExcelSyncEngine.cs` |
| `T3.6` | Rewire `ExcelToRevitWindow.xaml.cs:423` and `:468`; preserve `_suppressRowEvents` discipline (BUG-P3-01) | `T3.4` | `ArcTool.Core/UI/ExcelToRevitWindow.xaml.cs` |
| `T3.7` | Build gate + static isolation audit: grep-prove no Interop in the WPS file, no WPS in the MS file, zero `ExcelInteropService` references | `T3.5`, `T3.6` | result only |

`T3.1` and `T3.2` may run concurrently — different files, no shared write scope.
`T3.5` and `T3.6` may run concurrently — different files.

## Phase 4 — Parity review and operator runbooks

| Task | Purpose | Depends on | Exclusive write scope |
|---|---|---|---|
| `T4.1` | MS Excel parity review against contract R11, item by item, `file:line` old vs new | `T3.7` | result only |
| `T4.2` | Write the operator runbooks for `EV-1` (WPS ProgID/member probe), `EV-2` (WPS end-to-end + fidelity), `EV-3` (MS Excel non-regression) | `T3.7` | result only |

## Phase 5 — Runtime evidence and verdict

| Task | Purpose | Depends on | Exclusive write scope |
|---|---|---|---|
| `T5.1` | Analyze `EV-1` + `EV-2` WPS evidence; decide whether the WPS branch needs a follow-up patch | `T4.2` + `EV-1` + `EV-2` | result only |
| `T5.2` | Analyze `EV-3` MS Excel non-regression evidence | `T4.2` + `EV-3` | result only |
| `T5.3` | Final verdict GO / NO-GO for the package | `T5.1`, `T5.2` | result only |

If `T5.1` finds a WPS defect, the master opens `T5.1b` with exclusive write scope
`ArcTool.Core/Services/Excel/WpsWorkbookPdfExporter.cs` and re-runs `T3.7`.

## Phase 6 — Durable closure

| Task | Purpose | Depends on | Exclusive write scope |
|---|---|---|---|
| `T6.1` | Persist durable knowledge: update `Memory/project_excel_to_revit_wps_provider_split.md`, dossier §13.0 → closed record, `CLAUDE.md` pointers, locked-decisions register | `T5.3` | `Memory/`, `.Dossier/`, `CLAUDE.md`, `.Dossier/ArcTool Locked Technical Decisions.md` |
| `T6.2` | Draft the final master closure message | `T6.1` | result only |

---

## Source-file lock summary

- `ArcTool.Core/Services/Excel/ISpreadsheetPdfExporter.cs` — `T2.1` only.
- `ArcTool.Core/Services/Excel/PdfRasterImageService.cs` — `T2.2` only.
- `ArcTool.Core/Services/Excel/MsExcelWorkbookPdfExporter.cs` — `T3.1` only.
- `ArcTool.Core/Services/Excel/WpsWorkbookPdfExporter.cs` — `T3.2`, then `T5.1b` if opened. In that order.
- `ArcTool.Core/Services/Excel/SpreadsheetImageExportService.cs` — `T3.3` only.
- `ArcTool.Core/Services/ExcelInteropService.cs` (+ `_backup/`) — `T3.4` only.
- `ArcTool.Core/Services/ExcelSyncEngine.cs` — `T3.5` only.
- `ArcTool.Core/UI/ExcelToRevitWindow.xaml.cs` — `T3.6` only.
- `ArcTool.Core/ArcTool.Core.csproj` — **no task writes this file.** New `.cs` files are picked up by
  the SDK glob; the `COMReference` stays as-is. A worker that believes the csproj must change returns
  `BLOCKED` instead of editing it.
- Durable memory / dossier / `CLAUDE.md` — `T6.1` only.

Build tasks (`T2.3`, `T3.7`) run only after their preceding write tasks finish.

---

## Phase gates

- Stop the whole package if `T1.3` returns `NO_GO`.
- Do not start Phase 3 until `T2.3` is `PASS`.
- Do not start `T3.4` until `T3.3` is `PASS` — deleting the legacy file before the coordinator exists
  would break the build mid-package.
- Do not start Phase 4 until `T3.7` is `PASS`.
- Do not start Phase 6 until `T5.3` is `PASS`.

---

## Result-file convention

Every task writes exactly one detailed result file:

- `.claude/workpackages/excel-to-revit-wps-provider-split/results/<TASK_ID>_result.md`

`result only` tasks are dispatched **without** `isolation: "worktree"`. Source-writing tasks in this
package are serialized by the lock summary above, so they also do not need worktree isolation.

The master consumes only the compact envelope unless it must resolve a contradiction.
