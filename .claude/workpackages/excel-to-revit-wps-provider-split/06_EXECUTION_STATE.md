# EXCEL TO REVIT — WPS PROVIDER SPLIT — EXECUTION STATE

The master owns this file. Workers never edit it.

Update rules:
- Status is `PENDING`, `RUNNING`, `PASS`, `BLOCKED`, or `NO_GO`.
- Record the result file after every worker finishes.
- Keep notes short and factual. Do not delete history; append or update in place.

---

## Package facts locked at creation (2026-08-09)

- Provider precedence: MS Excel first, WPS fallback, auto-detect, no UI picker.
- Legacy `ExcelInteropService.cs` is backed up to `ArcTool.Core/Services/_backup/ExcelInteropService.cs.bak`
  then deleted (user decision).
- Dev machine has MS Excel **and** WPS Office 12.1.0.28032 (installed 2026-08-09; `KET.Application`
  resolves, per-user/HKCU only). `EV-1`/`EV-2`/`EV-3` all run on this one machine now. Superseded the
  earlier "no WPS" fact — see history below.
- `GetActiveSheetName` and `ExportPrintAreaAsHighResImage` are dead code and are not ported.

---

## Current state

| Task | Status | Owner | Evidence | Result file | Notes |
|---|---|---|---|---|---|
| `T1.1` | PASS | worker (sonnet) | — | `results/T1.1_result.md` | session shape + MS COM member lock; all 4 `Xl*` values sourced |
| `T1.2` | PASS | worker (sonnet) | — | `results/T1.2_result.md` | WPS late-binding strategy; 27-row call-shape table, all 4 constants sourced |
| `T1.3` | PASS | worker (sonnet) | — | `results/T1.3_result.md` | gate verdict **GO**; Phase 2 cleared |
| `T2.1` | PASS | worker (sonnet) | — | `results/T2.1_result.md` | wrote `ISpreadsheetPdfExporter.cs`; no deviation from `T1.1` lock |
| `T2.2` | PASS | worker (sonnet) | — | `results/T2.2_result.md` | wrote `PdfRasterImageService.cs`; `IsNonWhitePixel` parity gap resolved on re-dispatch, no code change needed |
| `T2.3` | PASS | worker (sonnet) | — | `results/T2.3_result.md` | build gate: exit 0, no error, no warning naming either new file; Phase 3 cleared |
| `T3.1` | PASS | worker (sonnet) | — | `results/T3.1_result.md` | wrote MS provider; repair pass 2026-08-09 actually landed the CS0104 qualification and corrected the earlier false PASS |
| `T3.2` | PASS | worker (sonnet) | — | `results/T3.2_result.md` | wrote WPS provider; repair pass renamed the four bare `Xl*` constants to provider-neutral names, numeric values unchanged |
| `T3.3` | PASS | worker (sonnet) | — | `results/T3.3_result.md` | wrote coordinator; MS-first/WPS-fallback, temp-PDF lifecycle moved here |
| `T3.4` | PASS | worker (sonnet) | — | `results/T3.4_result.md` | backup created, legacy deleted safely |
| `T3.5` | PASS | worker (sonnet) | — | `results/T3.5_result.md` | rewired `ExcelSyncEngine.cs` surgically; behavior/messages preserved |
| `T3.6` | PASS | worker (sonnet) | — | `results/T3.6_result.md` | rewired both UI construction sites; `_suppressRowEvents` discipline preserved |
| `T3.7` | PASS | worker (sonnet) | — | `results/T3.7_result.md` | authoritative rerun exit 0, zero warnings; all 9 static isolation audit checks passed; Phase 4 cleared |
| `T4.1` | PASS | worker (sonnet) | — | `results/T4.1_result.md` | parity review PASS; all 16 R11 items SAME/RELOCATED, no concerns |
| `T4.2` | PASS | worker (sonnet) | — | `results/T4.2_result.md` | runbooks written; `04_EVIDENCE_QUEUE.md` updated, structure preserved |
| `T5.1` | BLOCKED | master (user-directed runtime exception) | `EV-1` + `EV-1b` complete locally; `EV-2` pending | `results/T5.1_result.md` | EV-1b runtime re-check shows the `T5.1b` patch does NOT fix the WPS `Workbooks.Open` defect; every patched shape (and every wider diagnostic shape) still fails with `DISP_E_TYPEMISMATCH`; needs a new follow-up fix task before T5.1 can proceed to EV-2 |
| `T5.1b` | PASS (build) / INSUFFICIENT (runtime) | worker (sonnet) | `EV-1` | `results/T5.1b_result.md` | patched `WpsWorkbookPdfExporter.cs` open path with compact ordered late-bound `Open` fallbacks; build/static-isolation PASSed but `EV-1b` runtime re-check shows the patch did not resolve the defect |
| `EV-1b` | BLOCKED | master (user-directed runtime exception) | `evidence/EV-1b_output.txt` | `results/EV-1b_runtime_result.md` | runtime re-check of the `T5.1b` patch; all 4 patched `Open` shapes + 6 wider diagnostic shapes failed identically with `DISP_E_TYPEMISMATCH`; workbook never opened, all downstream WPS members remain unverified |
| `EV-1c` | PASS (diagnostic) | master (user-directed runtime exception) | `evidence/EV-1c_output.txt` | — | binder-comparison probe; first probe to null-check the result and reveal that `Application.Workbooks` returns null, reframing the defect as upstream of `Open` |
| `EV-1d` | PASS (diagnostic) | master (user-directed runtime exception) | `evidence/EV-1d_output.txt` | `results/EV-1d_root_cause_result.md` | root cause isolated: `KET.Application` activates and answers `Version`/`Visible`/`DisplayAlerts`/`Quit`, but `Application.Workbooks` is null across CreateInstance, GetActiveObject, 3 binders, and a 10x500ms readiness loop; `LocalServer32`/`InprocServer32`/`TypeLib` registry values read empty |
| `T5.2` | PENDING | — | `EV-3` | `results/T5.2_result.md` | MS non-regression analysis |
| `T5.3` | PENDING | — | — | `results/T5.3_result.md` | final verdict |
| `T6.1` | PENDING | — | — | `results/T6.1_result.md` | durable persistence |
| `T6.2` | PENDING | — | — | `results/T6.2_result.md` | closure message |

`T5.1b` was created because `T5.1` confirmed a WPS defect requiring a patch.

---

## History

- 2026-08-09 — package created. Plan approved shape: 6 phases, 20 tasks. No source file modified yet.
- 2026-08-09 — all 20 task files authored under `tasks/`. Scaffold complete and internally consistent
  with `03_TASK_MANIFEST.md`. Verified by `git status`: no file under `ArcTool.Core/Services/`,
  `ArcTool.Core/UI/`, or `ArcTool.Core.csproj` modified; `Services/Excel/` not yet created. Awaiting
  user approval before Phase 1 dispatch.
- 2026-08-09 — user installed WPS Office on this machine after package creation. Re-probed ProgIDs:
  `KET.Application` resolves (CLSID `45540001-5750-5300-4b49-4e47534f4655`, per-user/HKCU only,
  launches via `wps.exe /prometheus /et /Automation`); `ET.Application`, `Kingsoft.ET.Application`,
  `WPS.Application`, `ET.Sheet` still null; `KWPS.Application` resolves but is Writer, not the
  spreadsheet app. The earlier "no WPS on dev machine" fact recorded at package creation is void.
  Corrected in `01_SHARED_CONTRACT.md` §5/§6, `02_MASTER_ORCHESTRATOR.md` evidence routing,
  `04_EVIDENCE_QUEUE.md` environment constraint, and this file's facts block. `EV-1`/`EV-2`/`EV-3` now
  all run on one machine; still operator-run, not worker-run (R1 — launching WPS COM is runtime
  action reserved to the operator).
- 2026-08-09 — `T1.1` PASS. Session shape locked: `ISpreadsheetPdfExporter : IDisposable` with
  `Engine{get}`, `Open(path)->bool` (swallow, no throw), `GetSheetNames`/`GetNamedRanges`
  swallow-and-return-empty, `ExportRegionToPdf(sheet,region,outputPdfPath)->bool`; region resolution
  stays inside provider; coordinator owns temp-PDF deletion post-split. All 4 `Xl*` enum values sourced
  (none `UNSOURCED`): `xlTypePDF=0`, `xlQualityStandard=0`, `xlPaperEsheet=26`, `xlPaperA3=8`. COM
  release order confirmed child→parent, workbook→app. No source file touched.
- 2026-08-09 — `T1.2` PASS. ProgID chain locked `KET.Application` → `ET.Application` →
  `Kingsoft.ET.Application`; `KWPS.Application` explicitly denylisted (Writer, not spreadsheet). 27-row
  `InvokeMember` call-shape table produced (target/member/`BindingFlags`/args/fatal-or-tolerable per
  row). All 4 constants sourced on the Excel side (`xlTypePDF=0`, `xlQualityStandard=0`,
  `xlPaperEsheet=26`, `xlPaperA3=8`); WPS-side numeric equivalence flagged `ASSUMPTION` for `EV-1`, none
  `UNSOURCED`. Failure semantics mapped to `EngineAbsent` / `EngineFoundOpenFailed` /
  `EngineFoundExportFailed`, plus a new post-export `File.Exists` check (justified by the shell-mode
  server risk). Two open questions carried to `EV-1`: whether `Range.Address[false,false]` needs
  `GetProperty` alone or combined with `InvokeMethod` on WPS `IDispatch`; whether `ExportAsFixedFormat`'s
  8-arg signature is accepted as-is by `KET.Application`'s spreadsheet mode. No WPS COM instantiated by
  the worker (R1 respected). No source file touched.
- 2026-08-09 — `T1.3` PASS, verdict **GO**. Interface confirmed engine-neutral by inspection (no
  Interop leak possible into `WpsWorkbookPdfExporter.cs`); `ArcTool.Core.csproj` needs no edit, SDK glob
  covers `Services/Excel/`, `COMReference` untouched. R11 preserved with one contract-mandated
  relocation (temp-PDF cleanup provider→coordinator), not treated as a defect. Coordinator can
  distinguish `EngineAbsent` from `EngineFound*Failed` using construction-time vs call-time signals,
  no interface change needed. Remaining WPS behavioral claims stay `ASSUMPTION`, routed to `EV-1`/`EV-2`,
  explicitly non-blocking for Phase 2-4. **Phase 1 complete — Phase 2 (`T2.1`, `T2.2`) is cleared to
  start on user go-ahead.** No source file touched in Phase 1.
- 2026-08-09 — `T2.1` PASS. Wrote `ArcTool.Core/Services/Excel/ISpreadsheetPdfExporter.cs`: enum +
  interface only, matches `T1.1` signature lock verbatim. Namespace `ArcTool.Core.Services.Excel`
  (matches sibling convention, verified against `ExcelInteropService.cs`/`ExcelSyncEngine.cs`).
  Isolation grep (`Interop`, `Xl`, `KET`, `ET.Application`, `Worksheet`, `Workbook`, `Range`) returned
  3 hits, all non-meaningful (prose/identifier substrings), no COM type reference — neutrality gate
  passes. XML docs cover failure-return semantics, temp-PDF ownership (caller owns; method never
  deletes the output PDF), and region-resolution order (NamedRange → PrintArea → UsedRange). No
  deviation from `T1.1`. First source file of the package. Phase 2 continues with `T2.2`.
- 2026-08-09 — `T2.2` PASS. Wrote `ArcTool.Core/Services/Excel/PdfRasterImageService.cs` as a static
  class with a single entry point `bool RenderPdfToCroppedPng(string pdfPath, string outputPngPath)`.
  Zero COM / Interop / spreadsheet knowledge. Ported `GetRuntimeFolder`,
  `GetNativeLibraryCandidates`, `EnsurePdfiumLoaded`, `EnsureSkiaSharpLoaded`, the 300-DPI
  `Conversion.SavePng` render, and the four-direction white-margin crop; `WhiteThreshold = 240`,
  DPI 300, native candidate-path order (assembly dir → `native/` → `runtimes/win-x64/native/`) and
  load-once guards all unchanged. Temp-PDF deletion deliberately NOT ported (moves to the `T3.3`
  coordinator per contract) and that relocation is documented in the file's doc comments.
  First envelope carried one parity unknown: `IsNonWhitePixel` was reconstructed from the task
  description because its definition sits at `ExcelInteropService.cs:193`, outside the authorized
  read range (100-177, 231-359). Master re-dispatched the same worker with 188-200 authorized rather
  than reading/patching from the master chat (write-lock policy). Follow-up PASS: source predicate is
  an expression-bodied local function at 193-197, OR across `Alpha`/`Red`/`Green`/`Blue`, every
  channel compared `< threshold` — behaviorally identical to the port; only style differs
  (block-bodied private static vs expression-bodied local fn). No code change was required;
  `results/T2.2_result.md` updated, open questions now empty. `T4.1` no longer needs to re-verify
  this predicate. Phase 2 continues with the `T2.3` build gate.
