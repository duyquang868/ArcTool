# ArcTool — HANDOFF TO NEXT SESSION
**Updated:** 2026-08-09
**Status:** ACTIVE — Excel to Revit / WPS PDF-export provider split: work package AUTHORED and fact-corrected, Phase 1 NOT dispatched. Continue in a new chat.

> Previous handoff (research/design phase close) is archived at
> `.handoff/archive/HANDOFF_2026-08-09_excel-wps-provider-split-package-authored.md`.
> Its root-cause analysis, blast radius, and locked architecture are still valid and are not repeated here.

---

## What this phase closed

Phase unit for the chat that just ended: **author the multi-agent work package for the provider split, and correct the environment facts after WPS was installed. No source code.**

Delivered:
- full work package authored at `.claude/workpackages/excel-to-revit-wps-provider-split/` — 6 scaffold files, 20 task files (`T1.1` … `T6.2`), empty `results/`
- one design defect in the previously locked spec found and fixed (see "Correction to the locked spec")
- two dead methods identified and excluded from the port
- **WPS installed on this machine mid-phase** → ProgIDs re-probed and five package files corrected

Not done: no source file created, modified, or deleted. No build. No Revit, Excel, or WPS launched. No MCP call. No re-index.

---

## Package layout

`.claude/workpackages/excel-to-revit-wps-provider-split/`

| File | Role |
|---|---|
| `01_SHARED_CONTRACT.md` | invariants R1-R12, domain model, source ownership map, build command, 10 acceptance gates |
| `02_MASTER_ORCHESTRATOR.md` | dispatch rules, phase gates, evidence routing, write-lock policy, fresh-chat bootstrap prompt |
| `03_TASK_MANIFEST.md` | 6 phases, 20 tasks, dependency graph, source-file lock summary |
| `04_EVIDENCE_QUEUE.md` | `EV-1` / `EV-2` / `EV-3`, all `PENDING` |
| `05_RESULT_SCHEMA.md` | the envelope every worker returns |
| `06_EXECUTION_STATE.md` | 20-row status table, all `PENDING`; master-owned |
| `tasks/T1.1.md` … `tasks/T6.2.md` | one task = one worker |

Phase order and gates:

| Phase | Tasks | Gate |
|---|---|---|
| 1 — preflight | `T1.1 → T1.2 → T1.3` | `T1.3` `NO_GO` stops the package |
| 2 — foundation | `T2.1`, `T2.2`, `T2.3` | Phase 3 blocked until `T2.3` `PASS` |
| 3 — providers + rewire | `T3.1 ‖ T3.2`, `T3.3`, `T3.4`, `T3.5 ‖ T3.6`, `T3.7` | Phase 4 blocked until `T3.7` `PASS` |
| 4 — parity + runbooks | `T4.1`, `T4.2` | — |
| 5 — evidence analysis | `T5.1`, `T5.2`, `T5.3` (`T5.1b` only if a WPS patch is needed) | Phase 6 blocked until `T5.3` `PASS` |
| 6 — closure | `T6.1`, `T6.2` | — |

Only `T3.1 ‖ T3.2` and `T3.5 ‖ T3.6` run in parallel. No task uses worktree isolation.

---

## NEW FACT — WPS is now installed on this machine

The user installed WPS Office **after** the package was created. Re-probed 2026-08-09 with
`[Type]::GetTypeFromProgID`:

| ProgID | Result |
|---|---|
| `KET.Application` | **resolves** — CLSID `45540001-5750-5300-4b49-4e47534f4655` |
| `KWPS.Application` | resolves — CLSID `000209ff-0000-4b30-a977-d214852036ff`, but this is **Writer**, not the spreadsheet app; not a candidate |
| `ET.Application` | null |
| `Kingsoft.ET.Application` | null |
| `WPS.Application` | null |
| `ET.Sheet` | null |
| `Excel.Application` | resolves — CLSID `00024500-0000-0000-c000-000000000046` |

- install path: `C:\Users\ADMIN\AppData\Local\Kingsoft\WPS Office\12.1.0.28032\office6\`, `et.exe` present
- `KET.Application` `LocalServer32` is registered **per-user (HKCU) only** — there is no HKLM entry — and points at
  `wps.exe /prometheus /et /Automation`. The spreadsheet app runs as a mode of the WPS shell, not as a standalone `et.exe` server.

Two consequences that must survive into the code:
1. `KET.Application` stays **first** in the ProgID fallback chain. `ET.Application` and `Kingsoft.ET.Application` remain in the chain for other WPS builds but are unverified on any machine.
2. Detection must not assume machine-wide COM registration — per-user-only registration is the norm here.

**The earlier "no WPS anywhere on this machine" probe is VOID. Do not cite it.** It was recorded at package creation and has been corrected in:
`01_SHARED_CONTRACT.md` §5 and §6, `02_MASTER_ORCHESTRATOR.md` (evidence routing), `04_EVIDENCE_QUEUE.md` (environment constraint + `EV-1`/`EV-2` entries), `06_EXECUTION_STATE.md` (facts block + history), and `tasks/T1.2.md`, `tasks/T3.2.md`, `tasks/T4.2.md`.

Consequence for evidence: `EV-1`, `EV-2`, and `EV-3` now all run on **this one machine**. No separate WPS machine is needed. `EV-2` still requires a way to make the coordinator take the WPS path while MS Excel is present — otherwise it silently tests the MS branch and proves nothing. `T4.2` must solve that in the runbook.

Runtime is still operator-owned (R1). The master and workers do **not** launch WPS, Excel, or Revit. The user runs the runbooks.

---

## Correction to the locked spec (already applied in the contract)

The previously locked "`ISpreadsheetPdfExporter` with one export method" is **not implementable**.
`ExcelToRevitWindow` needs `GetSheetNames()` and `GetNamedRanges(sheet)` to populate the WorkSheet and Region dropdowns, and region resolution (NamedRange → PrintArea → UsedRange) is itself engine-specific COM work.

The abstraction is therefore a disposable **session**: open + enumerate + export-to-PDF. Proposed shape, for `T1.1` to confirm or refine:

```csharp
public interface ISpreadsheetPdfExporter : IDisposable
{
    SpreadsheetEngine Engine { get; }
    bool Open(string filePath);
    IReadOnlyList<string> GetSheetNames();
    IReadOnlyList<string> GetNamedRanges(string sheetName);
    bool ExportRegionToPdf(string sheetName, string regionName, string outputPdfPath);
}
```

The PDF file on disk remains the convergence point, so the user's separation constraint is untouched. Recorded as "Critical correction to the earlier locked spec" in contract §3.

---

## Verified symbol map — `ArcTool.Core/Services/ExcelInteropService.cs` (585 lines)

| Symbol | Lines | Destination |
|---|---|---|
| `OpenFile` | 32-48 | MS provider — `new Application()` at **36** is blocking call #1 |
| `GetActiveSheetName` | 54-66 | **dead code, zero callers — not ported** |
| `ExportPrintAreaAsHighResImage` | 68-98 | **dead code, zero callers — not ported** |
| `GetRuntimeFolder` | 100-109 | `PdfRasterImageService` |
| `GetNativeLibraryCandidates` | 111-127 | `PdfRasterImageService` (retarget `typeof(...)`) |
| `EnsurePdfiumLoaded` | 129-152 | `PdfRasterImageService` |
| `EnsureSkiaSharpLoaded` | 154-177 | `PdfRasterImageService` |
| `ExportRangeInternal` | 188-359 | splits: 201-229 PageSetup + PDF → MS provider (`ExportAsFixedFormat` at **221** is blocking call #2); 231-348 render + crop → raster service; temp cleanup at 357 → coordinator |
| `Dispose` / `ReleaseObject` | 361-388 | MS provider |
| `GetSheetNames` | 402-430 | MS provider |
| `GetNamedRanges` | 443-494 | MS provider |
| `ExportRegion` | 512-582 | region resolution → MS provider; raster half delegates |

Dead-code finding proved by graph `trace_path`, not by grep alone.

---

## User decisions already locked

1. **Legacy file** — back up `ExcelInteropService.cs` to `ArcTool.Core/Services/_backup/ExcelInteropService.cs.bak` first (in-place backup folder, easy to find), then delete it. Owned by `T3.4`.
2. **Provider pick** — auto-detect, MS Excel priority. **No UI picker, no settings override.**
3. **WPS verification** — the user runs the operator runbook. Now on this machine.
4. **No source code yet** — "chưa sửa code" has not been lifted. Phase 1 is analysis-only by design, so Phase 1 does not conflict with it. Phase 2 is the first phase that writes source.

---

## Write-lock policy — one owner per file

| File | Owner |
|---|---|
| `Services/Excel/ISpreadsheetPdfExporter.cs` | `T2.1` |
| `Services/Excel/PdfRasterImageService.cs` | `T2.2` |
| `Services/Excel/MsExcelWorkbookPdfExporter.cs` | `T3.1` |
| `Services/Excel/WpsWorkbookPdfExporter.cs` | `T3.2` → `T5.1b` if opened |
| `Services/Excel/SpreadsheetImageExportService.cs` | `T3.3` |
| `Services/ExcelInteropService.cs` + `Services/_backup/` | `T3.4` |
| `Services/ExcelSyncEngine.cs` | `T3.5` |
| `UI/ExcelToRevitWindow.xaml.cs` | `T3.6` |
| `ArcTool.Core.csproj` | **nobody** — a worker that thinks it needs editing returns `BLOCKED` |
| `Memory/`, `.Dossier`, `CLAUDE.md` | `T6.1` only |

The `COMReference Microsoft.Office.Interop.Excel` at `ArcTool.Core.csproj:20-28` **stays**. The SDK-style glob picks up `Services/Excel/` with no csproj change — `T1.3` confirms this.

---

## Build command (path verified present)

```bash
"/c/Program Files/Microsoft Visual Studio/18/Community/MSBuild/Current/Bin/MSBuild.exe" ArcTool.Core/ArcTool.Core.csproj -property:Configuration=Debug -property:Platform=x64 -verbosity:minimal -nologo
```

There is no unit-test project. Build plus static grep audit is the automated gate; `T2.3` and `T3.7` own it.

---

## Tool notes carried forward

- **Do not use `reg query` in a loop from git-bash.** Backslash mangling through a shell variable produced six false results, including a false "absent" for `Excel.Application`, even with `MSYS_NO_PATHCONV=1`. Use PowerShell `[Type]::GetTypeFromProgID($id)` instead.
- PowerShell regex through bash: avoid character classes containing backslashes (`[^\\]`) — bash eats the backslash and PowerShell reports "Unterminated [] set."
- WebSearch returned nothing usable for WPS COM ProgID or `ExportAsFixedFormat` behavior. `T1.2` is scoped to mark facts `UNSOURCED` rather than guess, and every WPS behavioral fact routes to `EV-1`/`EV-2`.
- `get_architecture` project id: `D-Quang mini-OneDrive - MSFT-Plugin Revit-ArcTool`.
- Broad greps must exclude `.claude/worktrees/agent-*` or they blow the 250-file limit.

---

## Open decision for the next chat

With WPS local, does `EV-1` (ProgID + late-bound member probe) stay an operator-run runbook, or may a worker run the probe directly? Launching WPS COM is runtime action, which R1 and `CLAUDE.md` reserve for the operator unless the user explicitly authorizes it. **Current package assumption: operator-run.** Ask the user before changing it; do not assume.

---

## Next-session starting point

Start a **NEW chat** with the "New-chat prompt" below. It dispatches Phase 1 only.

Carry-forward essentials:
- the bottleneck is `ExportAsFixedFormat`, not file parsing — do not re-litigate ClosedXML / OpenXML SDK / NPOI, they parse and do not render
- MS Excel logic file and WPS logic file stay separate; they meet only at the PDF on disk
- Phase 1 writes no source, only `results/T1.1_result.md` … `results/T1.3_result.md`
- stop immediately if `T1.3` returns `NO_GO` or `BLOCKED`

---

## Parallel open track (unchanged, different phase)

Quick Dimension phase-4 hardening remains open and independent:
- `R8_C07` closed as a concrete negative mid-run verdict; `T3.7` is the durable source
- `C01`/`C02`/`C03` were tested by the user — the gap is publication, not evidence
- next QD action: publish `T3.5_result.md`, then `T3.6_result.md`, then re-evaluate `T3.8`
- detail: `.handoff/archive/HANDOFF_2026-08-09_qd-t38-gate-carryforward.md`

Do not mix this track with the Excel/WPS phase in the same chat.

---

## Invariants to preserve

1. One chat = one phase. Package authoring is closed; Phase 1 dispatch starts in a new chat.
2. Revit runtime is operator-controlled: no Revit launch, `.rvt` open, MCP call, or smoke test without an explicit request. WPS and Excel COM launches are covered by the same rule.
3. The MS Excel exporter file and the WPS exporter file stay separate; they converge only at the PDF.
4. `WpsWorkbookPdfExporter.cs`: late-bound only, no `Microsoft.Office.Interop.Excel`, no `Xl*` enum, numeric constants with recorded provenance. Verified by static grep, not by assertion.
5. Engine precedence: MS Excel first, WPS fallback, auto-detected, no UI picker, no settings override.
6. The PDF → PDFtoImage 300 DPI → PNG → SkiaSharp threshold-240 crop → Revit `ImageType`/`ImageInstance` half is shared and must not be duplicated per provider.
7. Excel to Revit dossier invariants stay in force: COM release child → parent, `StoredWidth`/`StoredHeight` in millimetres, two-transaction image create/resize, local-time drift checks, `_suppressRowEvents` discipline.
8. Every worker/subagent dispatch carries `model: "sonnet"`.
9. `index_repository` is the final, optional, user-directed step. Never gate closure on it.

---

## New-chat prompt

```
ArcTool — Excel to Revit / WPS provider split: chạy PHASE 1 của work package.

Package: .claude/workpackages/excel-to-revit-wps-provider-split/
Đọc trước: 01_SHARED_CONTRACT.md, 03_TASK_MANIFEST.md, 05_RESULT_SCHEMA.md, 06_EXECUTION_STATE.md.
Không đọc CLAUDE.md đầy đủ. Không đọc 04_EVIDENCE_QUEUE.md ở phase này.

Bạn là master orchestrator. Dispatch Phase 1 như một chain tuyến tính:
  T1.1 → T1.2 → T1.3
mỗi task một worker, model: "sonnet", chỉ nhận về envelope theo 05_RESULT_SCHEMA.md.
Chuyển tiếp bằng ĐƯỜNG DẪN result file, không paste nội dung result vào chat master.

Phase 1 là phân tích, KHÔNG sửa source. Chỉ ghi results/T1.1_result.md … T1.3_result.md.
Cập nhật 06_EXECUTION_STATE.md sau mỗi worker.
Dừng ngay nếu T1.3 trả về NO_GO hoặc BLOCKED, báo lại cho tôi.

Lưu ý môi trường: máy này CÓ CẢ MS Excel và WPS Office 12.1.0.28032.
KET.Application resolve được (per-user/HKCU); ET.Application và Kingsoft.ET.Application vẫn null.
Không tự chạy Revit/Excel/WPS/MCP/smoke test — runtime do tôi chạy.
```

---

## Reference files

- Work package: `.claude/workpackages/excel-to-revit-wps-provider-split/`
- Previous handoff (research/design close): `.handoff/archive/HANDOFF_2026-08-09_excel-wps-provider-split-package-authored.md`
- Durable memory record: `Memory/project_excel_to_revit_wps_provider_split.md`
- Feature dossier: `.Dossier/Detailed Technical Dossier - Excel to Revit.md` §13.0
- Work package skill/scaffold: `.claude/skills/arctool-work-package/`, `.claude/workpackages/_TEMPLATE/`
- Root operating document: `CLAUDE.md`
