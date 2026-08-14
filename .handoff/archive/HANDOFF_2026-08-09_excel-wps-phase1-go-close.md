# ArcTool — HANDOFF TO NEXT SESSION
**Updated:** 2026-08-09
**Status:** ACTIVE — Excel to Revit / WPS PDF-export provider split: **PHASE 1 CLOSED, verdict GO**. Phase 2 (`T2.1`, `T2.2`, `T2.3`) is cleared and NOT started. Continue in a new chat.

> Previous handoff (package authored + WPS fact correction) is archived at
> `.handoff/archive/HANDOFF_2026-08-09_excel-wps-package-fact-corrected.md`.
> Its package layout table, symbol map, write-lock policy, tool notes, and locked user decisions are still
> valid and are not repeated in full here — read it if you need that detail.

---

## What this phase closed

Phase unit for the chat that just ended: **dispatch Phase 1 of the work package as a linear worker chain
`T1.1 → T1.2 → T1.3`, master-orchestrated, analysis only.**

Delivered:
- three worker result files under `.claude/workpackages/excel-to-revit-wps-provider-split/results/`
- `06_EXECUTION_STATE.md` rows `T1.1`/`T1.2`/`T1.3` = `PASS`, three history lines appended
- `T1.3` gate verdict **GO** → the package's only stop-everything gate did not fire

Not done, by design: no source file created, modified, or deleted; no `Services/Excel/` folder; no csproj
change; no build; no Revit/Excel/WPS launch; no MCP call; no re-index. Verified by `git status` — the only
churn under `ArcTool.Core/Services/` is pre-existing Quick Dimension work, unrelated to this package.

---

## Phase 1 results — where the design now lives

| Task | Status | Result file | Payload |
|---|---|---|---|
| `T1.1` | PASS | `results/T1.1_result.md` | locked session interface, MS COM member inventory with Microsoft Learn citations, 4 sourced `Xl*` values, COM release order |
| `T1.2` | PASS | `results/T1.2_result.md` | ProgID probe order + rationale, **27-row `InvokeMember` call-shape table**, numeric constant table, failure semantics, late-bound RCW release rule |
| `T1.3` | PASS, **GO** | `results/T1.3_result.md` | six reasoned gate answers, then the UNVERIFIED carry-forward list |

`T3.2` implements the `T1.2` table verbatim. `T4.2` turns its assumption list into the `EV-1` runbook.
Do not re-derive any of this in the next chat — read the result file.

---

## Locked design output (the minimum Phase 2 needs)

```csharp
public interface ISpreadsheetPdfExporter : IDisposable
{
    SpreadsheetEngine Engine { get; }
    bool Open(string filePath);                                  // swallow, no throw
    IReadOnlyList<string> GetSheetNames();                        // swallow → empty
    IReadOnlyList<string> GetNamedRanges(string sheetName);        // swallow → empty
    bool ExportRegionToPdf(string sheetName, string regionName, string outputPdfPath);
}
```

Confirmed by `T1.1`/`T1.3`:
- signature list is engine-neutral **by inspection** — only `string`, `bool`, `IReadOnlyList<string>`;
  no `Range`/`Worksheet` can leak into `WpsWorkbookPdfExporter.cs`
- region resolution (NamedRange → PrintArea → UsedRange) stays **inside** the provider
- temp-PDF deletion moves provider → coordinator. This is the **one contract-mandated R11 relocation**,
  documented and accepted, not a defect
- `ExcelSyncEngine.cs:162` can still throw its user-facing `InvalidOperationException` on open failure,
  so the observable outcome survives
- constants, all sourced, none `UNSOURCED`: `xlTypePDF=0`, `xlQualityStandard=0`, `xlPaperEsheet=26`,
  `xlPaperA3=8`
- COM release order child → parent, workbook → app, for both early- and late-bound paths

ProgID chain locked: `KET.Application` → `ET.Application` → `Kingsoft.ET.Application`.
`KWPS.Application` is **explicitly denylisted** — it resolves but is Writer, not the spreadsheet app.

Failure classes the coordinator must distinguish (so users get the right message):
`EngineAbsent` / `EngineFoundOpenFailed` / `EngineFoundExportFailed`, plus a **new post-export
`File.Exists` check** justified by the WPS shell-mode server risk.

`ArcTool.Core.csproj` needs **no edit** — the SDK-style glob picks up `Services/Excel/` automatically and
the `COMReference` with `EmbedInteropTypes=true` stays as-is. A worker that thinks otherwise returns
`BLOCKED`; nobody owns that file.

---

## Carried forward as UNVERIFIED — routed to Phase 5, non-blocking for Phase 2-4

Every remaining WPS behavioral claim is an `ASSUMPTION`, not a verified fact. The MS branch is what must
be fully verifiable for the code phases; WPS questions go to `EV-1`/`EV-2` and must **not** stall Phase 2-4.

Two specific open questions from `T1.2`:
1. whether `Range.Address[false,false]` on a WPS `IDispatch` needs `GetProperty` alone or `GetProperty`
   combined with `InvokeMethod` — the single most error-prone late-bound call in the file
2. whether `KET.Application`'s spreadsheet mode accepts the 8-argument `ExportAsFixedFormat` signature as-is

Plus, unchanged from the contract: PDF fidelity vs Excel, named-range scope parity, protected-sheet
PageSetup behavior, WPS-side numeric constant equivalence.

---

## Next-session starting point

Start a **NEW chat**. Phase 2 is the **first phase that writes source code** — the user's earlier
"chưa sửa code" hold applies here, so confirm the go-ahead before dispatching.

Phase 2 shape: `T2.1` (writes `Services/Excel/ISpreadsheetPdfExporter.cs`) ‖-free, then
`T2.2` (writes `Services/Excel/PdfRasterImageService.cs`), then `T2.3` build gate.
Phase 3 is blocked until `T2.3` is `PASS`.

Build command (path verified present 2026-08-09):

```bash
"/c/Program Files/Microsoft Visual Studio/18/Community/MSBuild/Current/Bin/MSBuild.exe" ArcTool.Core/ArcTool.Core.csproj -property:Configuration=Debug -property:Platform=x64 -verbosity:minimal -nologo
```

### New-chat prompt

```
ArcTool — Excel to Revit / WPS provider split: chạy PHASE 2 của work package.

Package: .claude/workpackages/excel-to-revit-wps-provider-split/
Đọc trước: 01_SHARED_CONTRACT.md, 03_TASK_MANIFEST.md, 05_RESULT_SCHEMA.md, 06_EXECUTION_STATE.md.
Không đọc CLAUDE.md đầy đủ. Không đọc 04_EVIDENCE_QUEUE.md ở phase này.

Phase 1 đã PASS toàn bộ, T1.3 verdict GO. Interface đã chốt trong results/T1.1_result.md,
recipe late-binding trong results/T1.2_result.md. KHÔNG dựng lại thiết kế — đọc result file.

Bạn là master orchestrator. Dispatch:
  T2.1 → T2.2 → T2.3 (build gate)
mỗi task một worker, model: "sonnet", chỉ nhận về envelope theo 05_RESULT_SCHEMA.md.
Chuyển tiếp bằng ĐƯỜNG DẪN result file, không paste nội dung result vào chat master.

Phase 2 LÀ phase đầu tiên ghi source: chỉ tạo
  ArcTool.Core/Services/Excel/ISpreadsheetPdfExporter.cs   (T2.1)
  ArcTool.Core/Services/Excel/PdfRasterImageService.cs     (T2.2)
Không sửa ArcTool.Core.csproj (không task nào own file đó → BLOCKED).
Cập nhật 06_EXECUTION_STATE.md sau mỗi worker. Phase 3 khoá đến khi T2.3 PASS.

Không tự chạy Revit/Excel/WPS/MCP/smoke test — runtime do tôi chạy.
```

---

## Master-discipline notes that held this phase (keep doing this)

- each dispatch carried only: the shared contract, one task file, and the upstream result-file **PATH**.
  Never paste worker result content into the master chat — that is what forced an avoidable compact on
  EV-1 (2026-08-07, `Memory/feedback_master_context_discipline.md`)
- master consumed only the three 25-line `<MICRO_RESULT>` envelopes
- `CLAUDE.md` was not read in full; `04_EVIDENCE_QUEUE.md` was not read at all
- "no source modified" was **verified with `git status`**, not asserted
- when an `Edit` on `06_EXECUTION_STATE.md` failed with "String to replace not found", the cause both
  times was that the row had already been updated earlier in the same session. Re-read the file to get
  the current text instead of retrying the stale match

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

1. One chat = one phase. Phase 1 is closed; Phase 2 dispatch starts in a new chat.
2. Revit/Excel/WPS runtime is operator-controlled: no launch, `.rvt` open, MCP call, or smoke test
   without an explicit request (R1). Local WPS availability does **not** authorize a worker to drive it.
3. The MS Excel exporter file and the WPS exporter file stay separate; they converge only at the PDF on disk.
4. `WpsWorkbookPdfExporter.cs`: late-bound only, no `Microsoft.Office.Interop.Excel`, no `Xl*` enum,
   numeric constants with recorded provenance. Verified by static grep, not by assertion (R9).
5. Engine precedence: MS Excel first, WPS fallback, auto-detected, no UI picker, no settings override (R10).
6. The PDF → PDFtoImage 300 DPI → PNG → SkiaSharp threshold-240 crop → Revit `ImageType`/`ImageInstance`
   half is shared and must not be duplicated per provider (R11).
7. `ExcelInteropService.cs` is backed up to `Services/_backup/ExcelInteropService.cs.bak` before deletion,
   and only `T3.4` may delete it (R12).
8. Every worker/subagent dispatch carries `model: "sonnet"`.
9. Durable `Memory/` / `.Dossier` / `CLAUDE.md` writes for this package belong to `T6.1` only — they are
   correctly deferred, not missing.
10. `index_repository` is the final, optional, user-directed step. Never gate closure on it.

---

## Reference files

- Work package: `.claude/workpackages/excel-to-revit-wps-provider-split/`
- Phase 1 results: `.claude/workpackages/excel-to-revit-wps-provider-split/results/T1.1_result.md`,
  `T1.2_result.md`, `T1.3_result.md`
- Previous handoff (package authored + fact correction): `.handoff/archive/HANDOFF_2026-08-09_excel-wps-package-fact-corrected.md`
- Durable memory record: `Memory/project_excel_to_revit_wps_provider_split.md`
- Feature dossier: `.Dossier/Detailed Technical Dossier - Excel to Revit.md` §13.0
- Work package skill/scaffold: `.claude/skills/arctool-work-package/`, `.claude/workpackages/_TEMPLATE/`
- Root operating document: `CLAUDE.md`
