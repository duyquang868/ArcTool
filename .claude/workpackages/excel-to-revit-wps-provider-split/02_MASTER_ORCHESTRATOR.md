# EXCEL TO REVIT — WPS PROVIDER SPLIT — MASTER ORCHESTRATOR

Use this file in the master chat that coordinates the package.

---

## Startup

1. Bootstrap only when the session is actually working this package.
2. Read only: `01_SHARED_CONTRACT.md`, `03_TASK_MANIFEST.md`, `05_RESULT_SCHEMA.md`,
   `06_EXECUTION_STATE.md`.
3. Read `04_EVIDENCE_QUEUE.md` only when a ready task needs operator evidence or a worker returns
   `BLOCKED` for evidence.
4. Do not read `CLAUDE.md` in full unless dispatching `T6.1`.
5. Do not paste full source files into worker prompts. `ExcelInteropService.cs` is 585 lines — name
   the file and the line range, let the worker read it.
6. Give each worker exactly: the shared contract, one task file, the immediate upstream result file
   when the chain depends on it, and the minimum evidence excerpt.

---

## Dispatch rules

- One worker = one task file. Every dispatch carries `model: "sonnet"`.
- Phase 1 (`T1.1 → T1.2 → T1.3`) is a light linear micro-chain: prefer one workflow script or one
  master turn that carries only the prior result path forward.
- `T3.1` ‖ `T3.2` may run concurrently. `T3.5` ‖ `T3.6` may run concurrently. Nothing else in this
  package runs in parallel.
- No task uses `isolation: "worktree"`. Source writes are serialized by the manifest lock summary.
- Do not start a task until every dependency is `PASS`.
- Workers return only the compact envelope. Read a result file only to feed the next sequential
  task, resolve a contradiction, or prepare closure.
- Only the master updates `04_EVIDENCE_QUEUE.md` and `06_EXECUTION_STATE.md`.
- Only the master asks the human for evidence.

---

## Phase gates

- Stop the package if `T1.3` returns `NO_GO`.
- Phase 3 blocked until `T2.3` is `PASS`.
- `T3.4` (delete legacy file) blocked until `T3.3` is `PASS`.
- Phase 4 blocked until `T3.7` is `PASS`.
- Phase 6 blocked until `T5.3` is `PASS`.
- If `T3.7` fails the isolation audit, re-dispatch the offending author task rather than patching the
  file from the master.

---

## Evidence routing

This machine has **both engines** as of 2026-08-09: MS Excel, plus WPS Office 12.1.0.28032 with
`KET.Application` resolving (CLSID `45540001-5750-5300-4b49-4e47534f4655`, per-user/HKCU registration
only). `ET.Application` and `Kingsoft.ET.Application` still resolve to null here. `EV-1`, `EV-2`, and
`EV-3` therefore all run on this one machine — no second machine is needed.

The earlier "no WPS on the dev machine" probe is superseded and void; do not carry it forward.

Runtime is still operator-owned (R1). Local WPS availability does **not** authorize the master or any
worker to launch WPS, Excel, or Revit. `EV-1`/`EV-2`/`EV-3` stay operator-run runbooks unless the user
explicitly says otherwise.

1. `T4.2` writes the runbooks. The master then opens `EV-1`, `EV-2`, `EV-3` in the queue.
2. Ask the human only for the exact runs named there.
3. When evidence arrives, record paths first, then forward the minimum routing excerpt.
4. Heavy artifacts (PDFs, PNGs, screenshots, journals) go to the worker by **path**. The master does
   not open them.
5. Record the handoff in `04_EVIDENCE_QUEUE.md`; update `06_EXECUTION_STATE.md`.

---

## Write-lock policy

- `Services/Excel/ISpreadsheetPdfExporter.cs` → `T2.1`
- `Services/Excel/PdfRasterImageService.cs` → `T2.2`
- `Services/Excel/MsExcelWorkbookPdfExporter.cs` → `T3.1`
- `Services/Excel/WpsWorkbookPdfExporter.cs` → `T3.2` → `T5.1b` (only if opened)
- `Services/Excel/SpreadsheetImageExportService.cs` → `T3.3`
- `Services/ExcelInteropService.cs` + `Services/_backup/` → `T3.4`
- `Services/ExcelSyncEngine.cs` → `T3.5`
- `UI/ExcelToRevitWindow.xaml.cs` → `T3.6`
- `ArcTool.Core.csproj` → nobody; a worker that thinks it needs editing returns `BLOCKED`
- Durable memory / dossier / `CLAUDE.md` → `T6.1` only

---

## Fresh-chat bootstrap prompt

```text
Read `.claude/workpackages/excel-to-revit-wps-provider-split/01_SHARED_CONTRACT.md`,
`03_TASK_MANIFEST.md`, `05_RESULT_SCHEMA.md`, and `06_EXECUTION_STATE.md` from that folder.
Read `04_EVIDENCE_QUEUE.md` only if a ready task needs operator evidence.

Act as the master orchestrator for this work package.
Do not auto-load the full package. Do not read heavy artifacts; route them to workers by path.
Dispatch every worker with `model: "sonnet"`.
Respect all dependencies and exclusive write scopes; no worktree isolation in this package.
Return only compact envelopes from workers.
Never launch Revit, Excel, WPS, or any smoke test — ask the human via the runbooks in the evidence queue.
Update `06_EXECUTION_STATE.md` after every worker result. Stop on any `NO_GO` gate.
```
