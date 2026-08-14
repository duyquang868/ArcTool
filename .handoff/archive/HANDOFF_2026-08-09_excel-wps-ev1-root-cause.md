# ArcTool — HANDOFF ARCHIVE
**Archived:** 2026-08-09
**Phase closed:** Excel to Revit / WPS provider split — EV-1 root-cause isolation only
**Result:** Root cause isolated; no source fix landed in this phase. Continue in a new chat.

---

## What this phase did

This chat closed one bounded phase only: **stop EV-2/EV-3, investigate why EV-1 still fails without Revit, and isolate the real WPS runtime defect before any further source change**.

Delivered:
- runtime re-check evidence: `.claude/workpackages/excel-to-revit-wps-provider-split/evidence/EV-1b_output.txt`
- deeper diagnostic evidence: `.claude/workpackages/excel-to-revit-wps-provider-split/evidence/EV-1c_output.txt`
- origin-isolation evidence: `.claude/workpackages/excel-to-revit-wps-provider-split/evidence/EV-1d_output.txt`
- diagnostic result write-up: `.claude/workpackages/excel-to-revit-wps-provider-split/results/EV-1d_root_cause_result.md`
- package state updates in `06_EXECUTION_STATE.md`

No source file was edited in this phase. No build gate rerun. No Revit. No smoke test.

---

## Key conclusion

The real defect is **not** the `Workbooks.Open(...)` argument list.

What the runtime evidence proved:
- `KET.Application` resolves and activates.
- `Application.Version`, `Visible`, `DisplayAlerts`, and `Quit` work.
- But `Application.Workbooks` returns **null** on this machine/WPS build.
- This reproduced across:
  - `Activator.CreateInstance(Type.GetTypeFromProgID("KET.Application"))`
  - `Marshal.GetActiveObject("KET.Application")`
  - reflection `InvokeMember(GetProperty)`
  - PowerShell COM adapter (`$app.Workbooks`)
  - VB `Interaction.CallByName(..., Get)`
  - a 10x500ms readiness retry loop

Therefore:
- EV-1 / EV-1b `DISP_E_TYPEMISMATCH` at `Workbooks.Open(...)` was a **downstream symptom** of binding on a null workbook collection.
- `T5.1b` was build-clean but targeted the wrong layer. Adding more `Open(...)` argument shapes cannot solve a null `Workbooks` surface.

---

## Strongest evidence excerpts

From `EV-1b_output.txt`:
- all four patched `Open(...)` shapes failed with the same `DISP_E_TYPEMISMATCH`
- wider diagnostic shapes also failed
- downstream members never executed because workbook open never succeeded

From `EV-1c_output.txt`:
- first null-checked probe showed `Workbooks get returned null`
- both headless and visible rounds showed null

From `EV-1d_output.txt`:
- `CreateInstance(KET.Application)` probe:
  - `Application.Version : 12.0`
  - `reflection once      -> NULL`
  - `ps-adapter once      -> NULL`
  - `CallByName once      -> NULL`
  - retry loop 10/10 remained NULL
- `GetActiveObject(KET.Application)` probe:
  - same result: all binders NULL, retry loop 10/10 NULL
- registry surface looked abnormal too:
  - CLSID resolved: `{45540001-5750-5300-4b49-4e47534f4655}`
  - `HKCR\CLSID\{...}\LocalServer32 = ` empty
  - `InprocServer32 = ` empty
  - `TypeLib = ` empty

Interpretation: the current WPS automation surface appears partial/non-standard in this environment — enough for some application-level members, not enough to yield an Excel-shaped workbook collection.

---

## Package state after this phase

`06_EXECUTION_STATE.md` now reflects:
- `T5.1` = `BLOCKED`
- `T5.1b` = `PASS (build) / INSUFFICIENT (runtime)`
- `EV-1b` = runtime BLOCKED result recorded
- `EV-1c` = diagnostic PASS
- `EV-1d` = diagnostic PASS

`EV-2` and `EV-3` should not be run until the WPS entry surface is corrected.

---

## Correct next direction

The next fix must target **WPS workbook/document-surface acquisition**, not `Open(...)` permutations.

Research/fix target inside `ArcTool.Core/Services/Excel/WpsWorkbookPdfExporter.cs`:
- identify which member on the active WPS application actually yields a usable workbook/document surface on this machine
- candidate members to probe first:
  - `ActiveWorkbook`
  - `ActiveWindow`
  - `Documents`
  - ET-specific collection/property names
  - application-level open method, if any, instead of `Application.Workbooks.Open(...)`
- if no workbook/document surface is usable, treat this WPS automation surface as unsupported and fail fast cleanly instead of pretending `Workbooks` is valid

The next phase is a **new chat** focused only on fixing this WPS runtime boundary.

---

## Scope and discipline notes preserved

- The user explicitly decided that EV-1 is worth running locally because it needs no Revit.
- EV-2/EV-3 are intentionally paused to avoid wasted time.
- No source edits were made in this phase; it was pure runtime diagnosis + package-state update.
- This phase should not be mixed with closure/persistence/reindex work. It ends at root-cause isolation.
