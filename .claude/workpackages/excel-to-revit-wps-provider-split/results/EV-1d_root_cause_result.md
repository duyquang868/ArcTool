# EV-1d — Root-cause isolation for WPS COM failure

- status: PASS
- date: 2026-08-09
- inputs_read: `.claude/workpackages/excel-to-revit-wps-provider-split/evidence/EV-1d_probe.ps1`; `.claude/workpackages/excel-to-revit-wps-provider-split/evidence/EV-1d_output.txt`; `.claude/workpackages/excel-to-revit-wps-provider-split/evidence/EV-1c_output.txt`; `.claude/workpackages/excel-to-revit-wps-provider-split/evidence/EV-1b_output.txt`; `ArcTool.Core/Services/Excel/WpsWorkbookPdfExporter.cs`
- write_scope_touched: `.claude/workpackages/excel-to-revit-wps-provider-split/results/EV-1d_root_cause_result.md`

## Findings
- The real defect is **upstream of `Workbooks.Open`**. On this machine and WPS build (`12.1.0.28032`), `KET.Application` activates successfully and returns a COM application object, but `Application.Workbooks` itself returns `null`.
- This was reproduced across:
  - a newly created COM instance via `Activator.CreateInstance(Type.GetTypeFromProgID("KET.Application"))`
  - an already-running instance via `Marshal.GetActiveObject("KET.Application")`
  - three different property binders: reflection `InvokeMember(GetProperty)`, PowerShell COM adapter (`$app.Workbooks`), and Visual Basic `Interaction.CallByName(..., Get)`
  - a 5-second readiness retry loop (10 retries at 500 ms) after app activation
- Because `Workbooks` is null, the later EV-1 / EV-1b `DISP_E_TYPEMISMATCH` at `Workbooks.Open(...)` was a downstream symptom of binding on a null/empty target, not evidence that the `Open` signature itself merely needed more optional arguments.
- Registry / COM registration evidence is also abnormal for a normal LocalServer automation class:
  - `Type.GetTypeFromProgID("KET.Application")` resolves a CLSID `{45540001-5750-5300-4b49-4e47534f4655}`
  - but `HKCR\CLSID\{45540001-5750-5300-4b49-4e47534f4655}\LocalServer32`, `InprocServer32`, and `TypeLib` all read empty in this environment
- This strongly suggests the current WPS exposure is a **partial or non-standard automation surface** on this machine: enough for `Application.Version`, `Visible`, `DisplayAlerts`, and `Quit`, but not enough to provide a usable workbook collection through the `KET.Application` object the .NET late binder sees.
- Therefore `T5.1b` was a good-faith but incorrect patch direction: it changed the `Open(...)` argument matrix, while the actual failing boundary is `Application.Workbooks` returning null.

## Decision
Do **not** spend any time on EV-2 or EV-3 until the WPS entry point is corrected. The next fix must target **WPS app acquisition / workbook-surface acquisition**, not the `Open` argument list.

## Recommended resolution direction
The most likely viable fix direction is to stop assuming `KET.Application.Workbooks` behaves like Excel and instead probe for the workbook/document surface that this WPS build actually exposes. Candidate strategies, in descending priority:
1. Inspect alternative app/document members on the active `KET.Application` object (`ActiveWorkbook`, `ActiveWindow`, `Documents`, `RecentFiles`, ET-specific collection names) and build the provider around the surface that is non-null.
2. Probe whether opening the file must happen through an application-level method instead of `Application.Workbooks.Open(...)` on this build.
3. If WPS only exposes usable automation through a different ProgID / object model on this machine, rework provider detection to target that surface explicitly rather than the current Excel-shaped assumption.
4. Treat the current `KET.Application` automation surface as unsupported for this machine if no workbook/document surface can be found, and fail fast with `EngineAbsent`/`EngineFoundOpenFailed` semantics rather than pretending `Workbooks` is valid.

## Handoff for downstream tasks
- Authoritative evidence: `.claude/workpackages/excel-to-revit-wps-provider-split/evidence/EV-1d_output.txt`
- The next investigation/fix should stay within exclusive write scope `ArcTool.Core/Services/Excel/WpsWorkbookPdfExporter.cs`
- The immediate research target is **which member on this WPS automation object actually yields an openable workbook/document surface**, not more `Open(...)` argument permutations.
