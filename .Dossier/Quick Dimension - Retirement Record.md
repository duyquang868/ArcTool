# Quick Dimension — Retirement Record

Status: RETIRED (closed)
Retired on: 2026-08-10
Retirement package: `.claude/workpackages/retire-quick-dimension/`
Prior mission history (not superseded by this record): `Memory/project_qd_chain_creation_audit_handoff.md`, `.Dossier/Quick Dimension - Implementation Roadmap.md`, `.Dossier/ArcTool Locked Technical Decisions.md`

This is a bounded closure dossier for the decision to retire Quick Dimension from the
active ArcTool product surface. It records why the feature was retired, exactly what
changed on the live surface, the full archive inventory, the build-exclusion mechanism,
verification evidence, and how to revive the feature later if ever needed.

---

## 1. Why retired

- The operator ran EV-4 (Quick Dimension Phase 4, Session 4.4 performance baseline
  evidence intake) and, based on that evidence, concluded Quick Dimension is no longer
  feasible or appropriate to continue developing.
- This is an explicit operator verdict, not a Claude-initiated judgment. No new runtime
  defect triggered this decision; it follows the BUG-10/BUG-11 mission closure
  (2026-08-04, EV-2/EV-3 PASS) and sits downstream of the Phase 4 hardening track.
- Retirement means the feature is removed from the active, user-facing product surface
  and from the active roadmap. It is explicitly NOT a deletion of history: all source,
  roadmap detail, ADRs, and runtime evidence records are preserved as-is.

## 2. What was removed from the live surface

- All 7 Quick Dimension ribbon buttons were removed from `App.OnStartup` in
  `ArcTool.Core/App.cs` (task T2.2 of the retirement package). The Annotation Tools
  ribbon panel no longer advertises any Quick Dimension entry point.
- No other active command registrations, menus, or UI surfaces referenced Quick
  Dimension; `App.OnStartup` was the sole live registration owner per the retirement
  package's source ownership map.
- Non-QD features (Arrange Dimension, Excel to Revit, Coordinate tools) were not
  touched and remain live.

## 3. Archive inventory

All previously active-source Quick Dimension files were moved from
`ArcTool.Core/Commands|Models|Services/` into a dedicated archive area, with namespaces
updated to `ArcTool.Core.Archive.QuickDimension.*` (tasks T2.3/T2.4).

Archive root: `ArcTool.Core/Archive/QuickDimension/`

### `Archive/QuickDimension/Commands/` — 7 files
- `QuickDimensionCreateChainSmokeCommand.cs`
- `QuickDimensionDoorWindowReferenceSpikeCommand.cs`
- `QuickDimensionFullMixedReferenceSpikeCommand.cs`
- `QuickDimensionGridReferenceSpikeCommand.cs`
- `QuickDimensionMixedReferenceSpikeCommand.cs`
- `QuickDimensionReadOnlySummaryCommand.cs`
- `QuickDimensionWallReferenceSpikeCommand.cs`

### `Archive/QuickDimension/Models/` — 8 files
- `QuickDimensionContract.cs`
- `QuickDimensionDoorWindowReferenceProbe.cs`
- `QuickDimensionFullMixedReferenceProbe.cs`
- `QuickDimensionGridReferenceProbe.cs`
- `QuickDimensionMixedReferenceProbe.cs`
- `QuickDimensionWallAxisAggregationTrace.cs`
- `QuickDimensionWallMidRunProbe.cs`
- `QuickDimensionWallReferenceProbe.cs`

### `Archive/QuickDimension/Services/` — 15 files
- `QuickDimensionChainCreationService.cs`
- `QuickDimensionDoorWindowCandidateCollector.cs`
- `QuickDimensionDoorWindowReferenceProbeService.cs`
- `QuickDimensionFullMixedReferenceProbeService.cs`
- `QuickDimensionGeometryService.cs`
- `QuickDimensionGridCandidateCollector.cs`
- `QuickDimensionGridReferenceProbeService.cs`
- `QuickDimensionMixedReferenceProbeService.cs`
- `QuickDimensionReadOnlyEngine.cs`
- `QuickDimensionReadOnlyXmlLogService.cs`
- `QuickDimensionWallAxisAggregatorService.cs`
- `QuickDimensionWallCandidateCollector.cs`
- `QuickDimensionWallMidRunProbeService.cs`
- `QuickDimensionWallReferenceProbeService.cs`
- `QuickDimensionWallSpikeXmlLogService.cs`

Total: 30 archived source files (7 commands + 8 models + 15 services), confirmed by
directory listing at retirement time. This matches the pre-retirement active-source
inventory recorded in the retirement package's shared contract (`01_SHARED_CONTRACT.md`,
section 4).

## 4. Build/project exclusion

- `ArcTool.Core/ArcTool.Core.csproj` gained one new item:
  `<Compile Remove="Archive\QuickDimension\**\*.cs" />` (task T2.5).
- Effect: the archived Quick Dimension source is preserved in the repository working
  tree (readable, greppable, revivable) but is not compiled into `ArcTool.Core.dll`.
  Archived files therefore cannot reference symbols that no longer resolve without
  breaking the build, and they no longer contribute any IL to the live add-in.
- No other `.csproj` or `.cs` files were changed by the archival/exclusion steps beyond
  the ribbon removal in `App.cs` and the moves that changed archived files' own
  namespaces.

## 5. Build verification result

- Command (locked, approved for this repo):
  ```bash
  "/c/Program Files/Microsoft Visual Studio/18/Community/MSBuild/Current/Bin/MSBuild.exe" ArcTool.Core/ArcTool.Core.csproj -property:Configuration=Debug -property:Platform=x64 -verbosity:minimal -nologo
  ```
- Result: PASS (task T2.6, 2026-08-10). MSBuild produced
  `ArcTool.Core/bin/x64/Debug/net8.0-windows/ArcTool.Core.dll` for `Debug|x64` with no
  compile errors or project-file errors in the captured output.
- No Revit runtime evidence was collected or is owed for this retirement mission — the
  retirement package's acceptance gates are static/build-only (`01_SHARED_CONTRACT.md`
  section 7/8, invariant R11). No Revit launch, `.rvt` open, Revit MCP call, or smoke
  test was performed during retirement.

## 6. How to revive Quick Dimension if ever needed

This is a documentation note for a hypothetical future decision, not a plan currently
scheduled or endorsed.

1. Read this record plus `.Dossier/Quick Dimension - Implementation Roadmap.md` (marked
   retired at its top, but the full phase history below that marker is intact) and
   `Memory/project_qd_chain_creation_audit_handoff.md` for the last known-good runtime
   state (BUG-10/BUG-11 CLOSED, EV-2/EV-3 PASS on the 2026-08-03 DLL).
2. Move the files listed in section 3 back from `ArcTool.Core/Archive/QuickDimension/{Commands,Models,Services}/`
   into `ArcTool.Core/{Commands,Models,Services}/` and revert their namespace from
   `ArcTool.Core.Archive.QuickDimension.*` back to `ArcTool.Core.*` (or whatever
   namespace convention is current at revival time).
3. Remove the `<Compile Remove="Archive\QuickDimension\**\*.cs" />` item from
   `ArcTool.Core/ArcTool.Core.csproj` (or narrow it if only some files are revived).
4. Re-add the required ribbon button registrations to `App.OnStartup` in
   `ArcTool.Core/App.cs`.
5. Rebuild with the locked VS MSBuild command above and treat the result as a fresh
   baseline — do not assume the 2026-08-03/2026-08-04 runtime evidence still applies
   without re-verification, since Revit, dependencies, or the rest of ArcTool may have
   changed in the interim.
6. Before resuming any roadmap phase, get an explicit operator decision to un-retire the
   feature; do not resume QD work based on this note alone.

## 7. Cross-references

- Retirement package (task manifest, shared contract, per-task results):
  `.claude/workpackages/retire-quick-dimension/`
- Root operating pointer: `CLAUDE.md` (header paragraph and sections 2/3/6/7)
- Durable memory pointer: `Memory/project_qd_retired_archive.md`
- Prior mission closure (preserved, not superseded): `Memory/project_qd_chain_creation_audit_handoff.md`
- Full historical roadmap (now marked retired at top, history intact below):
  `.Dossier/Quick Dimension - Implementation Roadmap.md`
- Deferred rollback validation track (moot now that the feature is retired, left
  unmodified as history): `.Dossier/Quick Dimension - Deferred Rollback Validation Task.md`
