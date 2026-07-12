# Quick Dimension — Implementation Roadmap

Last updated: 2026-07-12
Status: Active roadmap for the Quick Dimension feature in ArcTool.

## Purpose

This document is the long-form source of truth for the Quick Dimension feature. It exists so the roadmap can survive across many sessions without bloating `CLAUDE.md` or requiring repeated high-noise edits to the root operating document.

## Scope lock

The current MVP scope is:
- Active Revit Plan View only.
- Main-flow input: select one straight non-curtain host Wall and pick a placement side; the old two-picked-points input is retained only as a deprecated/optional cross-cutting path.
- Main-flow sources: selected wall end anchors plus hosted Door and Window openings in that wall; Grid and non-selected-wall sources are disabled in the wall-axis projection dispatch.
- Output: chain dimension and optional total dimension.
- No linked models, no column support, no arc host wall/grid support in the main flow, no rubberband preview, no automatic grouping.

## MODEL PIVOT — 2026-06-11 (ADR-2026-06-11): cross-cutting intersection → wall-axis projection

The original scope-lock above (two picked points + intersection across many elements) is SUPERSEDED for the main flow. It is retained as history because the Phase 1/2 spike and intersection code remain in the tree as evidence and as an optional/legacy path.

Why the pivot:
- The 2026-06-11 Revit smoke test confirmed the intersection model answers "what does the drawn line physically cross", which is structurally the wrong question for the user's intent.
- Symptoms: Window collected 17 / accepted 0; Door returned a single jamb instead of both; walls running parallel to the drawn line were rejected by the `ParallelToDimensionLine` guard — exactly the walls the user wants to dimension along.
- The desired output (worked example chain 406/406/915 etc.) measures openings ALONG one wall, which is how Revit "dimension along a wall" behaves.

New main-flow model:
- Input: user selects ONE host Wall, then picks a side (left/right). No drawn dimension line.
- Axis: the selected wall's straight `LocationCurve` IS the dimension axis, even when the wall is skewed in plan. Arc/non-line host walls are rejected with diagnostics.
- References gathered ONLY from the selected wall: its two end edges plus every hosted Door/Window opening, each opening contributing BOTH left and right jambs.
- Participation test: project each reference point onto the wall axis via `QuickDimensionLineContext.ProjectParameter`; keep when the projected parameter is within `[0, Length]`. This replaces `TryIntersectSegmentWithDimensionLine2D` and the `IsNearlyParallel` parallel guard for the main flow.
- Jamb points are built along the WALL direction. Do NOT mix wall direction with a separate drawn-line direction — that mixing was the root cause of the window projection failure.
- Side sign: the left/right pick is captured in the contract from Phase 2 onward so Phase 3 placement has enough information.
- Grids and non-selected walls drop out of the main flow; they may return later as an opt-in/legacy cross-cutting path.

Carry-over (unchanged by the pivot):
- Wall side-face reference strategy `HostObjectUtils.GetSideFaces()` and `QuickDimensionReferenceStrategy.WallSideFace`.
- Opening Left/Right reference strategy `FamilyInstance.GetReferences(Left/Right)` with `HostWallOpeningGeometry` fallback.
- Source-aware, conservative dedupe so both jambs of one instance survive.
- Read-only-before-NewDimension discipline; diagnostics-first; do not destabilize existing commands.
- Contract rule: store `ElementId`/`Reference`/`XYZ`/diagnostics, never live `Element` objects.

Contract impact:
- `QuickDimensionLineContext` is rebuilt from a Wall + side sign (not two picked points); axis/direction/length derive from the wall curve.
- `QuickDimensionCandidate.ParameterOnDimensionLine` now means projected coordinate on the wall axis (not intersection point on a drawn line).
- Phase 1 spike services/models/commands (`QuickDimension*ReferenceProbe*`, `*ReferenceSpikeCommand`) are UNTOUCHED — they have their own models and `CreateDimensionLine` and do not depend on `QuickDimensionLineContext`.

## Non-negotiable invariants

- Revit `Reference` correctness comes before UI polish.
- Read-only extraction and sort/dedupe must be proven before any production `NewDimension` call.
- Door and Window are first-class source types, not edge cases.
- Unsupported cases must fail loudly with diagnostics, not silently.
- The feature must not destabilize existing ArcTool commands.

## Roadmap structure

### Phase 0 — Preparation and safety

#### Session 0.1 — Baseline, branch, scope lock
- Verify repo state and separate pre-existing changes from Quick Dimension changes.
- Lock the MVP scope.
- Confirm baseline build status before touching feature code.
- Decide the file boundary for all Quick Dimension work.

Pass criteria:
- Project state is understood.
- MVP scope is frozen.
- Build baseline is documented.

Blocked if:
- Baseline build is unknown.
- Scope is still ambiguous.

Completion record — 2026-05-27:
- Repo state reviewed before Quick Dimension implementation: working tree already contains many pre-existing changes and generated `.vs` / `Obj` artifacts; no Quick Dimension source code has been added in Phase 0.
- Code map verified against the current folder structure: `Commands`, `Services`, `UI`, `Models`, `Utilities`, `Resources`, `Properties`, and `.Dossier` still match the root operating file.
- Knowledge graph status checked: project index is ready under `D-Quang mini-OneDrive - MSFT-Plugin Revit-ArcTool`; graph search confirms existing dimension-related code is limited to `ArrangeDimensionCommand` and `LinearDimensionSelectionFilter`.
- MVP scope is frozen: active Plan View only; two picked points define the dimension line; sources are Grids, Wall boundaries, hosted Door openings, and hosted Window openings; output is chain dimension plus optional total dimension.
- File boundary is locked for implementation: add Quick Dimension command/service/model/filter files under existing `Commands`, `Services`, `Models`, and `Utilities`; defer WPF settings UI until the reference engine is proven.
- Baseline build command attempted with `dotnet build ArcTool.slnx --no-restore`, but the current shell has no `dotnet` executable. Treat this as a documented environment limitation, not a source-code result; the first Windows/Revit-side check before Phase 1 code must confirm the build in the normal developer environment.

### Phase 1 — Reference feasibility spikes

#### Session 1.1 — Grid reference spike
- Prove a stable `Reference` strategy for Grid dimensions.
- Test `new Reference(grid)` and `grid.Curve.Reference` against `NewDimension`.
- Record which strategy works in Revit 2026.

Closure record — 2026-05-29:
- Added and runtime-tested temporary ribbon command `QuickDimensionGridReferenceSpikeCommand` under Annotation Tools as `QD Grid Spike`.
- Runtime result in Revit 2026: `new Reference(grid)` PASS for straight Grid dimensions; `grid.Curve.Reference` FAIL with `Invalid number of references`; Grid MVP must use element references.
- Spike collector caveat: midpoint projection can over-collect in mixed/slanted grid layouts; production Grid collector must use true line intersection between the picked dimension line and each straight grid line before sorting.
- Build status remains environment-limited: the current shell still has no `dotnet`; compile/load verification must run in the normal Windows/Revit developer environment.

#### Session 1.2 — Wall face reference spike
- Prove a stable `Reference` strategy for wall boundary faces.
- Use `Options.ComputeReferences = true` and evaluate planar faces.
- Lock which wall face rule is acceptable for the MVP.

Implementation record — 2026-05-31:
- Added temporary ribbon command `QuickDimensionWallReferenceSpikeCommand` under Annotation Tools as `QD Wall Spike`.
- Two reference strategies tested:
  1. `HostObjectUtils.GetSideFaces(wall, ShellLayerType.Exterior/Interior)` — recommended Revit API approach for host objects, returns `Reference` directly.
  2. `Options.ComputeReferences = true` + `Face.Reference` from planar faces — general geometry extraction approach.
- Wall filtering logic: skip curtain walls, skip walls parallel to dimension line (dot product < 0.02), skip arc walls, skip walls outside picked span.
- Face selection logic: choose the face closer to the dimension line (perpendicular distance from face centroid to dimension line).
- Runtime validation in Revit 2026: `HostObjectUtilsSideFaces` PASS consistently; `GeometryComputeReferences` PASS on axis-aligned walls but FAIL on rotated walls due to face normal filtering bug.

Closure record — 2026-05-31:
- **Locked strategy for MVP:** `HostObjectUtils.GetSideFaces()` with closest-face selection.
- **Rationale:** API is purpose-built for host objects, simpler code, consistent PASS results.
- **Face selection rule:** When both Exterior and Interior faces exist, pick the face whose centroid is closer to the dimension line (perpendicular distance). This ensures dimension snaps to the wall edge nearest the picked dimension line position.
- `GeometryComputeReferences` approach is not needed for walls; may be revisited for Door/Window openings in Session 1.4.

#### Session 1.3 — Mixed Grid + Wall reference array spike
- Prove Grid and Wall references can coexist in the same `ReferenceArray`.
- Validate sort order and failure behavior when references are intentionally reversed.

Implementation record — 2026-06-01:
- Added temporary ribbon command `QuickDimensionMixedReferenceSpikeCommand` under Annotation Tools as `QD Mixed Spike`.
- Created unified candidate model `QuickDimensionMixedCandidate` with `SourceType` enum (Grid/Wall) for merged sorting.
- Service collects Grid candidates using Session 1.1 proven strategy: `new Reference(grid)`.
- Service collects Wall candidates using Session 1.2 proven strategy: `HostObjectUtils.GetSideFaces()` with closest-face selection.
- All candidates merged and sorted by `ParameterOnDimensionLine` before testing.
- Four test scenarios implemented:
  1. `SortedByPosition` — references in physical order along dimension line (primary success criterion).
  2. `ReversedOrder` — references in reverse order to test if Revit rejects or auto-sorts.
  3. `GridsOnly` — baseline comparison with grids only.
  4. `WallsOnly` — baseline comparison with walls only.

Closure record — 2026-06-01:
- Runtime validation in Revit 2026: ALL SCENARIOS PASS.
- **Primary conclusion: Mixed Grid + Wall references work in the same ReferenceArray.**
- ReversedOrder also PASS — Revit auto-sorts references internally, does not reject reversed order.
- Test coverage: horizontal, vertical, and diagonal dimension lines across various grid/wall configurations.
- **Known limitation confirmed**: Midpoint projection causes inaccurate collection on slanted/diagonal dimension lines. Some grids/walls that visually intersect the dimension line are missed because their midpoint falls outside the picked span. This is acceptable for the spike but must be fixed in Phase 2 with true line-line intersection.
- Session 1.3 objective achieved: reference compatibility proven. Collection accuracy deferred to Phase 2 geometry service.

#### Session 1.4 — Door/Window opening reference spike
- Prove a stable `Reference` strategy for hosted Door and Window openings.
- Test family-instance geometry, host wall opening geometry, and family reference-plane strategies.
- Lock the supported opening strategy or stop the roadmap if no stable strategy exists.

Implementation record — 2026-06-01:
- Added temporary ribbon command `QuickDimensionDoorWindowReferenceSpikeCommand` under Annotation Tools as `QD Door/Win Spike`.
- Created model file `QuickDimensionDoorWindowReferenceProbe.cs` with candidate, strategy result, and summary types.
- Created service file `QuickDimensionDoorWindowReferenceProbeService.cs` testing three strategies.
- Three reference strategies tested:
  1. `FamilyInstance.GetReferences(FamilyInstanceReferenceType.Left/Right)` — API-recommended approach using family reference planes.
  2. `Options.ComputeReferences = true` + `Face.Reference` — general geometry extraction approach.
  3. Host Wall opening geometry via vertical edge detection within instance bounding box.
- Candidate filtering logic: skip non-wall-hosted instances, skip instances in walls parallel to dimension line, skip instances outside picked span.
- Reference extraction: each strategy extracts left and right references independently; candidates track all six possible references.

Closure record — 2026-06-01:
- Runtime validation in Revit 2026: 7 test cases covering horizontal, vertical, and multiple diagonal dimension lines.
- **FamilyInstanceReferences: 100% PASS** — works consistently across all test cases including complex diagonal layouts.
- **HostWallOpeningGeometry: 100% PASS** — backup strategy also works consistently.
- **GeometryComputeReferences: 100% FAIL** — extracted 0 references in all cases; Door/Window family geometry does not have perpendicular faces suitable for this approach.
- **Locked strategy for MVP:** `FamilyInstance.GetReferences(FamilyInstanceReferenceType.Left/Right)` as primary; `HostWallOpeningGeometry` as fallback.
- **Drop GeometryComputeReferences** for Door/Window — not viable for family instances.
- Filtering logic verified: parallel skip and outside-span skip work correctly.
- Session 1.4 objective achieved: two stable reference strategies proven for Door/Window openings.

#### Session 1.5 — Full mixed source spike
- Prove one chain can mix Grid, Wall, Door, and Window references.
- Validate a model with one wall, one door, one window, and at least one grid.
- Confirm that the engine can sort and hand the references to `NewDimension` in valid order.

Implementation record — 2026-06-01:
- Added temporary ribbon command `QuickDimensionFullMixedReferenceSpikeCommand` under Annotation Tools as `QD Full Mixed`.
- Created model file `QuickDimensionFullMixedReferenceProbe.cs` with unified candidate, test result, and summary types supporting all four source types.
- Created service file `QuickDimensionFullMixedReferenceProbeService.cs` merging proven strategies from Sessions 1.1-1.4.
- Reference strategies used:
  1. Grid: `new Reference(grid)` — Session 1.1 proven strategy.
  2. Wall: `HostObjectUtils.GetSideFaces()` with closest-face selection — Session 1.2 proven strategy.
  3. Door/Window: `FamilyInstance.GetReferences(FamilyInstanceReferenceType.Left/Right)` with `HostWallOpeningGeometry` fallback — Session 1.4 proven strategy.
- All candidates merged into unified list and sorted by `ParameterOnDimensionLine` before testing.
- Six test scenarios implemented:
  1. `FullMixed` — all four source types together (primary success criterion).
  2. `GridWall` — Grid + Wall only (Session 1.3 baseline).
  3. `WallOpening` — Wall + Door/Window only.
  4. `GridsOnly` — grids only baseline.
  5. `WallsOnly` — walls only baseline.
  6. `OpeningsOnly` — Door + Window only baseline.

Closure record — 2026-06-01:
- Runtime validation in Revit 2026: **ALL 7 TEST CASES PASS**.
- Test coverage: vertical, horizontal, and multiple diagonal dimension lines across complex wall/door/window layouts.
- Diagonal/slanted dimension lines (hardest cases) all passed Full Mixed test.
- Stress test: horizontal line accepted 53 references (7 grids, 6 walls, 12 doors, 28 windows).
- "Grids Only: FAIL" in some tests is expected behavior — dimension line only intersected 1 grid (need minimum 2).
- **Primary conclusion: Full mixed references (Grid + Wall + Door + Window) work in the same ReferenceArray.**
- **Phase 1 reference feasibility spikes are COMPLETE.**
- Known limitation confirmed: midpoint projection can miss elements whose midpoint falls outside the picked span even if they visually intersect the dimension line. Defer fix to Phase 2 geometry service with true line-line intersection.

Pass criteria for Phase 1:
- At least one stable reference strategy exists for each source type.
- Mixed reference chains are either proven valid or explicitly rejected with a new design rule.

### Phase 2 — Read-only engine

#### Session 2.1 — QD contract models
- Create immutable or near-immutable QD contract types.
- Define candidate, summary, diagnostic, and options models.

Closure record — 2026-06-04:
- Added production contract model file `ArcTool.Core/Models/QuickDimensionContract.cs`.
- Locked MVP source enum: Grid, Wall, Door, Window.
- Locked production reference strategy enum: Grid element reference, Wall side face, FamilyInstance Left/Right, and Host Wall opening geometry fallback.
- Added read-only contract types for options, line context, candidates, diagnostics, per-source summaries, and final read-only result.
- Contract layer stores `ElementId`, `Reference`, `XYZ`, and scalar diagnostics, but does not hold live `Element` objects and performs no transactions or `NewDimension` calls.
- Existing Phase 1 spike models/services were intentionally left unchanged; Session 2.2 should build geometry helpers against `QuickDimensionLineContext` instead of mutating spike code.
- Verification: structural file check passed (balanced braces, no TODO placeholders, correct namespace). User validated in the Windows/Revit developer environment: build succeeded, Revit add-in loaded without errors, and the existing `QD Full Mixed` spike command still ran without errors.

#### Session 2.2 — Geometry service
- Build math-only helpers for distance, sorting, dedupe, and line/segment validation.
- Keep this layer free of transactions.

Closure record — 2026-06-05:
- Added production geometry helper file `ArcTool.Core/Services/QuickDimensionGeometryService.cs`.
- The service is transaction-free and document-free: it performs finite-coordinate guards, straight-curve endpoint extraction, planar direction checks, near-parallel checks, point projection, dimension-line distance, true 2D segment/dimension-line intersection, stable sorting, and conservative source-aware dedupe.
- True segment intersection is the production replacement for the Phase 1 midpoint-projection limitation; future collectors must call this service instead of projecting element midpoints to decide picked-span hits.
- Dedupe intentionally preserves Door/Window left/right opening references by requiring matching source identity, host identity, display label, reference strategy, near-identical parameter, and near-identical hit point before removing a candidate.
- Verification: structural file check passed for balanced braces and required helper names; no `Document`, collector, transaction-opening, or `NewDimension` production call was introduced. Shell build remains environment-limited because `dotnet` is unavailable in the current Linux workspace.
- User validation: Windows developer-environment build succeeded, Revit loaded the add-in successfully, and rerunning `QD Full Mixed` produced no errors.

#### Session 2.3 — Grid candidate collector
- Collect Grid candidates from the active view.
- Reject arc grids in V1.
- Produce hit points, references, and diagnostics.

Closure record — 2026-06-05:
- Added production read-only collector `ArcTool.Core/Services/QuickDimensionGridCandidateCollector.cs`.
- Collector scope is Grid-only and plan-view-only; it returns `QuickDimensionReadOnlyResult` with candidates, diagnostics, and a Grid source summary.
- Candidate extraction uses `FilteredElementCollector(doc, view.Id).OfClass(typeof(Grid))`, rejects curved/arc grids, rejects parallel grids, and rejects grids outside the picked span with explicit diagnostics.
- Production hit detection uses `QuickDimensionGeometryService.TryIntersectSegmentWithDimensionLine2D()`; midpoint projection remains only in Phase 1 spike services.
- Grid references use the Phase 1 locked strategy `new Reference(grid)` and record `QuickDimensionReferenceStrategy.GridElementReference`.
- Verification: structural checks passed for balanced braces, required helper usage, no TODO placeholders, no transaction creation, no `ReferenceArray`, and no `NewDimension` call.
- User validation: Windows developer-environment build succeeded, Revit loaded the add-in successfully, and rerunning `QD Full Mixed` produced no errors.

#### Session 2.4 — Wall boundary collector
- Collect wall boundary candidates from the active view.
- Filter faces by normal and projection rules.
- Keep compound-wall behavior explicit.

Closure record — 2026-06-07:
- Added production read-only collector `ArcTool.Core/Services/QuickDimensionWallCandidateCollector.cs`.
- Collector scope is Wall-only and plan-view-only; it returns `QuickDimensionReadOnlyResult` with candidates, diagnostics, and a Wall source summary.
- Candidate extraction uses `FilteredElementCollector(doc, view.Id).OfClass(typeof(Wall)).OfCategory(BuiltInCategory.OST_Walls)`, rejects curtain walls, rejects arc/non-line walls, rejects parallel walls, and emits explicit diagnostics for unsupported or failed cases.
- Wall references use the Phase 1 locked strategy `HostObjectUtils.GetSideFaces()` for Exterior/Interior major side faces and record `QuickDimensionReferenceStrategy.WallSideFace`.
- Compound-wall behavior is explicit for MVP: only major Exterior/Interior boundary faces are considered; core/layer-level wall dimensioning is not exposed.
- Production hit detection builds a planar side-face segment from the wall location curve and resolved face centroid/normal, then uses `QuickDimensionGeometryService.TryIntersectSegmentWithDimensionLine2D()` for picked-span hits.
- Verification: structural checks passed for balanced braces, required helper usage, no TODO placeholders, no transaction creation, no `ReferenceArray`, and no `NewDimension` call. Shell build remains environment-limited because `dotnet` is unavailable in the current Linux workspace.
- User validation after local fixes: Windows developer-environment build succeeded, Revit loaded the add-in successfully, and rerunning `QD Full Mixed` produced no errors.

#### Session 2.5 — Door/Window opening collector
- Collect hosted opening candidates from Doors and Windows.
- Extract left/right opening edges as references.
- Record family/type/host diagnostics.

Implementation record — 2026-06-09:
- Added production read-only collector `ArcTool.Core/Services/QuickDimensionDoorWindowCandidateCollector.cs`.
- Collector scope is Door/Window-only and plan-view-only; it returns `QuickDimensionReadOnlyResult` with candidates, diagnostics, and separate Door/Window source summaries.
- Candidate extraction uses `FilteredElementCollector(doc, view.Id).OfCategory(OST_Doors/OST_Windows).OfClass(typeof(FamilyInstance))`, rejects non-wall-hosted or non-line-wall-hosted instances, and records family/type/host diagnostics.
- Reference strategy prioritizes `FamilyInstance.GetReferences(FamilyInstanceReferenceType.Left/Right)` and falls back to host-wall opening edge references from wall geometry when enabled.
- Production hit detection uses `QuickDimensionGeometryService.TryIntersectSegmentWithDimensionLine2D()` for picked-span hits instead of midpoint projection.
- Local structural checks passed for balanced braces and banned APIs. User validation on 2026-06-09: Windows developer-environment build succeeded, Revit loaded the add-in successfully, and rerunning `QD Full Mixed` produced no errors; dedicated production collector smoke remains pending until Session 2.6/integration invokes the Phase 2 collectors together.

#### Session 2.6 — Merge, sort, dedupe, and summarize
- Merge all source candidates into a single ordered sequence.
- Deduplicate cautiously with source-aware rules.
- Return a read-only summary only.

Implementation record — 2026-06-10:
- Added production read-only merge engine `ArcTool.Core/Services/QuickDimensionReadOnlyEngine.cs`.
- The engine calls the production Grid, Wall, and Door/Window collectors, merges their candidates, runs final source-aware sort/dedupe through `QuickDimensionGeometryService.DeduplicateCandidates()`, recomputes source summaries, and returns one `QuickDimensionReadOnlyResult`.
- Final summary semantics are explicit: `QuickDimensionReadOnlyResult` candidate counts remain final candidate records, while `QuickDimensionSourceSummary.AcceptedCount` is recomputed as distinct accepted source elements so Door/Window left/right edge records do not inflate accepted element counts.
- Added temporary read-only smoke command `ArcTool.Core/Commands/QuickDimensionReadOnlySummaryCommand.cs` and ribbon button `QD ReadOnly Summary`; it picks two plan-view points, invokes the production read-only engine, and shows ordered final candidates plus diagnostics without creating dimensions.
- Local structural checks passed for balanced braces/parentheses/brackets, no placeholder text, and no `Transaction`, `NewDimension`, or `ReferenceArray` usage in the read-only engine. Shell build remains environment-limited because `dotnet` is unavailable in the current Linux workspace.

Smoke-fix record — 2026-07-12:
- First real Revit 2026 smoke run of the wall-axis projection path confirmed selected-wall-only dispatch and Door/Window left/right jamb collection work: Window accepted count is no longer 0 and each hosted opening contributes both jamb records.
- The smoke-fix loop exposed that `WallEndFace` anchors were semantically wrong in two separate ways: opening reveal/jamb faces can share wall-axis normals, and `LocationCurve` endpoints can lie on joining-wall centerlines instead of the visible wall corner. Fix: wall-end candidates are wall-direction-aligned planar faces whose stations are computed directly with `QuickDimensionLineContext.ProjectParameter`; the min and max projected stations are selected as the two physical wall solid end caps, while opening jamb faces sit between those caps and are not used as wall anchors.
- The engine now performs global projected-station dedupe after conservative source-aware dedupe. Candidates sharing a station within duplicate tolerance are removed with `DuplicateStation` diagnostics because Phase 3 chain dimensions cannot use zero-length segments. `QuickDimensionReadOnlyResult.CanCreateChainDimension` now requires at least two final records and distinct projected stations for all final records.
- The read-only summary command now displays wall-axis length and ordered candidate `t` values in millimeters using `UnitUtils.ConvertFromInternalUnits(..., UnitTypeId.Millimeters)`, not Revit internal feet.
- Verification: source-level static checks and file-tail inspection passed; Linux shell still has no `dotnet`, so Windows/Revit build and re-smoke remain required before Phase 3.

Pass criteria for Phase 2:
- The engine can explain exactly what it found and what it ignored.
- The sorted candidate output matches the expected physical order.
- Door/Window candidates are handled as first-class source records.

### Phase 3 — Create dimension for real

#### Session 3.1 — Chain dimension creation
- Convert sorted references into a production `ReferenceArray`.
- Create the first working chain dimension in Revit.

#### Session 3.2 — DimensionType policy
- Decide how the feature chooses or receives a dimension type.
- Avoid hardcoded language-dependent names.

#### Session 3.3 — Total dimension creation
- Create a second total dimension using first and last references.
- Keep chain and total behavior deterministic.

#### Session 3.4 — Failure isolation and rollback
- Ensure all transaction failures roll back cleanly.
- Preserve document integrity on unsupported or invalid input.

#### Session 3.5 — Mixed-source stress test
- Validate Grid + Wall + Door + Window chains on denser layouts.
- Increase opening count progressively and watch for sort or tolerance drift.

Pass criteria for Phase 3:
- A real Revit dimension is created reliably.
- Failure leaves no dirty transaction state.

### Phase 4 — Hardening on real models

#### Session 4.1 — Clean-model acceptance test
- Use a controlled test model with known spacing.
- Compare output against expected segment values.

#### Session 4.2 — Wall + Door/Window complexity matrix
- Test empty walls, walls with openings, many openings, flipped instances, and close-spaced openings.
- Record support vs unsupported cases explicitly.

#### Session 4.3 — Grid complexity matrix
- Test straight, cropped, hidden, and arc grids.
- Ensure the MVP fails safely where support is not intended.

#### Session 4.4 — Performance test
- Measure collector and geometry extraction costs on larger models.
- Improve prefiltering if wall/opening extraction becomes slow.

#### Session 4.5 — Regression test against ArcTool
- Verify existing ArcTool commands still load and run.
- Confirm no namespace or startup regression was introduced.

Pass criteria for Phase 4:
- The feature works on real project-like content.
- It does not destabilize the rest of ArcTool.

### Phase 5 — Official integration

#### Session 5.1 — Ribbon integration
- Add the Quick Dimension command to the Annotation Tools panel.
- Keep the tooltip honest about current support.

#### Session 5.2 — Diagnostics and logging
- Add user-facing summary output.
- Add a dedicated QD log file for unsupported cases and failure analysis.

#### Session 5.3 — Documentation
- Update `CLAUDE.md` with a short pointer only.
- Create or update a dedicated dossier once implementation becomes stable.

Pass criteria for Phase 5:
- The feature is discoverable in the UI.
- The behavior is explainable and supportable.

### Phase 6 — Optional expansion

#### Session 6.1 — Column support
- Add rectangular column support only after the wall/opening pipeline is stable.

#### Session 6.2 — Minimal settings UI
- Add a compact dialog only if the engine proves stable enough to expose options.

#### Session 6.3 — Linked model research
- Treat linked models as a separate research branch, not part of the MVP.

#### Session 6.4 — Rubberband / preview R&D
- Research preview only after the core engine is reliable.

### Phase 7 — Release hardening

#### Session 7.1 — Release candidate regression
- Verify build, load, command startup, and QD behavior together.

#### Session 7.2 — User acceptance on real projects
- Validate the feature against the user's real production workflow.

## Status tracking block

Current phase: Phase 2.6 WALL-AXIS PROJECTION REWRITE IMPLEMENTED / PENDING WINDOWS-REVIT VALIDATION — read-only engine
Current session target: Build/load in Windows/Revit, then smoke `QD ReadOnly Summary` with select-one-wall + side-pick input; expected output is selected wall start/end anchors plus both jambs for each hosted Door/Window, ordered by wall-axis projection parameter.
Last verified against ArcTool code map: 2026-07-12
Session 1.1 closure note: `new Reference(grid)` PASS, `grid.Curve.Reference` FAIL in Revit 2026; production collector must use true line intersection instead of midpoint projection.
Session 1.2 closure note: `HostObjectUtils.GetSideFaces()` PASS consistently; locked as MVP strategy with closest-face selection rule. `GeometryComputeReferences` approach dropped for walls.
Session 1.3 closure note: Mixed Grid + Wall references PASS in same ReferenceArray; ReversedOrder also PASS (Revit auto-sorts). Midpoint projection inaccuracy confirmed — defer fix to Phase 2 geometry service.
Session 1.4 closure note: `FamilyInstance.GetReferences(Left/Right)` PASS 100%; `HostWallOpeningGeometry` PASS 100%; `GeometryComputeReferences` FAIL 100% (0 refs extracted). Locked primary strategy: FamilyInstanceReferences; fallback: HostWallOpeningGeometry. Drop GeometryComputeReferences for Door/Window.
Session 1.5 closure note: ALL 7 TEST CASES PASS in Revit 2026. Full mixed references (Grid + Wall + Door + Window) work in same ReferenceArray. Stress test: 53 refs accepted. Diagonal dimension lines (hardest cases) all passed. Phase 1 reference feasibility spikes COMPLETE.
Session 2.1 closure note: `QuickDimensionContract.cs` added immutable/near-immutable production contracts for options, line context, candidates, diagnostics, source summaries, and read-only result. User validated build success, clean Revit add-in load, and clean `QD Full Mixed` smoke run. Spike models remain untouched; next work is transaction-free geometry helpers.
Session 2.2 closure note: `QuickDimensionGeometryService.cs` added transaction-free/document-free helpers for finite guards, straight-curve endpoints, planar direction/parallel checks, projection, distance, true segment/dimension-line intersection, stable sorting, and conservative source-aware dedupe. User validated build success, clean Revit add-in load, and clean `QD Full Mixed` smoke run. Future collectors must use true intersection instead of midpoint projection.
Session 2.3 closure note: `QuickDimensionGridCandidateCollector.cs` added the production read-only Grid collector. It collects visible straight grids from the active plan view, rejects arc/parallel/outside-span grids with diagnostics, uses true 2D segment intersection for picked-span hits, and uses `new Reference(grid)` with `GridElementReference`. Structural checks passed; user validated Windows build success, clean Revit add-in load, and clean `QD Full Mixed` smoke run.
Session 2.4 closure note: `QuickDimensionWallCandidateCollector.cs` added the production read-only Wall collector. It collects visible straight non-curtain walls from the active plan view, uses `HostObjectUtils.GetSideFaces()` Exterior/Interior major side faces, builds side-face boundary segments, uses true 2D segment intersection for picked-span hits, and records `WallSideFace`. Structural checks passed; user validated Windows build success, clean Revit add-in load, and clean `QD Full Mixed` smoke run after local fixes.
Session 2.5 implementation note: `QuickDimensionDoorWindowCandidateCollector.cs` added the production read-only Door/Window collector. It collects visible wall-hosted Doors/Windows from the active plan view, prefers `FamilyInstance.GetReferences(Left/Right)`, falls back to host-wall opening edge references, uses true 2D segment intersection for picked-span hits, and records family/type/host diagnostics. Structural checks passed; user validated Windows build success, clean Revit add-in load, and clean `QD Full Mixed` regression run. Dedicated production collector smoke remains pending until merged/invoked by Session 2.6.
Session 2.6 projection rewrite note — 2026-07-12: production read-only main flow now uses the wall-axis projection model in code. `QuickDimensionReadOnlySummaryCommand` no longer asks for two picked points; it uses `Selection.PickObject(ObjectType.Element, ISelectionFilter, string)` to select one straight non-curtain host Wall, then `PickPoint(ObjectSnapTypes.None, string)` to capture placement side. `QuickDimensionLineContext.CreateFromWallAxis()` builds the axis from the selected Wall `LocationCurve` and rejects ambiguous side picks. `QuickDimensionReadOnlyEngine` routes wall-axis contexts to selected-wall-only collection. `QuickDimensionWallCandidateCollector.CollectSelectedWallEndAnchors()` creates wall start/end anchors from wall geometry planar end faces using `Options.ComputeReferences = true` because Revit 2026 `HostObjectUtils` has no `GetEndFaces()` method. `QuickDimensionDoorWindowCandidateCollector.CollectOpeningsAlongWallAxis()` collects only hosted Doors/Windows whose host id equals the selected wall id and projects both left/right jamb candidates onto the wall axis. `QuickDimensionReferenceStrategy.WallEndFace` records the new wall-end anchor strategy separately from the legacy side-face strategy. Local brace/forbidden-token checks passed; shell build remains unavailable because `dotnet` is not installed in the Linux workspace.

## Change policy

- Update this file after each meaningful session or phase.
- Keep `CLAUDE.md` short; only update it when the high-level scope or status cha
