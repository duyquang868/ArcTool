# Quick Dimension — Implementation Roadmap

RETIRED STATUS — 2026-08-10
- Operator EV-4 concluded Quick Dimension is no longer feasible/appropriate to continue developing.
- Live status changed from active roadmap to retired/archived on 2026-08-10.
- Source archive path: `ArcTool.Core/Archive/QuickDimension/{Commands,Models,Services}/`.
- Archived source is preserved in-repo but excluded from compilation via `<Compile Remove="Archive\QuickDimension\**\*.cs" />` in `ArcTool.Core/ArcTool.Core.csproj`.
- No further Quick Dimension roadmap phases are planned unless a future operator explicitly revives the feature.
- Closure record: `.Dossier/Quick Dimension - Retirement Record.md`.

Last updated: 2026-08-10
Status: Retired/archived historical roadmap for the Quick Dimension feature in ArcTool.

## Purpose

This document is the long-form source of truth for the Quick Dimension feature. It exists so the roadmap can survive across many sessions without bloating `CLAUDE.md` or requiring repeated high-noise edits to the root operating document.

## Scope lock

The current MVP scope is:
- Active Revit Plan View only.
- Main-flow input: select one straight non-curtain host Wall and pick a placement side; the old two-picked-points input is retained only as a deprecated/optional cross-cutting path.
- Main-flow sources: selected wall end anchors plus hosted Door and Window openings in that wall; Grid and non-selected-wall sources are disabled in the wall-axis projection dispatch.
- Output: one reviewable chain dimension and optional total dimension for that selected wall only.
- One invocation always handles one selected wall axis; no bulk or automatic multi-wall dimension creation.
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
- One selected straight wall per invocation, one reviewable chain (ADR-2026-07-17C). Bulk/automatic multi-wall dimension creation is forbidden because a human cannot validate high-volume geometry-dependent output.
- Per-joint left/right anchor correctness is an input invariant only. A correct continuous chain over a mixed L-joint/T-joint wall axis must still be proven separately by research and self-critique before code.
- Before writing or porting aggregation code, run a research-then-self-critique loop against mixed-joint counterexamples (L-L, T-T, L-T, T-L, reversed-axis, coincident stations). Code is written only after the model survives that critique.

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
- The smoke-fix loop exposed that `WallEndFace` anchors were semantically wrong in two separate ways: opening reveal/jamb faces can share wall-axis normals, and `LocationCurve` endpoints can lie on joining-wall centerlines instead of the visible wall corner. Fix (later superseded — see Wall Spike reset 2026-07-14): wall-end candidates are wall-direction-aligned planar faces whose stations are computed directly with `QuickDimensionLineContext.ProjectParameter`; the min and max projected stations are selected as the two physical wall solid end caps, while opening jamb faces sit between those caps and are not used as wall anchors.
- The engine now performs global projected-station dedupe after conservative source-aware dedupe. Candidates sharing a station within duplicate tolerance are removed with `DuplicateStation` diagnostics because Phase 3 chain dimensions cannot use zero-length segments. `QuickDimensionReadOnlyResult.CanCreateChainDimension` now requires at least two final records and distinct projected stations for all final records.
- The read-only summary command now displays wall-axis length and ordered candidate `t` values in millimeters using `UnitUtils.ConvertFromInternalUnits(..., UnitTypeId.Millimeters)`, not Revit internal feet.
- Verification: source-level static checks and file-tail inspection passed; Linux shell still has no `dotnet`, so Windows/Revit build and re-smoke remain required before Phase 3.

Wall Spike reset — 2026-07-14:
- User directive: "test each logic first before mixing". Wall Spike (`QuickDimensionWallReferenceSpikeCommand` + `QuickDimensionWallReferenceProbeService`) was fully rewritten to isolate wall-end anchor extraction from Grid/Door/Window sources.
- Iteration 1 (side-face vertical edges via `HostObjectUtils.GetSideFaces()`): failed — `Vertical edges on side face: 0`. Route rejected.
- Iteration 2 (all wall-geometry side faces aligned with side normal, vertical edges min/max): Left OK, Right lands on wall-join solid extension instead of the plan-visible corner.
- Iteration 3 (largest side face by Area only): same Right-side lean; largest face still owns the join extension.
- Iteration 4 (current, pending re-smoke): use `wall.get_Geometry(ComputeReferences=true)` → largest horizontal top `PlanarFace` → outer-loop footprint vertices on picked side (offset filter `max(50 mm, Width*0.35)` against `targetSideNormal`) → min/max station along wall axis to select the two plan-visible corners; each corner XY is mapped to the nearest vertical solid `Edge.Reference` within `max(50 mm, Width/2)`. Static checks pass; Revit build + Right-side re-smoke is the next required step. See `Memory/project_qd_wall_spike_handoff.md` for the ordered follow-up list, including fallbacks if Right still lags.
- Production `QuickDimensionWallCandidateCollector.CollectSelectedWallEndAnchors` and ADR-2026-07-12 still describe the older wall-direction planar-face min/max rule (from 2026-07-12) and must not be reconciled until the spike is confirmed on both sides.

Wall Spike isolated closure + scope clarification — 2026-07-17:
- User smoke confirmed 100% PASS on walls 379467, 379468, 379469, 379470, 379933, 380187 Left/Right (12 XML logs). The side-face boundary + directional full-height resolver (ADR-2026-07-17B) is the accepted spike model.
- User clarified the durable product intent (ADR-2026-07-17C): Quick Dimension always dimensions ONE operator-selected straight wall at a time and produces ONE reviewable chain. Bulk/automatic multi-wall dimensioning is rejected on purpose because a human cannot validate ~100 geometry-dependent dimensions for correctness.
- Therefore the two proven L-joint and T-joint anchor logics are only the foundation. The open problem is combining them: along one selected wall axis a real building mixes L-joint and T-joint corners (and mid-run T-joints from intersecting walls), so the chain must aggregate the two end anchors plus every relevant intermediate joint station in ascending order without duplicates.

Session 2.7 — Mixed L/T one-axis aggregation research (ACTIVE, research-before-code):
- Goal: define and prove the one-axis station-aggregation contract for one selected straight wall before writing or porting any production code.
- Method is a research-then-self-critique loop: propose the aggregation model, then attack it with counterexamples (L-L ends, T-T ends, L-T, T-L, mid-run T-joint, reversed pick direction, coincident/near-coincident stations) before accepting it.
- Output of this session is a validated model + diagnostics plan, not source edits. Code (spike or production port) begins only after the model survives self-critique.
- Session 2.7 Section 11 log-only experiment is complete after Windows/Revit smoke on the mixed fixture. Evidence: join APIs are blind to the oblique mid-run T-wall 381185; both real jamb references appeared only after logging all side-line vertical `Edge.Reference` hits; Right/Interior captured two `MidRunCrossing` hits at 2665.787mm and 2867.289mm; Left/Exterior correctly had no mid-run hits; parallel near wall 381035 remained `ParallelNonJoined`; end-join wall 380858 produced a coincident hit at the Finish Anchor station, proving the future aggregator needs `DuplicateStation` handling at anchors.
- Post-closure audit on 2026-07-19 across four real smoke sets found a log-classification defect: endpoint join walls can expose raw vertical side-line references and were overcounted as `MidRunCrossing`. Fix applied in `QuickDimensionWallMidRunProbeService`: `EndJoinOnly` wins before mid-run; accepted mid-run stations exclude `ElementsAtJoinStart/End`, endpoint zones, and resolved anchor duplicates, require `candidateReferenceNormalAlongAxis=true`, sort/dedupe by station, and require `acceptedMidRunStationCount >= 2`. `QuickDimensionWallSpikeXmlLogService` now passes resolved anchors into the probe and logs `acceptedMidRunStationCount`. VS MSBuild build passed; Revit re-smoke of the same four sets is pending.
- Section 10 acceptance gates are satisfied for the research/log-only contract; implementation gates remain separate: re-smoke the refined log classifier, port the side-face directional end-anchor resolver into production first, then implement the reference-preserving one-axis wall-joint aggregator read-only before any `NewDimension` work.
- Constraint: do not merge Wall + Door/Window + Grid until each isolated logic passes; keep the XML diagnostic format comparable to prior Wall Spike smoke reports.

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

Runtime evidence record — 2026-08-01:
- **Smoke #1, wall 380815: PASS creation and audit.** Left/Exterior dim 383577: 2 refs, 4952.493mm. Right/Interior dim 383578: 4 refs, 2608.052/201.502/2128.527mm; mid-run T-wall 381185 resolved correctly. All three audit gates true on both shells. Cosmetic logger issue: two-reference dim reports raw `NumberOfSegments=0` / `Segments.Size=0`, but `Dimension.Value` validates one correct segment; `actualSegmentCount` must use validated value count.
- **Smoke #2, wall 379467: PASS creation/geometry; audit is an output-level false negative but a valid BUG-11 invariant failure.** Left/Exterior dim 383579: 10 refs, 9 segments 966.965/406/539.5/915/341.378/406/1037.622/915/925.03mm. Right/Interior dim 383580: 12 refs, 11 segments 827.035/406/539.5/915/341.378/406/300/201.502/536.12/915/759.97mm; mid-run T-wall 379933 resolved correctly. Screenshot labels match all rounded values exactly.
- **Smoke #3, wall 379469: PASS creation/geometry; BUG-11 broadened.** Left/Exterior dim 384631: 10 refs, 9 segments 1501.707/500/1092.5/915/930.838/500/554.162/915/2493.83mm. Right/Interior dim 384632: 12 refs, 11 segments 1196.398/500/1092.5/915/316.876/200.963/413/500/554.162/915/2266.043mm; mid-run T-wall 379933 resolved correctly only on Interior/Right. Both screenshots match every rounded segment exactly.
- Smokes #2/#3 report `referenceOrderRelation=Mismatch` and all three gates false because local opening pairs commit in a different stable-reference sequence than their collector station order. Smoke #3 expands the affected set from one Window to Window 379475 and Doors 379472/379471 on both shells; Window 379484 remains unswapped despite sharing the same family/type as 379475. The defect is therefore per-instance, not a global per-family rule. Flip/mirror/orientation is the leading hypothesis but is not yet proven because the XML does not log those flags. Do not accept `LocalPairSwap`; prove each reference's physical station directly.
- Smoke #2 confirms BUG-10 as metadata-only on door 379481 fallback: candidate `elementId=379481`, live stable-reference owner=host wall 379467. Smoke #3 does not exercise fallback; every candidate and committed reference owner matches its declared element metadata.
- Required next implementation boundary: prove each named FamilyInstance reference's physical projected station directly from reference geometry, correct the collector reference↔station association, then rerun walls 379467 and 379469 both shells before changing the audit order gate. Do not auto-whitelist local pair swaps. Keep `CreateChainDimension`, wall aggregation, and reference geometry unchanged unless a concrete runtime defect appears.
- **Smoke #4, wall 379470: PASS creation/geometry on both shells; BUG-11 confirmed, BUG-10 confirmed.** Left/Exterior dim 384894: 8 refs, 7 segments 805.802/406/500/915/255.938/406/1150.527mm (sum=4439.267mm=resolved span; anchors extend to joined walls 379469/379467). Right/Interior dim 384895: 10 refs, 9 segments 578.015/406/500/915/255.938/406/366.949/206.084/437.563mm (sum≈4071.549mm; mid-run wall 380187 at 206.084mm on Interior only; anchors snap inward to selected wall 379470). Screenshots match all rounded values exactly in both shells. Committed mappings `[1,3,2,4,5,7,6,8]` / `[1,3,2,4,5,7,6,8,9,10]`: Windows 379479 and 379478 (both `M_Fixed 0406 x 0610mm`) swap on both shells; Door 379482 fallback pair stays in order. BUG-11 per-instance defect confirmed across all four candidate walls; BUG-10 on Door 379482 fallback metadata (candidate `elementId=379482`, live owner=host wall 379470). Audit `Mismatch` cascade is not evidence of wrong geometry.

Pass criteria for Phase 3:
- A real Revit dimension is created reliably.
- Failure leaves no dirty transaction state.

Resolution of the 2026-08-01 evidence record (added 2026-08-04): the collector reference↔station correction landed, and the EV-2 re-smokes on walls 379467/379469/379470 (both shells) all report `Exact` order with every audit gate true. The 2026-08-01 block above is preserved as history only — its "required next implementation boundary" line is no longer the live next step. The "no dirty transaction state" half of the Phase 3 pass criteria is not yet runtime-proven and is tracked by the standalone task in `.Dossier/Quick Dimension - Deferred Rollback Validation Task.md`.

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
- 2026-08-05 packaging note: the full Phase 4 hardening work package is scaffolded at `.claude/workpackages/quick-dimension-phase4-hardening/`; status = `NOT STARTED`, scaffold complete through `T7.2`, first dispatch target = `T1.1`, and no operator evidence has been requested yet.

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

Current phase: Phase 3 `NewDimension` RUNTIME AUDIT — **BUG-10/BUG-11 MISSION CLOSED 2026-08-04; EV-2 PASS; EV-3 REOPEN PASS; FORCED ROLLBACK DEFERRED**. The BUG-11 collector fix landed: named `FamilyInstanceReferenceType.Left/Right` references now derive projected station from that same reference geometry, keeping identity+station atomic. The BUG-10 fallback metadata fix also landed: fallback candidates now align `elementId` with the live reference owner while preserving `hostElementId` as the selected wall. ChainCreationAudit logger fixes landed too: `actualSegmentCount` now uses the normalized measured-value count, and each `<Segment>` now records `valueSource`. Locked VS MSBuild produced the regression candidate DLL on 2026-08-03. EV-2 ran on that rebuilt DLL for walls 379467/379469/379470, both shells (six runs): every run committed with `Exact` audit order, `referenceIdentityMatched`/`referenceOwnersMatched`/`segmentValuesMatched` all true, and unchanged geometry. EV-3 reopen persistence passed on dimensions 385355/385356/385632/385584/385719/385720. Forced-rollback validation was removed from this mission's gates by operator decision on 2026-08-04 and now lives as an independent future task in `.Dossier/Quick Dimension - Deferred Rollback Validation Task.md`; deferral is safe because no task in the work package edited `QuickDimensionChainCreationService.cs`, so rollback behavior is byte-identical to the pre-fix build. Claude must not run Revit/MCP/smoke without explicit request. Authoritative handoff: `Memory/project_qd_chain_creation_audit_handoff.md`.

### Next-session handoff prompt (persisted 2026-08-04; BUG-10/BUG-11 mission closed)
- Read first, do not re-derive: `CLAUDE.md`; `Memory/project_qd_chain_creation_audit_handoff.md` (authoritative current handoff); `Memory/feedback_revit_runtime_operator_control_and_journal_analysis.md`; then this roadmap. Read package state from `.claude/quick-dimension-bugfix/04_EVIDENCE_QUEUE.md`, `.claude/quick-dimension-bugfix/06_EXECUTION_STATE.md`, and `.claude/quick-dimension-bugfix/HANDOFF_TO_NEXT_SESSION.md`.
- Source/build state: `QuickDimensionDoorWindowCandidateCollector.cs` derives each named `FamilyInstanceReferenceType.Left/Right` station from that same reference geometry and aligns fallback candidate `elementId` with the live reference owner; `QuickDimensionReadOnlyXmlLogService.cs` reports normalized `actualSegmentCount` and per-segment `valueSource`; locked VS MSBuild passed and produced the DLL used by EV-2.
- Runtime state: EV-2 (six runs) and EV-3 (reopen persistence) are both SUPPLIED and PASS. No further runtime evidence is owed by this mission.
- Do not reopen BUG-10 or BUG-11 without a new runtime defect. Both are runtime-confirmed fixed on the rebuilt DLL; treat any new local swap as a fresh regression with its own evidence, not as a reason to whitelist `LocalPairSwap`.
- Forced-rollback validation is a separate, not-started task. Entry point: `.Dossier/Quick Dimension - Deferred Rollback Validation Task.md`. It needs a fault-injection harness or debug-only switch that forces a post-`Transaction.Start()` failure; every operator-reachable invalid input currently returns `Result.Cancelled` before the transaction starts, so it cannot be requested as a plain smoke.
- Non-goals remain unchanged: no audit weakening, no collector-wide rewrite beyond the landed reference↔station association, no ribbon/UX redesign, no Grid/linked-model/column/arc expansion, and no bulk multi-wall creation.
- This 2026-08-04 update persisted closure state only. No Revit runtime action, Revit MCP action, or codebase-memory re-index was performed by Claude.
Last verified against ArcTool code map: 2026-08-04 (BUG-10/BUG-11 mission closed; EV-2 + EV-3 reopen PASS; forced rollback split into a standalone future task)
Session 2.7 closure note (2026-07-18): Research report produced, Section 11 log-only mid-run probe implemented and smoked (2 rounds), Section 10 acceptance gates satisfied for research/log contract. New locked knowledge: join APIs blind to mid-run T-joints; correct signal = vertical `Edge.Reference` on selected side line with normal along axis; reference ownership lives on joining wall; one T-joint contributes two jamb stations; mid-run detection is shell-specific; end-join collision at anchor station needs `DuplicateStation` guard. ADR-2026-07-18A added. Two new log-only files: `QuickDimensionWallMidRunProbe.cs` (Models) and `QuickDimensionWallMidRunProbeService.cs` (Services). Production `CollectSelectedWallEndAnchors` still uses superseded min/max model — porting it is Gate 1 after re-smoke.
Session 2.7 post-closure re-smoke note (2026-07-20): four real Wall Spike smoke sets passed after the accepted-mid-run classifier fix. Sets: 380815 (true mid-run 381185 on Right/Interior), 379467 (true mid-run 379933 on Right/Interior), 379469 (true mid-run 379933 on Right/Interior), and 379470 (true mid-run 380187 on Right/Interior). Opposite/clean shells had `mid-run crossings: 0`; end-join candidates stayed `EndJoinOnly` with `acceptedMidRunStationCount=0` even when raw `referenceHitCount > 0`. Wall Spike classifier is cleared for production collector + read-only aggregator port. `NewDimension` remains blocked until reference-preserving read-only aggregation passes.
Production read-only re-smoke note (2026-07-22): four real production `QuickDimensionReadOnlySummaryCommand` smoke sets passed after BUG-09 metadata fix. XML/image pairs: 380815, 379467, 379469, 379470, both shells. All visible survey labels matched `FinalCandidates` coordinates; all final anchors/mid-run candidates carried stable refs with blank `stableReferenceError`; `FinalCandidate.elementId` matched the stable-reference owner while `hostElementId` preserved the selected wall. Classifier stayed PASS: 380815→381185, 379467→379933, 379469→379933, 379470→380187 only on Interior/Right; opposite shells clean; end joins `EndJoinOnly`. `NewDimension` remains gated by: using resolved final candidate span for dimension line construction, correcting XML Grid-options metadata, and verifying close-spaced opening semantics from test 379470 where Window/Door/Window references interleave by owner.
Production read-only smoke #3 note (2026-07-27): selected wall 379469, both shells, annotated survey images + XML (`ArcTool_QD_ReadOnlySummary_379469_Left_20260727_170819.xml`, `..._Right_20260727_170743.xml`). Read-only PASS, no new defect. Interior/Right chain `2266,915,601,406,460,201,317,915,1140,406,1243` (span 8869.941mm) and Exterior/Left chain `2494,915,601,406,978,915,1140,406,1549` (span 9403.038mm) match the visible Revit labels exactly; adjacent deltas sum to the anchor span in both shells. Crossing wall 379933 stayed `MidRunCrossing` (`acceptedMidRunStationCount=2`, stations 4173.428/4374.391, ≈200.963mm apart) only on Interior/Right and `Ignored` on Exterior/Left; end joins 379468/379470 stayed `EndJoinOnly`. BUG-09 owner invariant held (every stable-reference owner token matched `elementId`; all `hostElementId=379469`); BUG-10 did not manifest — all four openings resolved via `FamilyInstanceLeftRight` (doors 915mm, windows 406mm), no `HostWallOpeningGeometry` fallback. This run's XML shows `includeGrids="false"` with matching "Grid collection is disabled" diagnostic (the 2026-07-22 Grid-options mismatch is absent here). Re-read `QuickDimensionChainCreationService.CreateChainDimension` (no edit): dimension line is built planar from final-candidate min/max stations via `LineContext.Evaluate`, so the elevation split (wall-edge LINEAR at 4000mm vs opening SURFACE at 0) does not corrupt line geometry, and a post-create `dimension.References.Size` rollback gate exists. `NewDimension` remains OPEN: `canCreateChainDimension` proves only distinct stations, not that Revit accepts the mixed LINEAR/SURFACE `ReferenceArray`; a real `NewDimension` creation/commit smoke on 379469 (Interior/Right preferred) is still required before Phase 3. No source changed this session. See [[project_qd_midrun_smoke_evidence]] for full detail.
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
- Keep `CLAUDE.md` short; only update it when the high-level scope or status changes.
- Keep long-form phase/session detail here, not in the root technical context file.
