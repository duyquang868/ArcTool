# Quick Dimension - Session 2.7 Closure Dossier

Last updated: 2026-07-20

## Scope

Session 2.7 closed the research/log phase for mixed L/T one-axis wall-joint aggregation. The work stayed intentionally limited to research, diagnostic logging, XML evidence, ADR updates, and roadmap/memory persistence. No production collector was ported, no dimension was created, no ReferenceArray was built.

## What was completed

### 1. Research report (Section 10 research/log phase)

The one-axis station-aggregation research report was produced from the supplied prompt. Conclusion: aggregation was conditionally plausible but not supportable until mid-run T-joint reference evidence was tested in Revit. The report defined station vocabulary, a pipeline model, 14-case counterexample matrix, and acceptance gates. ADR-2026-07-18A was added as part of this phase.

### 2. Section 11 log-only probe — added source files

Two new diagnostic-only files were added:

- `ArcTool.Core/Models/QuickDimensionWallMidRunProbe.cs` — value-only models: `QuickDimensionWallMidRunRelation` enum, `QuickDimensionWallMidRunReferenceHit`, `QuickDimensionWallMidRunCandidate`, `QuickDimensionWallMidRunProbeResult`. No live Element/Reference retained.
- `ArcTool.Core/Services/QuickDimensionWallMidRunProbeService.cs` — static read-only probe service. Scans view-visible straight non-curtain walls for vertical `Edge.Reference` hits on the selected side line; records join provenance from `ElementsAtJoin(0/1)` and `JoinGeometryUtils` separately; classifies candidates; calls no Transaction/NewDimension/ReferenceArray.

`QuickDimensionWallSpikeXmlLogService.cs` was extended with a `<MidRunProbe>` block appended after `<JoinedWalls>`, using `doc.ActiveView`. Existing `<ProbeResult>`, `<SelectedWall>`, `<Corners>`, `<JoinedWalls>` blocks unchanged.

`QuickDimensionWallReferenceProbeService.RunWallReferenceProbe` was NOT modified.

### 3. Two smoke rounds

**Round 1 (2026-07-17):** Mid-run T-wall 381185 was classified `NonJoinedProximity` and reported only one jamb. Root causes: (a) the classifier required `inAnyJoin` and `isPerpendicular`, both false for an oblique T-joint invisible to all join APIs; (b) the probe selected one midspan-best hit, hiding the second jamb.

**Round 2 (2026-07-18):** After correcting the classifier and switching to all-hits-per-candidate logging, wall 381185 produced `MidRunCrossing` with two `ReferenceHit` stations: 2665.787mm (N 20417.016 / E 20116.839) and 2867.289mm (N 20502.175 / E 20299.461), both `candidateReferenceNormalAlongAxis=true`. Parallel non-join wall 381035 remained `ParallelNonJoined`. Shell specificity confirmed: Left/Exterior had no mid-run hit. End-join wall 380858 produced a coincident hit at the Finish Anchor station (DuplicateStation case).

### 4. Section 10 acceptance gates satisfied for research/log contract

All research/log phase gates passed: mixed fixture smoked, correct signal identified, false-positive guard held, shell-specific behavior understood, operator accepted the result. Production implementation gates remain separate.

### 5. ADR and memory records updated

- `ADR-2026-07-18A` added to `.codebase-memory/adr.md` via `manage_adr`.
- `Memory/project_qd_lt_aggregation_research.md` — research contract.
- `Memory/project_qd_midrun_smoke_evidence.md` — full smoke evidence including both rounds.
- `Memory/MEMORY.md` — index updated with both new entries.
- `Memory/gemma4_error_learning_log.xml` — one correction entry added (Gemma enum contract violation during log-only probe generation).

## New locked knowledge (ADR-2026-07-18A)

- `LocationCurve.get_ElementsAtJoin` and `JoinGeometryUtils.GetJoinedElements` are blind to mid-run T-joints; they only report joins at the two selected-wall curve ENDS.
- The correct detection signal for a mid-run wall-joint station is: a vertical `Edge.Reference` on the selected side line (`distanceToSideLine <= 5mm`) with the adjacent face normal along the selected wall axis (`candidateReferenceNormalAlongAxis=true`), at a station strictly inside the axis span. Join-set membership and perpendicularity are provenance only.
- Reference ownership for mid-run T lives on the joining wall, not the selected wall (`selectedWallExposesRefAtStation=false`, `candidateWallExposesRefAtStation=true` in this fixture).
- One mid-run T-joint contributes TWO jamb stations (~200mm apart for a 200mm wall). Logging/aggregation must preserve all distinct side-line reference hits.
- Mid-run detection is shell-specific: the T-joint only exposed hits on the shell the joining wall butts into (Interior in this fixture). Reversed side pick is NOT symmetric for mid-run stations.
- An end-join wall whose reference lands on the same station as an end anchor is a DuplicateStation case and must be filtered by the DuplicateStation guard in the aggregator, not by the MidRunCrossing classifier.

## Post-closure follow-up — 2026-07-19 / 2026-07-20

Four additional real smoke sets were reviewed against the log-only mid-run classifier. They found a smaller but production-relevant defect: endpoint join walls can expose raw vertical side-line `ReferenceHit` evidence and were being overcounted as `MidRunCrossing` because the old classifier checked candidate reference-normal evidence before `ElementsAtJoinStart/End`.

Applied source fix:
- `QuickDimensionWallMidRunProbeService.Probe` now receives the resolved `QuickDimensionWallSpikeResult` and uses Start/Finish anchors as the trusted accepted span.
- `ProbeCandidate` separates raw `ReferenceHits` from accepted mid-run stations.
- `Classify` now returns `EndJoinOnly` before considering `MidRunCrossing`.
- Accepted mid-run stations require `candidateReferenceNormalAlongAxis=true`, candidate not in start/end join sets, station strictly inside resolved anchor span, not near Start/Finish anchors, not near raw location-curve endpoint zones, sorted and deduped by `StationEps`.
- `MidRunCrossing` now requires `acceptedMidRunStationCount >= 2`.
- `QuickDimensionWallSpikeXmlLogService` emits `acceptedMidRunStationCount` in each `<Candidate>`.

Verification: `dotnet build ArcTool.slnx --no-restore` still fails on .NET Core MSBuild `ResolveComReference`; Visual Studio MSBuild 18 with dash switches passed. Revit re-smoke of the four audited sets passed on 2026-07-20: selected walls 380815, 379467, 379469, and 379470 kept true mid-run walls accepted only on the correct shell while opposite/clean shells and end-join artifacts stayed clean.

## Open items / next session

In priority order:

1. **Port spike end-anchor resolver to production** — `QuickDimensionWallCandidateCollector.CollectSelectedWallEndAnchors` still uses the superseded min/max planar-face model (ADR-2026-07-12). Port the side-face directional resolver from `QuickDimensionWallReferenceProbeService` without changing Door/Window/Grid logic.
2. **Implement read-only one-axis wall-joint aggregator** — collect the two end anchors (from the ported resolver) plus accepted side-line reference hits from joining walls strictly inside the resolved anchor span; preserve live `Reference`, source wall id, station, hit point, side/shell, and diagnostics; order by projected station; apply `DuplicateStation` guard for coincident anchor/joint/opening hits; emit a read-only `QuickDimensionReadOnlyResult` before any `NewDimension` call.
3. **Apply the production acceptance gate from ADR-2026-07-19A** — exclude `ElementsAtJoinStart/End`, endpoint zones, and anchor duplicates; require `candidateReferenceNormalAlongAxis=true`; require at least two distinct accepted stations per true mid-run wall; treat raw `ReferenceHits` as diagnostic evidence only.
4. **Smoke wall-only read-only chains** on the four 2026-07-20 re-smoked sets plus the original 12 L/T anchor cases; verify ascending station order, two end anchors, true mid-run jamb pairs, explicit duplicate diagnostics, and no false-positive from parallel/proximity-only walls.
5. **Then smoke Door/Window isolated and full mixed chains.** Phase 3 `NewDimension` work begins only after read-only full-chain aggregation is accepted and geometric-reference eligibility is verified.

## Files changed this session

| File | Change |
|---|---|
| `ArcTool.Core/Models/QuickDimensionWallMidRunProbe.cs` | NEW + UPDATED — log-only value models; added `AcceptedMidRunStationCount` |
| `ArcTool.Core/Services/QuickDimensionWallMidRunProbeService.cs` | NEW + UPDATED — log-only probe service; accepted mid-run filtering now excludes end-join artifacts and endpoint zones |
| `ArcTool.Core/Services/QuickDimensionWallSpikeXmlLogService.cs` | MODIFIED — `<MidRunProbe>` block + `RevitView` alias; now passes resolved anchors and emits `acceptedMidRunStationCount` |
| `.codebase-memory/adr.md` | MODIFIED — ADR-2026-07-18A and ADR-2026-07-19A added |
| `.Dossier/Quick Dimension - Implementation Roadmap.md` | MODIFIED — Session 2.7 status + Section 10/11 results |
| `Memory/project_qd_lt_aggregation_research.md` | NEW |
| `Memory/project_qd_midrun_smoke_evidence.md` | NEW |
| `Memory/MEMORY.md` | MODIFIED — two new index entries |
| `Memory/gemma4_error_learning_log.xml` | MODIFIED — one correction entry |
| `CLAUDE.md` | MODIFIED (already updated at session start per header) |
