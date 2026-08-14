# Quick Dimension — Session 2.7 Closure

Date: 2026-07-18; post-closure updates: 2026-07-19, 2026-07-20

## Scope

Session 2.7 closed the research/log-only phase for mixed L/T one-axis wall-joint aggregation.
It did NOT implement production aggregation, port the spike resolver, or create any `NewDimension`.
The Wall Spike remains open per ADR-2026-07-17C.

---

## What was done this session

### 1. Session 2.7 research report (Section 1–12 of prompt)
- Proposed the one-axis station-aggregation contract vocabulary, pipeline, classification tiers,
  and dedupe policy.
- Ran a full adversarial self-critique against 14 required counterexamples.
- Identified the core gap: `LocationCurve.get_ElementsAtJoin` and `JoinGeometryUtils` are blind to
  mid-run T-joints; the detection signal must come from vertical `Edge.Reference` evidence on the
  selected side line.
- Recorded no-code decision pending Section 11 experiment.

### 2. Section 11 log-only experiment — design, implement, smoke (Rounds 1 and 2)
- **New files added to the repo:**
  - `ArcTool.Core/Models/QuickDimensionWallMidRunProbe.cs`
    — `QuickDimensionWallMidRunRelation` enum, `QuickDimensionWallMidRunReferenceHit`,
      `QuickDimensionWallMidRunCandidate`, `QuickDimensionWallMidRunProbeResult`.
  - `ArcTool.Core/Services/QuickDimensionWallMidRunProbeService.cs`
    — Read-only view-scoped probe: collects all distinct vertical `Edge.Reference` hits per
      candidate wall, deduplicates by `Reference.ConvertToStableRepresentation`, classifies by
      reference evidence not by join API or perpendicularity.
  - `ArcTool.Core/Services/QuickDimensionWallSpikeXmlLogService.cs` — extended with `<MidRunProbe>`
    block, `<Candidates>/<ReferenceHits>/<ReferenceHit>` per-hit XML output.
- **Round 1 smoke findings (2026-07-17):** T-joint wall 381185 was `NonJoinedProximity` because
  old classifier required `inAnyJoin` and `isPerpendicular`; both signals were false for the oblique
  mid-run T. Log also collapsed two T-joint jambs to one midspan-best representative.
- **Fix:** classifier now uses `candidateReferenceNormalAlongAxis` only; log preserves ALL
  distinct side-line hits via stable-key dedupe. Both fixes are additive/read-only.
- **Round 2 smoke (2026-07-18) PASS:**
  - Wall 381185 → `MidRunCrossing`, `referenceHitCount=2`, stations 2665.787mm and 2867.289mm.
  - End-join wall 380858 coincident with Finish Anchor → reveals `DuplicateStation` requirement.
  - Parallel near wall 381035 → `ParallelNonJoined` (no hit). False-positive guard holds.
  - Left/Exterior 381185 → `Ignored`. Mid-run is shell-specific, not symmetric.

### 3. Section 10 acceptance gates — evaluated and closed for research phase
All research-phase gates satisfied; implementation gates are separate next-session work.

### 4. ADR updated
- **ADR-2026-07-18A** added: mid-run station detection uses side-line reference evidence, not join
  APIs. Join membership and perpendicularity are provenance/diagnostics only.

### 5. Roadmap and memory updated
- `.Dossier/Quick Dimension - Implementation Roadmap.md` — Session 2.7 block + Section 10
  closure note added to status tracking.
- `Memory/project_qd_midrun_smoke_evidence.md` — complete smoke evidence, root causes,
  acceptance gate results.
- `Memory/project_qd_lt_aggregation_research.md` — full research contract, counterexample matrix,
  no-code decision.
- `Memory/gemma4_error_learning_log.xml` — entry 2026-07-17-01 for Gemma contract-violation
  on enum naming.
- `MEMORY.md` index updated with two new entries.

---

## New durable knowledge locked this session

| Finding | Where persisted |
|---|---|
| `LocationCurve.get_ElementsAtJoin` only reports end joins; invisible to mid-run T | ADR-2026-07-18A, midrun smoke memory |
| `JoinGeometryUtils.GetJoinedElements` also returned nothing for mid-run T in tested fixture | ADR-2026-07-18A |
| Correct raw mid-run detection signal: vertical `Edge.Reference` on selected side line, normal along axis, station inside span | ADR-2026-07-18A |
| Accepted mid-run classification must exclude candidates in `ElementsAtJoinStart/End`, endpoint zones, and resolved anchor duplicates; require at least two distinct accepted stations | ADR-2026-07-18A refinement, roadmap, midrun smoke memory |
| Reference ownership for mid-run T-joint lives on the joining wall, not the selected wall | ADR-2026-07-18A, smoke memory |
| One mid-run T-joint contributes TWO jamb stations (~1 wall thickness apart) | Smoke memory |
| Mid-run detection is shell-specific: joining wall is visible on the shell it butts into only | Smoke memory, ADR |
| End-join wall can expose raw side-line references near anchors/endpoints and must be `EndJoinOnly`, not `MidRunCrossing` | Smoke memory, ADR refinement |
| `Reference.ConvertToStableRepresentation(doc)` is the correct deduplication key for same-edge hits | Service implementation |
| Gemma 4 (12B QAT) times out on large service-file generation (>~600 tokens output) via MCP | Gemma error log |
| ArcTool must use Visual Studio MSBuild for COM reference build verification; `dotnet build` fails on `ResolveComReference` | Visual Studio MSBuild memory |

---

## Post-closure updates

### 2026-07-19

Four additional real smoke sets were audited before closing this follow-up session. They showed that raw vertical side-line `ReferenceHit` evidence can come from endpoint join artifacts. The log-only classifier was fixed in source: `QuickDimensionWallMidRunProbeService` now uses resolved Start/Finish anchors from `QuickDimensionWallSpikeResult`, lets `EndJoinOnly` win before mid-run, excludes endpoint/anchor-zone hits, sorts/dedupes accepted stations, and requires `acceptedMidRunStationCount >= 2` for `MidRunCrossing`. `QuickDimensionWallSpikeXmlLogService` now passes anchors into the probe and emits `acceptedMidRunStationCount`. Visual Studio MSBuild build passed.

### 2026-07-20

Re-smoke of four real Wall Spike sets passed after the classifier fix. Wall 380815 accepted true mid-run 381185 only on Right/Interior; wall 379467 accepted true mid-run 379933 only on Right/Interior; wall 379469 accepted true mid-run 379933 only on Right/Interior; wall 379470 accepted true mid-run 380187 only on Right/Interior. Opposite/clean shells reported zero mid-run crossings, proximity-only candidates stayed ignored, and end-join artifacts stayed `EndJoinOnly` with `acceptedMidRunStationCount=0` even when raw `referenceHitCount > 0`. Wall Spike classifier is now cleared for production collector + read-only aggregator port; `NewDimension` remains gated.

## Remaining open work (priority order for next session)

1. **Port the Wall Spike side-face directional resolver** from `QuickDimensionWallReferenceProbeService`
   into production `QuickDimensionWallCandidateCollector.CollectSelectedWallEndAnchors`.
   Preserve live Revit `Reference` objects, source wall ids, stations, and diagnostics.
2. **Implement the one-axis wall-joint read-only aggregator** in production:
   — Collect both resolved end anchors (from ported resolver) and all accepted mid-run joining-wall
     reference hits inside the resolved anchor span.
   — Apply the ADR-2026-07-19A gate: exclude `ElementsAtJoinStart/End`, endpoint zones, and anchor duplicates; require `acceptedMidRunStationCount >= 2`.
   — Order ascending by projected station.
   — Deduplicate coincident anchor/joint/opening hits with explicit `DuplicateStation` diagnostics.
   — Reject geometry-only horizontal endpoints from the dimension-eligible tier.
   — Smoke on the same mixed fixture before accepting.
3. **Smoke wall-only read-only chains** on the four accepted re-smoke sets plus the original 12 L/T anchor set; compare final ordered stations and references against XML evidence.
4. **Smoke Door/Window separately** after wall-only aggregation passes.
5. **Smoke full mixed chain** (wall anchors + mid-run joints + door/window jambs) after each isolated stage passes.
6. Phase 3: `NewDimension` / `ReferenceArray` only after read-only full-chain is accepted and geometric-reference eligibility is verified.

---

## Files touched this session

### New source files
- `ArcTool.Core/Models/QuickDimensionWallMidRunProbe.cs`
- `ArcTool.Core/Services/QuickDimensionWallMidRunProbeService.cs`

### Modified source files
- `ArcTool.Core/Models/QuickDimensionWallMidRunProbe.cs` — `AcceptedMidRunStationCount` added to separate raw hits from accepted stations
- `ArcTool.Core/Services/QuickDimensionWallMidRunProbeService.cs` — accepted-mid-run filtering now excludes end-join artifacts, endpoint zones, and anchor duplicates
- `ArcTool.Core/Services/QuickDimensionWallSpikeXmlLogService.cs` — `<MidRunProbe>` block added; later updated to pass resolved anchors and emit `acceptedMidRunStationCount`

### Documentation / memory
- `.Dossier/Quick Dimension - Implementation Roadmap.md` — Session 2.7 + Section 10 updates
- `.Dossier/Quick Dimension - Session 2.7 Closure.md` — this file (new)
- `.codebase-memory/adr.md` — ADR-2026-07-18A and ADR-2026-07-19A added
- `Memory/project_qd_midrun_smoke_evidence.md` — new
- `Memory/project_qd_lt_aggregation_research.md` — new
- `Memory/gemma4_error_learning_log.xml` — entry added
- `Memory/MEMORY.md` — two entries added

---

## Constraints that must not be violated in future sessions

- Wall Spike stays OPEN until read-only full-chain aggregation passes smoke.
- Production `CollectSelectedWallEndAnchors` must not use join-role heuristics or `LocationCurve`
  endpoints directly; use the side-face directional resolver from the spike.
- `NewDimension`/`ReferenceArray` code only after read-only engine is accepted.
- Mid-run reference ownership: collect from joining wall, never assume selected wall has the ref.
- One T-joint contributes two jambs; do not dedupe to one midspan representative.
- `DuplicateStation` guard required at anchor/joint collision points.
- Side-line reference evidence (`candidateReferenceNormalAlongAxis`) is the gate, not join APIs.
- Product scope: one selected straight non-curtain wall → one reviewable chain (ADR-2026-07-17C).
- Never add bulk/automatic multi-wall dimension creation.
- `dotnet` unavailable in Linux workspace; all build/smoke in Windows/Revit dev environment.
- Chief Architect = Claude; code generation = Gemma 4 via MCP; Claude reviews before applying.
