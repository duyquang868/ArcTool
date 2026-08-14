# ADR store — verbatim recovery of entries lost to full-replace overwrites

Do not treat this file as the live ADR store. It is a byte-faithful archive of ADR entries that
were written to `.codebase-memory/adr.md` and later overwritten. Each entry below is reproduced
exactly as it was last submitted, with no editing, no summarizing, and no supersession
annotation. Root cause, timeline, and the restoration decision live in
`.Dossier/ADR Store Loss - Root Cause and Recovery Inventory.md`.

The four oldest entries (ADR-2026-06-11 and three ADR-2026-07-12 entries) plus the store's
`PURPOSE` / `STACK` / `ARCHITECTURE` / `PATTERNS` / `TRADEOFFS` / `PHILOSOPHY` prose sections are
NOT copied here because they survive durably in git and can be read with
`git show 3a935f3:.codebase-memory/adr.md`.

---

## Recovered: ADR-2026-07-17B

Last submitted 2026-07-19T10:57:26.853Z.

```markdown
### ADR-2026-07-17B: Wall Spike side-face directional full-height resolver
Status: ACCEPTED FOR ISOLATED WALL SPIKE after Revit 2026 smoke PASS on 2026-07-17. Do not claim production Quick Dimension is updated until this resolver is ported to `QuickDimensionWallCandidateCollector.CollectSelectedWallEndAnchors` and smoked there.
Context: Prior Wall Spike models failed across L/T joint cases: top-face/intersection and join-role heuristics were rejected; vertical-only side-face edges passed T-joints but failed L-joints; simple joined-wall outward extension fixed some exterior cut corners but overshot interior corners or fell back to raw side-run endpoints. XML logs showed correct anchors depend on shell and end direction, not a single min/max or always-outward rule.
Decision: Keep the selected-wall side-face boundary model. Collect vertical side-face edge midpoints (`Reference != null`) and horizontal side-run endpoints. Choose the longest horizontal side-run endpoints as base anchors. Resolve each end directionally: for `Interior`, choose the nearest full-height vertical reference in the inward direction into the selected wall span from selected+joined side-line candidates; for `Exterior`, choose an outward joined full-height vertical reference when one exists on the side line, otherwise keep the selected-wall base. `CollectJoinedWallBoundaryCandidates` remains limited to candidates within `DistanceToSideLine <= 5mm` and `JoinExtensionMargin=500mm`.
Critical invariant: Full-height filtering must compute `maxMidpointZ` only from candidates with `Reference != null` and `Point != null`; horizontal endpoints at top elevation must never define the full-height threshold.
Evidence: User smoke on 2026-07-17 passed 12/12 Left/Right cases for walls 379467, 379468, 379469, 379470, 379933, and 380187. XML diagnostics recorded expected survey N/E anchors for selected wall cutting others, selected wall being cut, and T-joint cases.
Next step: Prove one-axis mixed L/T station aggregation for one selected wall through research and self-critique before porting the resolver or creating dimensions.
```

## Recovered: ADR-2026-07-17C

Last submitted 2026-07-19T10:57:26.853Z.

```markdown
### ADR-2026-07-17C: Quick Dimension remains one selected-wall chain at a time
Status: ACCEPTED.
Context: A production plan contains mixed L-joint and T-joint conditions. The Wall Spike 100% smoke pass proves individual left/right end-anchor extraction only; it does not prove that mixed joint stations form a correct continuous dimension chain. Bulk creation across many walls would prevent reliable operator validation of geometry-dependent output.
Decision: The main flow accepts exactly one straight, non-curtain host Wall and one side pick per invocation. It may create only one dimension chain on that wall axis. The chain contract must aggregate relevant L-joint/T-joint stations and later hosted opening jamb stations in ascending axis order, with explicit duplicate-station diagnostics. The resolver must be researched and self-criticized against mixed L-L, T-T, L-T, T-L, reversed-axis, and coincident-station counterexamples before any production code is written or ported.
Consequences: No automatic or multi-wall batch dimension creation. Per-joint anchor correctness is a required input invariant, not chain-level acceptance evidence. The Wall Spike remains open until aggregation is isolated, diagnosable, Revit-smoked, and only then integrated with the production collector.
```

## Recovered: ADR-2026-07-18A

Last submitted 2026-07-19T10:57:26.853Z.

```markdown
### ADR-2026-07-18A: Mid-run wall-joint station detection uses side-line reference evidence with endpoint-join exclusion
Status: ACCEPTED FOR RESEARCH/LOG-ONLY CONTRACT after Section 11 smoke PASS on selected wall 380815; REFINED on 2026-07-19 after four additional real smoke audits. Do not treat this as production collector implementation until the side-face end-anchor resolver is ported and the reference-preserving read-only wall-joint aggregator is implemented/smoked separately.
Context: Session 2.7 Section 11 smoke used a mixed fixture with one selected straight wall, two end joins, one oblique mid-run T-wall (381185), and a nearby parallel non-join wall (381035). `LocationCurve.get_ElementsAtJoin(0/1)` reported only the two selected-wall ends; `JoinGeometryUtils.GetJoinedElements` returned no mid-run wall. The mid-run T-wall was oblique, so a perpendicularity gate rejected a real station. A first log version collapsed two T-joint jambs into one midspan-best representative. Four follow-up smoke sets on 2026-07-19 then showed endpoint join walls can expose raw vertical side-line references and be overcounted if raw reference evidence is accepted before endpoint join provenance.
Decision: Raw mid-run evidence is a vertical `Edge.Reference` on the selected side line whose adjacent face normal is along the selected wall axis. Accepted mid-run classification must additionally exclude candidate walls in `ElementsAtJoinStart` or `ElementsAtJoinEnd`, reject endpoint-zone and resolved-anchor duplicate stations, sort/dedupe accepted stations, and require at least two distinct accepted stations (`acceptedMidRunStationCount >= 2`) before classifying a candidate wall as `MidRunCrossing`. Join API membership and perpendicularity remain diagnostics/provenance; join membership is an exclusion for true mid-run acceptance, not a positive mid-run signal. Reference ownership for the tested mid-run T-joint lives on the joining wall, not the selected wall.
Evidence: Right/Interior smoke on wall 380815 logged wall 381185 as `MidRunCrossing` with two `ReferenceHit` stations: 2665.787mm and 2867.289mm, both `candidateReferenceNormalAlongAxis=true`; all join provenance flags were false and `isPerpendicular=false`. Left/Exterior logged no mid-run hits for 381185, proving shell-specific behavior. Parallel near wall 381035 stayed `ParallelNonJoined`. 2026-07-19 audits found false `MidRunCrossing` records for end-join walls such as 380858, 379470, 379468, 379467, and 379469; each is now expected to remain `EndJoinOnly` even if raw `referenceHitCount > 0`.
Consequences: Future read-only wall-joint aggregation must assemble resolved Start/Finish anchors from ADR-2026-07-17B plus only accepted side-line reference hits from non-end-join candidate walls inside the resolved anchor span, ordered by projected station. It must dedupe coincident anchor/joint hits, preserve near-distinct jamb stations, reject endpoint artifacts, retain live `Reference` values for production, and still precede any `ReferenceArray`/`NewDimension` work.
```

## Recovered: ADR-2026-07-19A

Last submitted 2026-07-21T07:53:20.523Z.

```markdown
### ADR-2026-07-19A: Accepted mid-run stations exclude endpoint join artifacts
Status: ACCEPTED FOR LOG-ONLY CLASSIFIER AND FUTURE PRODUCTION AGGREGATION GATE. Follow-up Revit re-smoke on 2026-07-20 passed across four real Wall Spike sets; production collector + read-only aggregator port may proceed. `NewDimension` work remains blocked until reference-preserving read-only aggregation passes.
Context: Four real Wall Spike smoke audits showed that raw vertical side-line `Edge.Reference` evidence alone overcounts endpoint joins. Walls in `ElementsAtJoinStart` or `ElementsAtJoinEnd` can expose side-line reference hits at or near the selected wall endpoint/anchor zones; those hits were incorrectly classified as `MidRunCrossing` even though they are end-join artifacts. Some artifacts do not exactly equal the resolved Start/Finish anchor station, so anchor-station duplicate filtering alone is insufficient.
Decision: Treat `ElementsAtJoinStart` and `ElementsAtJoinEnd` as exclusion provenance for true mid-run acceptance. A true mid-run candidate must: expose vertical `Edge.Reference` hits on the selected side line; have `candidateReferenceNormalAlongAxis=true`; not belong to either selected-wall end join set; produce at least two distinct accepted stations after filtering; keep stations strictly inside the resolved Start/Finish anchor span; drop hits near resolved Start/Finish anchors and raw location-curve endpoint zones (`0`/`axisLength`); sort and dedupe by station tolerance. Raw `ReferenceHits` remain diagnostic evidence only; `acceptedMidRunStationCount` is the gate for `MidRunCrossing`.
Consequences: `QuickDimensionWallMidRunProbeService` receives `QuickDimensionWallSpikeResult`, separates raw hits from accepted station count, makes `EndJoinOnly` win before `MidRunCrossing`, and requires `acceptedMidRunStationCount >= 2`. `QuickDimensionWallSpikeXmlLogService` emits `acceptedMidRunStationCount` per candidate. Production aggregation must use the same acceptance rule while preserving live Revit `Reference` objects; the current probe remains value-only/log-only and must not be directly promoted into `ReferenceArray` creation. Re-smoke evidence: 380815→381185, 379467→379933, 379469→379933, and 379470→380187 were accepted only on the true mid-run shell; opposite shells and endpoint joins stayed clean.
```

## Recovered: ADR-2026-07-20A

Last submitted 2026-07-21T07:53:20.523Z.

```markdown
### ADR-2026-07-20A: Production wall-axis mid-run aggregator + read-only XML audit log
Status: ACCEPTED. Source complete and VS-MSBuild verified; awaiting Revit re-smoke on the four real wall sets using the new XML log. `NewDimension` still gated until reference-preserving read-only aggregation is Revit-confirmed.
Context: The proven Section 11 gates lived only in the value-only `QuickDimensionWallMidRunProbeService`. Production `QuickDimensionWallAxisAggregatorService.CollectMidRunCandidates` already ported the same acceptance gates (side-line tolerance 5mm, station eps 5mm, end-join exclusion, `NormalAlongAxis` gate, `>=2` distinct accepted stations, live `Reference` preserved) and is called from `QuickDimensionReadOnlyEngine.CollectWallAxisCandidates`. The read-only path could not be audited because `QuickDimensionReadOnlySummaryCommand` only showed a TaskDialog.
Decision: Keep the classifier as the single source of truth inside the aggregator. Add an optional `QuickDimensionWallAxisAggregationTrace` (plus candidate/reference-hit trace DTOs) populated by the aggregator during its single acceptance pass (no second geometry scan for non-end-join candidates; end-join artifacts are still recorded for audit). Carry the trace on `QuickDimensionReadOnlyResult`. Add `QuickDimensionReadOnlyXmlLogService` as a serializer-only writer (no classification, no collection, no transaction): it writes `ArcTool_QD_ReadOnlySummary_{wallId}_{side}_{timestamp}.xml` next to the `.rvt`, reusing the Wall Spike survey-coordinate convention (`GetProjectPosition`, n/e/elevation meters + `*_mm`). Keep the read-only XML block shape aligned with Wall Spike for side-by-side review: root metadata, `ReadOnlyResult` as the production equivalent of `ProbeResult` (anchors + options), `SelectedWall`, `WallMidRunAggregation` shaped like `MidRunProbe` (per-candidate relation + per-hit accept/reject reason), then production-only `FinalCandidates` and full unfiltered `Diagnostics`. `QuickDimensionReadOnlySummaryCommand` writes the XML after `CollectCandidates` regardless of `CanCreateChainDimension`, and surfaces the full path (or failure reason) in the TaskDialog.
Consequences: Read-only wall-only chains can now be audited from XML with the same fidelity and block vocabulary as the Wall Spike log while still retaining production-only data (`stableReference`, `accepted`, `rejectedReason`, `FinalCandidates`, full diagnostics). Production acceptance behavior is unchanged (gates only refactored into shared `PassesMidRunHitGates`/`GetMidRunHitRejectedReason` helpers used by both the acceptance list and the trace, so trace accept/reject never drifts from production). `dotnet build` cannot compile ArcTool.Core (COM Excel `ResolveComReference` unsupported on .NET Core MSBuild); use VS MSBuild (`Program Files/Microsoft Visual Studio/18/.../MSBuild.exe`).
```

## Recovered: ADR-2026-07-22A

Last submitted 2026-07-21T17:39:40.290Z.

```markdown
### ADR-2026-07-22A: Production read-only smoke passes; NewDimension line and opening semantics gates
Status: ACCEPTED. Production read-only collector + wall-axis aggregator are Revit-confirmed on four real wall sets after BUG-09. Phase 3 `NewDimension` remains gated by implementation-specific checks below.
Context: After fixing `QuickDimensionWallCandidateCollector.TryCollectWallEndAnchors` to use the resolved reference owner as `FinalCandidate.elementId` and selected wall as `hostElementId`, `QuickDimensionReadOnlySummaryCommand` was re-smoked on 380815, 379467, 379469, and 379470, both shells, with annotated plan images and XML logs. All visible survey labels matched XML `FinalCandidates`; all anchors/mid-run candidates carried stable references with blank `stableReferenceError`; final `elementId` now matches stable-reference owner. The refined classifier did not regress: 380815→381185, 379467→379933, 379469→379933, and 379470→380187 were accepted only on Interior/Right with `acceptedMidRunStationCount=2`; opposite shells stayed clean and end joins stayed `EndJoinOnly`.
Decision: Treat production read-only aggregation as passed for these four real fixtures, but do not start Phase 3 from raw wall-axis endpoints. `NewDimension` line creation must derive from the resolved final candidate span/points and cover valid exterior anchors outside the selected wall `LocationCurve` range. The read-only XML options block must accurately report wall-axis Grid collection as disabled. Before trusting generated dimensions on close-spaced openings, verify FamilyInstance Left/Right reference semantics where candidate owners interleave: test 379470 has Window 379479, Door 379482, and Window 379478 stations ordered 1452.409, 1858.409, 1997.909, 2912.909, 3052.409, 3458.409; spacing looks plausible but owner/side-label semantics need live `ReferenceArray` validation.
Consequences: BUG-09 is fixed/confirmed and may be retained only as a regression guard. Do not rewrite the proven collector/classifier when starting Phase 3. Next implementation work should be surgical: build dimension lines from final candidate span, correct XML options metadata, add a minimal `NewDimension` smoke path using existing ordered live `Reference` objects, and smoke 379470 first because it combines end joins, close openings, and a mid-run wall.
```

## Recovered: ADR-2026-07-30A

Last submitted 2026-08-03T12:46:12.339Z.

```markdown
### ADR-2026-07-30A: Failure-isolated, sequence-strict post-commit Quick Dimension chain audit
Status: ACCEPTED. Instrumentation/build/deployment complete; smoke #1 (wall 380815) both-shell audit PASS 2026-07-31; smokes #2/#3 (walls 379467/379469) both-shell creation/geometry PASS with sequence audit failures exposing BUG-11; smoke #4 (wall 379470, 2026-08-02) both-shell creation/geometry PASS confirming per-instance BUG-11 on Windows 379479/379478 and BUG-10 on Door 379482. Source/build follow-up landed 2026-08-03: the collector now derives each named `FamilyInstanceReferenceType.Left/Right` station from that same reference geometry, fallback candidate `elementId` now aligns with the live reference owner, and audit logging now reports normalized `actualSegmentCount` plus per-segment `valueSource`. EV-2 is the next runtime confirmation on the rebuilt DLL; EV-3 rollback remains blocked pending a verified safe post-start rollback fixture.
Context: `canCreateChainDimension` and pre-commit reference counts do not prove that a committed mixed-reference dimension retains per-reference identity/owners or produces expected segment values. Audit logging must not convert a committed model mutation into command failure, and BUG-10 candidate metadata must remain visible without becoming a false live-reference creation blocker. Revit may geometrically re-sort `Dimension.References`, so output can be correct even when the collector associated a stable reference with the wrong projected station.
Decision: After `CreateChainDimension` returns, read the committed `Dimension` by id without a transaction and append one `<ChainCreationAudit>` to the existing pre-mutation XML through a temp file plus `File.Replace`. Report creation and audit statuses independently. Compare stable-reference sequences as Exact, complete Reversed, or Mismatch; validate live-reference owners and segment values against the matching adjacent-station delta order. Keep sequence identity strict: do not auto-whitelist a local adjacent-pair swap and do not replace sequence comparison with unordered-set matching. BUG-11 root cause boundary was the collector's named `FamilyInstanceReferenceType.Left/Right` identity paired with an independently estimated physical station. The landed fix derives station directly from each named reference geometry and associates identity+station atomically while preserving named-reference `[0]` selection and fallback-only escape hatches. Log candidate `elementId` versus live-reference owner separately so BUG-10 remains observable but non-blocking. Handle the observed two-reference API shape through nullable `Dimension.Value`; use nullable `DimensionSegment.Value` for multi-segment dimensions. Report `actualSegmentCount` from values actually validated, not raw `Dimension.NumberOfSegments`, and identify the value source for the fallback.
Consequences: Smoke #1 wall 380815: dims 383577/383578 pass all gates. Smoke #2 wall 379467: dims 383579/383580 commit and match screenshots; Window 379477 exposes BUG-11 and Door 379481 exposes BUG-10 metadata-only divergence. Smoke #3 wall 379469: dims 384631/384632 commit with exact screenshot geometry (10/12 refs; 9/11 segments); local committed-to-expected mappings Left `[1,3,2,5,4,6,7,9,8,10]` Right `[1,3,2,5,4,6,7,8,9,11,10,12]`. Smoke #4 wall 379470: dims 384894/384895 commit with exact screenshot geometry (8/10 refs; 7/9 segments); Left `[1,3,2,4,5,7,6,8]` Right `[1,3,2,4,5,7,6,8,9,10]`; mid-run wall 380187 correct on Interior only at 206.084mm; old manual 379470 expected sequence superseded. All candidate and committed live owners correct in smokes #3/#4; top-level owner/segment false values are strict-sequence cascade, not attachment failure. BUG-10 runtime evidence now exists on doors 379481 and 379482, and the metadata owner fix is landed in source awaiting EV-2 confirmation. T5.2/T5.3 logging fixes are landed in source awaiting EV-2 confirmation on rebuilt output. Runtime remains operator-controlled.
```

