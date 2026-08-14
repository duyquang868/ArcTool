---
name: project_qd_chain_creation_audit_handoff
description: Quick Dimension ChainCreationAudit runtime handoff; BUG-11/BUG-10 mission CLOSED 2026-08-04 with EV-2 six-run matrix PASS and EV-3 reopen PASS; forced-rollback validation split into a standalone future task.
metadata:
  type: project
---

# Quick Dimension ChainCreationAudit handoff — updated 2026-08-04

## 2026-08-04 closure update

- **Mission CLOSED.** BUG-11 and BUG-10 are both runtime-confirmed fixed on the rebuilt DLL. No runtime evidence is owed by this mission.
- **EV-2 PASS (six runs).** `QuickDimensionCreateChainSmokeCommand` on walls 379467/379469/379470, both shells each: every run committed with `referenceOrderRelation=Exact`, `referenceIdentityMatched=true`, `referenceOwnersMatched=true`, `segmentValuesMatched=true`, and unchanged geometry. Committed dimensions: 385355/385356 (379467), 385632/385584 (379469), 385719/385720 (379470). Historical swap fixtures are clean — windows 379477/379475/379479/379478 and doors 379472/379471 no longer reproduce.
- **EV-3 reopen PASS.** All six committed dimensions survived save/close/reopen with unchanged displayed values and unchanged side/position.
- **Forced-rollback validation DEFERRED, not dropped** (operator decision 2026-08-04). Every operator-reachable invalid input returns `Result.Cancelled` *before* `Transaction.Start()`, so the rollback branches are unreachable by normal modelling and cannot be requested as a plain smoke. Deferral is safe because no task in the work package edited `QuickDimensionChainCreationService.cs`, making rollback behavior byte-identical to the pre-fix build and therefore not a regression surface for BUG-10/BUG-11. Standalone future task: `.Dossier/Quick Dimension - Deferred Rollback Validation Task.md`. Source analysis preserved in `.claude/quick-dimension-bugfix/results/T6.5_result.md`.
- **Do not reopen BUG-10 or BUG-11 without a new runtime defect,** and do not whitelist `LocalPairSwap` — only `Exact` and complete `Reversed` are acceptable audit order relations.

## 2026-08-03 source/build update

- BUG-11 source fix landed in `ArcTool.Core/Services/QuickDimensionDoorWindowCandidateCollector.cs`: named `FamilyInstanceReferenceType.Left/Right` references still use `GetReferences(...)[0]`, but each named reference now derives its projected station from that same reference geometry instead of pairing identity with a proxy/estimated point.
- BUG-10 metadata fix landed in `ArcTool.Core/Services/QuickDimensionDoorWindowCandidateCollector.cs`: fallback candidates now align `elementId` with the live reference owner while preserving `hostElementId` as the selected host wall.
- Audit logger fixes landed in `ArcTool.Core/Services/QuickDimensionReadOnlyXmlLogService.cs`: `actualSegmentCount` now uses the normalized measured-value count, and each `<Segment>` now records `valueSource` (`DimensionSegment.Value` vs `Dimension.Value`).
- Locked Visual Studio MSBuild build passed on 2026-08-03 and produced `ArcTool.Core/bin/x64/Debug/net8.0-windows/ArcTool.Core.dll`. This is the exact DLL EV-2 and EV-3 were run against.

## Scope completed

The section below is the preserved historical record of the earlier audit-slice work; it no longer describes the full 2026-08-03 code state.

Two source files were changed:

1. `ArcTool.Core/Services/QuickDimensionReadOnlyXmlLogService.cs`
   - Added `TryAppendChainCreationAudit(Document, string, QuickDimensionReadOnlyResult, QuickDimensionChainCreationResult)`.
   - Loads the exact read-only XML written before mutation and appends one `<ChainCreationAudit>` block after `CreateChainDimension` returns.
   - Saves to `<original>.audit.tmp`, then uses `File.Replace` so the original read-only XML remains intact if building/saving/replacing the audit fails.
   - Audit failure is returned as a status string; it does not throw into the creation result path.
   - Reads the committed `Dimension` back by `dimensionId`; it does not retain a live `Dimension` in `QuickDimensionChainCreationResult` and opens no audit transaction.
   - Emits required top-level attributes: `attempted`, `succeeded`, `message`, `transactionStatus`, `dimensionId`, expected/created reference counts, expected/actual segment counts, `referenceOrderRelation`, `referenceIdentityMatched`, `referenceOwnersMatched`, and `segmentValuesMatched`.
   - Emits `<ResolvedDimensionLine>`, `<ExpectedCandidates>`, `<CreatedReferences>`, and `<Segments>` with per-index identity/owner/value evidence.
   - `referenceOrderRelation` accepts only `Exact`, complete `Reversed`, or `Mismatch`; full reversal is valid evidence because Revit may normalize orientation.
   - Stable representations are the identity key. Owner IDs are not treated as unique reference keys because several references may share one owner.
   - Segment values are checked against adjacent station deltas in the matching forward/reversed orientation with a `0.1mm` tolerance.
   - Handles the observed Revit single-segment shape: for two references, `Dimension.NumberOfSegments` and `Dimension.Segments.Size` may both report 0 while the validated length remains available through nullable `Dimension.Value`; multi-segment chains use nullable `DimensionSegment.Value`.
   - Separates three ownership signals deliberately:
     - `referenceIdentityMatched`: stable reference identity survived creation.
     - top-level `referenceOwnersMatched`: committed live-reference owners match expected live-reference owners; this is the true Revit binding gate and is not falsely failed by BUG-10 metadata.
     - `ExpectedCandidates.elementIdMatchesReferenceOwner` and per-created-reference `ownerEqualToExpected`: candidate metadata versus live-reference owner, which exposes BUG-10 independently without making it a `NewDimension` blocker.
2. `ArcTool.Core/Commands/QuickDimensionCreateChainSmokeCommand.cs`
   - Calls the new audit append method immediately after `QuickDimensionChainCreationService.CreateChainDimension` returns.
   - Wraps audit append separately so no audit/log exception can turn a committed dimension into command failure.
   - TaskDialog now reports `Creation status`, `Creation message`, `Transaction status`, and `Audit status` independently.
   - Command result still follows `creationResult.Succeeded` only.

## API research completed

Revit 2026 documentation was verified before coding for:
- `ItemFactoryBase.NewDimension(View, Line, ReferenceArray)` — https://www.revitapidocs.com/2026/47b3977d-da93-e1a4-8bfa-f23a29e5c4c1.htm
- `Dimension.References` — https://www.revitapidocs.com/2026/fc3bc889-b274-3262-a126-849df2af9019.htm
- `Dimension.NumberOfSegments` — https://www.revitapidocs.com/2026/3e01937d-a001-8fd4-9cc8-270ad4b4ba10.htm
- `Dimension.Segments` — https://www.revitapidocs.com/2026/d7fcdab2-ca81-0ed1-4813-f7aa092430d7.htm
- `DimensionSegment.Value` — https://www.revitapidocs.com/2026/d4a5ac3d-c5c4-b7d8-2555-b04d2f26e422.htm
- `Reference.ElementId` — https://www.revitapidocs.com/2026/909ec304-3c41-8319-4c80-efedce795d7f.htm
- `Reference.ConvertToStableRepresentation(Document)` — https://www.revitapidocs.com/2026/9d821d63-5b4a-b814-25b2-b92f7d5d1425.htm
- `Transaction.Commit()` — https://www.revitapidocs.com/2026/32714010-7138-f64f-8fde-a310354448e3.htm

Documentation boundary retained: the `ItemFactoryBase.NewDimension` page explicitly documents geometric references and `ArgumentException` for non-geometric references; stronger parallel/reference-count constraints observed at runtime or on related API pages should not be misquoted as text from this exact overload.

## Build evidence

Visual Studio MSBuild was run repeatedly after implementation and after ownership-semantics corrections:

`C:\Program Files\Microsoft Visual Studio\18\Community\MSBuild\Current\Bin\MSBuild.exe ArcTool.slnx -m -nologo -v:minimal`

Final result: PASS, `0 errors`. The only C# warning is the pre-existing `CS8600` at `QuickDimensionReadOnlyXmlLogService.cs:77` (`Path.GetDirectoryName` assigned to non-nullable `string`), unrelated to this slice. Earlier verbose build also showed the known `MSB3246` Revit link/native reference warnings; output DLL was still produced successfully.

## Deployment state

The final built audit DLL was copied to the active Revit add-in folder because exact `Revit.exe` was not running:

- Build output: `ArcTool.Core/bin/x64/Debug/net8.0-windows/ArcTool.Core.dll`
- Deployed target: `C:\ProgramData\Autodesk\Revit\Addins\2026\ArcTool\ArcTool.Core.dll`
- Verified SHA-256 for source and deployed DLL: `08459e5aa1a9e171fe0e2049dbcee369be78dabce184cfd045022baf65ba71c2`
- Previous deployed `ArcTool.Core.dll/.pdb/.deps.json` backup: `C:\ProgramData\Autodesk\Revit\Addins\2026\ArcTool\_backup_20260730_074254`

No Revit model was opened and no runtime smoke was run by Claude during the implementation session. Subsequent operator-controlled smoke tests #1 and #2 are recorded below.

## Journal evidence discovered (historical, not a substitute for the new audit)

Reading recent Revit journals proved useful for cross-source diagnosis and fixture discovery:

- Fixture/model: `C:\Users\ADMIN\Desktop\PA4\Project3.rvt`, active plan `Floor Plan: Level 1`.
- Read-only/creation XML directory: `C:\Users\ADMIN\Desktop\PA4\`.
- Journal `journal.0444.txt` records a historical 2026-07-27 run of `QuickDimensionCreateChainSmokeCommand` on wall 379470 Exterior/Left:
  - command `Outcome: SUCCESS`
  - transaction `Committed`
  - 8 final candidates / 8 references
  - dimension id `383148`
  - resolved span `-113.89..4325.37mm`
  - XML `ArcTool_QD_ReadOnlySummary_379470_Left_20260727_212453.xml`
- Older journal `journal.0449.txt` also records successful committed create-chain runs on 380815 and 379467, including a fallback-capable fixture, but those historical runs predate `<ChainCreationAudit>` and therefore do not close the durable identity/segment gate.
- Journals show a pre-existing non-fatal assembly-version conflict on ArcTool commands: RevitAPI/RevitAPIUI `26.0.10.0` versus Revit preloaded `26.0.4.0`, and SkiaSharp `3.119.0.0` versus preloaded `2.88.0.0`. Commands continued to execute and commit; this packaging issue is outside the current audit slice.

Use journal analysis as an independent evidence source alongside XML, images, code, and user observations. Follow [[revit-runtime-operator-control-and-journal-analysis]].

## Runtime status and exact next-session boundary

### Real smoke test #1 — wall 380815 — 2026-07-31 — PASS

Operator ran `QuickDimensionCreateChainSmokeCommand` on wall 380815, both shells. XML files:
- `C:\Users\ADMIN\Desktop\PA4\ArcTool_QD_ReadOnlySummary_380815_Left_20260731_075004.xml`
- `C:\Users\ADMIN\Desktop\PA4\ArcTool_QD_ReadOnlySummary_380815_Right_20260731_075016.xml`

Wall 380815 context: selected straight wall on `Floor Plan: Level 1` of `Project3.rvt`. Has a T-junction mid-run crossing wall 381185 on the Interior/Right shell and no mid-run on the Exterior/Left shell. End-join wall 380858 at the Finish Anchor on Interior/Right.

**Left / Exterior — dim 383577:**
- Transaction: Committed. References: 2 created / 2 expected.
- Stable-reference identity: `referenceIdentityMatched="true"`, order `Exact`.
- Live owners: `referenceOwnersMatched="true"`. Both references owned by 380815 (WallSideFace).
- Segments: expected 1, `segmentValuesMatched="true"`. Span = **4952.493 mm** → visible label 4952 ✓.
- Both candidates: WallSideFace Start Anchor and Finish Anchor on wall 380815.

**Right / Interior — dim 383578:**
- Transaction: Committed. References: 4 created / 4 expected.
- Stable-reference identity: `referenceIdentityMatched="true"`, order `Exact`.
- Live owners: `referenceOwnersMatched="true"`.
- Segments: expected 3, `segmentValuesMatched="true"`. Values: **2608.052 / 201.502 / 2128.527 mm** → visible labels 2608 / 202 / 2129 ✓.
- Candidates: Start Anchor (380815), mid-run crossing wall 381185 × 2 (`MidRunCrossing`), Finish Anchor from end-join wall 380858 (`EndJoinOnly`).

**BUG-10 status:** Not triggered on wall 380815 — no `HostWallOpeningGeometry` fallback openings present. All candidates resolved via `FamilyInstanceLeftRight` or `WallSideFace`. BUG-10 metadata attributes `elementIdMatchesReferenceOwner` were not stressed; still requires fallback fixture.

### Single-segment reporting inconsistency (cosmetic, non-blocking)

Root cause confirmed by source, this runtime XML, and Revit API documentation:
- Runtime evidence from dim 383577: `dimension.NumberOfSegments` was observed as **0** and `Dimension.Segments.Size` as **0** for this two-reference dimension.
- Revit 2026 documents `Dimension.Value` as the nullable value for a single-segment linear dimension and as null for multi-segment linear dimensions; therefore it is the correct fallback source here. The documentation defines `NumberOfSegments` but does not by itself justify generalizing this observed zero to every two-reference shape.
- `ExtractSegmentValues` handles the observed API shape correctly: fallback to `Dimension.Value` when `segmentArray.Size == 0` and `expectedReferenceCount == 2`. The fallback is guarded and produces the correct `4952.493 mm` value.
- `actualSegmentCount` attribute is written from `dimension.NumberOfSegments` (= 0), not from `segmentValuesInternal.Count` (= 1). The XML entry and the `<Segment>` child element are therefore contradictory: `actualSegmentCount="0"`, `segmentArraySize="0"`, but one `<Segment matched="true">` exists.
- `segmentValuesMatched="true"` is accurate; the value check ran correctly.

**Fix required (next session, low priority):** In `TryAppendChainCreationAudit`, compute `reportedSegmentCount` as `segmentValuesInternal.Count` (the number of values actually validated) rather than `dimension.NumberOfSegments`. Add `valueSource="Dimension.Value"` attribute to `<Segment>` when the fallback was used. This is one-line cosmetic — no transaction, no reference, no gate-blocking logic changes.

**Fix must NOT be confused with a creation defect.** The dimension 383577 is correct. The segment value is correct. Only the XML count attribute is misleading.

### Real smoke test #2 — wall 379467 — 2026-07-31 — PASS geometry, FAIL audit (false negative)

Operator ran `QuickDimensionCreateChainSmokeCommand` on wall 379467, both shells. XML files:
- `C:\Users\ADMIN\Desktop\PA4\ArcTool_QD_ReadOnlySummary_379467_Left_20260731_090712.xml`
- `C:\Users\ADMIN\Desktop\PA4\ArcTool_QD_ReadOnlySummary_379467_Right_20260731_090717.xml`

Wall 379467 context: selected straight wall on `Floor Plan: Level 1` of `Project3.rvt`. Selected wall type `Generic - 200mm`, width 200mm, `LocationCurve` axis length 6300mm. Survey-coordinate axis endpoints are N 15102.527 / E 4982.136 and N 18441.018 / E 10324.839; axis direction `(0.8480481, 0.5299193, 0)`. Has no mid-run crossing on Exterior/Left shell; has a T-junction mid-run crossing wall 379933 on the Interior/Right shell producing three distinct consecutive segments at that junction. Includes windows 379476/379477 and doors 379480/379481; window 379477 uses `FamilyInstanceLeftRight`, while door 379481 uses `HostWallOpeningGeometry` fallback.

**Fixture reconstruction map:**

| Shell | Station mm | Source | Element / live reference owner | Stable-reference suffix / role |
|---|---:|---|---|---|
| Left/Exterior | -69.965 | Start anchor | 379467 / 379467 | `:36:LINEAR` |
| Both | 897 | Window jamb | 379476 / 379476 | `:0:SURFACE` |
| Both | 1303 | Window jamb | 379476 / 379476 | `:2:SURFACE` |
| Both | 1842.5 | Door fallback jamb | candidate 379481 / live owner 379467 | `:109:LINEAR` — BUG-10 metadata divergence |
| Both | 2757.5 | Door fallback jamb | candidate 379481 / live owner 379467 | `:117:LINEAR` — BUG-10 metadata divergence |
| Both | 3098.878 | Window candidate | 379477 / 379477 | collector associates `:0:SURFACE`; committed geometric order shows `:2` first — BUG-11 |
| Both | 3504.878 | Window candidate | 379477 / 379477 | collector associates `:2:SURFACE`; committed geometric order shows `:0` second — BUG-11 |
| Right/Interior | 3804.878 | Mid-run wall | 379933 / 379933 | `:28:LINEAR` |
| Right/Interior | 4006.38 | Mid-run wall | 379933 / 379933 | `:30:LINEAR` |
| Both | 4542.5 | Door jamb | 379480 / 379480 | `:0:SURFACE` |
| Both | 5457.5 | Door jamb | 379480 / 379480 | `:2:SURFACE` |
| Right/Interior | 69.965 | Start anchor | 379470 / 379470 | `:23:LINEAR` |
| Right/Interior | 6217.47 | Finish anchor | 379468 / 379468 | `:23:LINEAR` |
| Left/Exterior | 6382.53 | Finish anchor | 379467 / 379467 | `:25:LINEAR` |

Key survey-coordinate evidence: Left anchors N/E 15150.256/4869.810 and 18569.557/10341.836; Right anchors N/E 15054.798/5094.462 and 18312.479/10307.842; Right mid-run points N/E 17034.000/8261.847 and 17140.780/8432.731. Resolved dimension-line side offsets were 682.862mm Left and 726.208mm Right.

**Left / Exterior — dim 383579:**
- Transaction: Committed. References: 10 created / 10 expected.
- Exact segment values: 966.965 / 406 / 539.5 / 915 / 341.378 / 406 / 1037.622 / 915 / 925.03 mm.
- Screenshot-visible labels: 967 / 406 / 540 / 915 / 341 / 406 / 1038 / 915 / 925 — exact rounded match ✓.
- Audit XML: `referenceIdentityMatched="false"`, `referenceOwnersMatched="false"`, `segmentValuesMatched="false"`, `referenceOrderRelation="Mismatch"`.
- Window 379477 is the only sequence divergence: expected candidates associate stable-reference suffix `:0:SURFACE` with station 3098.878mm and `:2:SURFACE` with 3504.878mm, while committed `Dimension.References` returns `:2` then `:0`. Revit geometrically sorted the pair and produced correct segments.
- BUG-10 confirmed: door 379481 fallback candidate logs `elementId=379481`, but its committed stable-reference owner is host wall 379467. The live-reference owner is correct (379467); only candidate metadata diverges.

**Right / Interior — dim 383580:**
- Transaction: Committed. References: 12 created / 12 expected.
- Exact segment values: 827.035 / 406 / 539.5 / 915 / 341.378 / 406 / 300 / 201.502 / 536.12 / 915 / 759.97 mm.
- Mid-run wall 379933 contributed two references creating three consecutive segments: 300 / 201.502 / 536.12 mm.
- Screenshot-visible labels: 827 / 406 / 540 / 915 / 341 / 406 / 300 / 202 / 536 / 915 / 760 — exact rounded match ✓.
- Audit XML: the same window 379477 `:0` / `:2` local sequence divergence and the same door 379481 BUG-10 metadata divergence.

**Analysis and classification:**
- **Geometry/NewDimension: PASS on both shells.** Revit committed both dimensions with the expected reference counts; every displayed and API segment value matches the physical station deltas. The Interior mid-run T-junction is correct.
- **Audit as an output-only verdict: false negative.** Its three top-level booleans are false although the committed dimensions are visually/numerically correct; owner and segment checks cascade to false when order is `Mismatch`.
- **Audit as an input-invariant detector: valid failure.** The collector's candidate sequence says window 379477 `:0` is the lower station and `:2` is the higher station, but Revit's committed geometric order proves the opposite. This exposes **BUG-11**: named `FamilyInstanceReferenceType.Left/Right` identities can be cross-associated with separately estimated/fallback physical station points. Instance flip/mirror/orientation state is a hypothesis to inspect, not established evidence. Revit auto-sorting hides this in the final dimension, but downstream logic must not rely on the candidate identity↔station association.
- **BUG-10: confirmed Low, metadata-only, non-blocking.** `QuickDimensionChainCreationService` uses `candidate.Reference`, never `candidate.ElementId`, so fallback metadata divergence does not affect attachment.

**Required next-session implementation boundary (HISTORICAL — all four items were completed by 2026-08-04; see the 2026-08-04 closure update at the top of this file):**
1. Diagnose window 379477 to prove the physical projected station of each named Left/Right `Reference` directly from the reference geometry rather than pairing named references with independently estimated points by label.
2. Correct the collector's reference↔station association. After the fix, rerun wall 379467 both shells. Expected audit order is `Exact` or complete `Reversed`; do not auto-whitelist `LocalPairSwap` and do not reduce identity checking to an unordered set.
3. Keep `CreateChainDimension`, wall anchors, mid-run aggregation, and geometry math unchanged unless the corrected collector produces a concrete runtime defect.
4. After BUG-11 is fixed, apply the independent cosmetic single-segment count/value-source logger fix and the Low BUG-10 metadata fix.

### Operator matrix — final state 2026-08-04

1. **BUG-11 regression re-smokes after collector fix — DONE, PASS.** Walls 379467, 379469, and 379470, both shells (EV-2, six runs). All reported `Exact` order, all identity/owner/segment gates true, unchanged geometry.
2. **Reopen persistence — DONE, PASS.** Committed dimensions 385355/385356/385632/385584/385719/385720 all persisted across save/close/reopen with unchanged visible labels and unchanged side/position (EV-3).
3. **Intentional invalid-reference rollback — DEFERRED to a standalone future task**, not required for this mission. See `.Dossier/Quick Dimension - Deferred Rollback Validation Task.md`. It needs a fault-injection harness or debug-only switch; it cannot be requested as a plain operator smoke because invalid input cancels before `Transaction.Start()`.

Wall 379470 both-shell creation/geometry was closed by smoke #4. The earlier manual expected sequences (`578,406,500,915,540,406,83,206,438` and `806,406,500,915,540,406,867`) are superseded: current XML and screenshots agree on the actual fixture sequences recorded below. BUG-10 fallback evidence is independently confirmed by doors 379481 and 379482. The BUG-11/BUG-10 mission is now closed; only the deferred rollback track remains open, tracked outside this package.

### Real smoke test #3 — wall 379469 — 2026-08-01 — PASS geometry both shells, FAIL audit (BUG-11 broadened)

Operator ran `QuickDimensionCreateChainSmokeCommand` on wall 379469, both shells (NEW create smoke; both XML carry a committed `<ChainCreationAudit>`). Files:
- `C:\Users\ADMIN\Desktop\PA4\ArcTool_QD_ReadOnlySummary_379469_Left_20260801_144328.xml` (Exterior, sideSign +1, 10 candidates, dim 384631)
- `C:\Users\ADMIN\Desktop\PA4\ArcTool_QD_ReadOnlySummary_379469_Right_20260801_144333.xml` (Interior, sideSign −1, 12 candidates, dim 384632)

**Left / Exterior — dim 384631 (Committed, 10/10 refs, 9 segments):** actual segments 1501.707/500/1092.5/915/930.838/500/554.162/915/2493.83mm. Reversed to lower-left→upper-right and rounded → 2494/915/554/500/931/915/1093/500/1502 = image exactly. Segment sum 9403.037mm = ResolvedDimensionLine span.

**Right / Interior — dim 384632 (Committed, 12/12 refs, 11 segments):** actual segments 1196.398/500/1092.5/915/316.876/200.963/413/500/554.162/915/2266.043mm. Reversed+rounded → 2266/915/554/500/413/201/317/915/1093/500/1196 = image exactly. Segment sum 8869.942mm vs span 8869.941mm (0.001mm serialization rounding). Mid-run wall 379933 present only on Interior/Right with two refs at 4173.428/4374.391mm (200.963mm ≈ wall width), `Ignored` on Exterior/Left — shell-directional mid-run holds.

**Geometry/NewDimension: PASS both shells.** Every displayed and API segment matches the physical station deltas; anchors resolve correctly (Left extends outward to 379468/379469; Right snaps inward/end-joins to 379469/379470).

**Audit: `referenceOrderRelation="Mismatch"`, all three booleans false on both shells — output-level false negative but a valid BUG-11 input-invariant detection.**

**BUG-11 BROADENED (key new evidence).** `CreatedReferences.matchedExpectedCandidateIndex`:
- Left: [1,3,2,5,4,6,7,9,8,10] → swapped local pairs = window **379475**, door **379472**, door **379471**; NOT swapped = window **379484**.
- Right: [1,3,2,5,4,6,7,8,9,11,10,12] → swapped = window **379475**, door **379472**, door **379471**; NOT swapped = window **379484** and mid-run wall **379933**.
- Reproducible across both shells and both categories (Door + Window), no longer a single-window symptom (smoke #2 was only window 379477).
- **Decisive diagnostic: window 379475 (swapped) and window 379484 (NOT swapped) are the SAME family/type `M_Fixed 0406 x 0610mm`.** The inversion is per-instance, not per-type. Instance orientation/flip/mirror relative to wall direction becomes the leading hypothesis, but these XML files do not log those flags, so it is not yet proven. It is NOT a global rule, so `LocalPairSwap` whitelisting is still forbidden.

**No live-reference ownership loss:** every ExpectedCandidate `elementIdMatchesReferenceOwner=true` and every CreatedReference `ownerEqualToExpected=true`. Top-level `referenceOwnersMatched=false` is only the strict-positional-sequence cascade after `Mismatch`, not a real owner defect. BUG-10 does NOT manifest — all openings resolved via `FamilyInstanceLeftRight` (no `HostWallOpeningGeometry` fallback).

**Cosmetic logger note unchanged:** multi-segment path here reports `actualSegmentCount` correctly (9 and 11); the single-segment `NumberOfSegments=0` fallback issue from smoke #1 is still queued.

**Phase 3 gate impact:** smoke #3 adds strong feasibility evidence (12-ref mixed LINEAR/SURFACE + mid-run chain commits intact) but does NOT close the input-invariant gate. Required next source change is still the BUG-11 collector reference↔station fix (now with orientation-dependent per-instance evidence), then re-smoke 379467 and 379469 expecting `Exact`/complete-`Reversed`. Remaining matrix items: 379470 both shells, intentional rollback, reopen persistence.

### Real smoke test #4 — wall 379470 — 2026-08-02 — PASS geometry both shells, FAIL audit (BUG-11 confirmed, BUG-10 confirmed)

Operator ran `QuickDimensionCreateChainSmokeCommand` on wall 379470, both shells. XML files:
- `C:\Users\ADMIN\Desktop\PA4\ArcTool_QD_ReadOnlySummary_379470_Left_20260802_092358.xml` (Exterior, sideSign +1, 8 candidates, dim 384894)
- `C:\Users\ADMIN\Desktop\PA4\ArcTool_QD_ReadOnlySummary_379470_Right_20260802_092403.xml` (Interior, sideSign −1, 10 candidates, dim 384895)

Wall 379470 context: selected straight wall type `Generic - 200mm` on `Floor Plan: Level 1` of `Project3.rvt`. Raw axis length 4255.409mm. Hosts Window 379479 (`M_Fixed 0406 x 0610mm`, stations 691.909/1097.909mm), Door 379482 (`HostWallOpeningGeometry` fallback, stations 1597.909/2512.909mm, nominal width 915mm), and Window 379478 (`M_Fixed 0406 x 0610mm`, stations 2768.847/3174.847mm). Shell-specific mid-run wall 380187 appears only on Right/Interior.

**Fixture reconstruction — Left / Exterior (shell sideSign +1):**

| Station mm | Source | Element / live reference owner |
|---:|---|---|
| −113.893 | Start anchor (extends outward) | wall 379469 |
| 691.909 | Window 379479 jamb | 379479 |
| 1097.909 | Window 379479 jamb | 379479 |
| 1597.909 | Door 379482 fallback jamb | candidate 379482 / live owner 379470 — BUG-10 |
| 2512.909 | Door 379482 fallback jamb | candidate 379482 / live owner 379470 — BUG-10 |
| 2768.847 | Window 379478 jamb | 379478 |
| 3174.847 | Window 379478 jamb | 379478 |
| 4325.374 | Finish anchor (extends outward) | wall 379467 |

Survey-coordinate anchors: Start N 10807.308 / E 5789.536mm; Finish N 15150.256 / E 4869.810mm.

**Left / Exterior — dim 384894 (Committed, 8/8 refs, 7 segments):**
- Actual segment values: 805.802 / 406 / 500 / 915 / 255.938 / 406 / 1150.527mm.
- Segment sum = 4439.267mm = resolved span (4325.374 − (−113.893)).
- Screenshot labels top-to-bottom: 1151 / 406 / 256 / 915 / 500 / 406 / 806 — reversed-rounded form matches exactly ✓.
- Committed mapping `[1,3,2,4,5,7,6,8]`: Windows 379479 and 379478 both locally swap. Door 379482 pair (indices 4,5) stays in order.
- `referenceOrderRelation="Mismatch"`, all three gates false (BUG-11 cascade).
- BUG-10: Door 379482 candidate `elementId=379482`; live/stable reference owner = host wall 379470. Dimension attachment correct.

**Fixture reconstruction — Right / Interior (shell sideSign −1):**

| Station mm | Source | Element / live reference owner |
|---:|---|---|
| 113.893 | Start anchor (snaps inward) | wall 379470 |
| 691.909 | Window 379479 jamb | 379479 |
| 1097.909 | Window 379479 jamb | 379479 |
| 1597.909 | Door 379482 fallback jamb | candidate 379482 / live owner 379470 — BUG-10 |
| 2512.909 | Door 379482 fallback jamb | candidate 379482 / live owner 379470 — BUG-10 |
| 2768.847 | Window 379478 jamb | 379478 |
| 3174.847 | Window 379478 jamb | 379478 |
| 3541.796 | Mid-run wall 380187 | 380187 / `:23:LINEAR` |
| 3747.881 | Mid-run wall 380187 | 380187 / `:25:LINEAR` |
| 4185.444 | Finish anchor (snaps inward) | wall 379470 |

Survey-coordinate anchors: Start N 11071.588 / E 5938.004mm; Finish N 15054.798 / E 5094.462mm. Mid-run refs: N 14425.116 / E 5227.812mm and N 14626.729 / E 5185.116mm.

**Right / Interior — dim 384895 (Committed, 10/10 refs, 9 segments):**
- Actual segment values: 578.015 / 406 / 500 / 915 / 255.938 / 406 / 366.949 / 206.084 / 437.563mm.
- Segment sum = 4071.549mm ≈ resolved span 4071.550mm (0.001mm serialization rounding — within tolerance).
- Screenshot labels top-to-bottom: 438 / 206 / 367 / 406 / 256 / 915 / 500 / 406 / 578 — reversed-rounded form matches exactly ✓.
- Committed mapping `[1,3,2,4,5,7,6,8,9,10]`: same Windows 379479 and 379478 swap as Left. Door and mid-run pairs (indices 4,5 and 9,10) in order.
- `referenceOrderRelation="Mismatch"`, all three gates false (BUG-11 cascade).

**Geometry/NewDimension: PASS both shells.** All segment values match adjacent station deltas exactly; anchor logic is directional-correct (Left extends outward to joined walls; Right snaps inward to selected wall). Mid-run wall 380187 correct on Right/Interior only.

**BUG-11 CONFIRMED ON 379470.** Affected per-instance: Windows 379479 and 379478 (both `M_Fixed 0406 x 0610mm`) — same family/type, both swap consistently on both shells. Door 379482 fallback pair does NOT swap. This further disproves any per-type or global rule. Flip/mirror/orientation per-instance remains the leading hypothesis.

**BUG-10 CONFIRMED ON 379470.** Door 379482 fallback: candidate `elementId=379482`, live/stable owner = host wall 379470. Committed attachment correct; metadata-only divergence.

**No new creation or geometry defect.** `segmentValuesMatched=false` is the designed strict-gate cascade when `referenceOrderRelation=Mismatch`, not evidence of numerically wrong geometry.

**Note on old expected sequences.** The prior manual estimate (`Right: 578,406,500,915,540,406,83,206,438`; `Left: 806,406,500,915,540,406,867`) used stale gap widths (540 and 867/83 instead of 255.938 and 366.949/1150.527). Current XML and screenshots agree — the prior estimate was based on an approximated or earlier fixture state. No geometry defect.

## Closure verification — 2026-08-02 (smoke #4, wall 379470)

- Both XML files were read in full; screenshot labels independently verified against rounded segment values.
- Numerical identity verified: Left sum = 4439.267mm = resolved span exactly; Right sum = 4071.549mm vs span 4071.550mm (0.001mm serialization rounding, under 0.1mm tolerance).
- BUG-11 confirmed per-instance on Windows 379479/379478 (both shells, consistent swap pattern); BUG-10 confirmed on Door 379482 fallback (both shells, metadata-only, non-blocking).
- Mid-run wall 380187 resolves correctly on Right/Interior only; Left/Exterior has no mid-run as expected.
- Old manual expected matrix for 379470 is superseded; current XML/screenshot values are authoritative.
- Durable records updated: this handoff, `.Dossier/Quick Dimension - Implementation Roadmap.md`, `CLAUDE.md`, `Memory/MEMORY.md`, and ADR-2026-07-30A.
- No source code changed, no build run, no Revit model opened, no Revit MCP invoked, no codebase-memory re-index performed by Claude.

## Closure verification — 2026-08-01 (smoke #2 batch)

- Durable records updated: this handoff, `.Dossier/Quick Dimension - Implementation Roadmap.md`, `CLAUDE.md`, `Memory/MEMORY.md`, and ADR-2026-07-30A.
- Fixture reconstruction checks passed: Left segment sum = 6452.495mm, exactly matching resolved span; Right segment sum = 6147.505mm, exactly matching resolved span.
- Visual Studio MSBuild command passed with output `ArcTool.Core/bin/x64/Debug/net8.0-windows/ArcTool.Core.dll` and no reported build error.
- No Revit model was opened, no Revit MCP was invoked, and no smoke test was run during closure.
- No codebase-memory re-index was run; re-index remains the final optional user-directed step and does not block this handoff.

## Closure verification — 2026-08-01 (smoke #3, wall 379469)

- Analysis was read-only over the two provided create-chain XML files plus source (`QuickDimensionReadOnlyXmlLogService`, `QuickDimensionChainCreationService`, `QuickDimensionDoorWindowCandidateCollector`) and Revit 2026 docs for `FamilyInstance.GetReferences`, `FamilyInstanceReferenceType`, and `Dimension.References`. Independently cross-checked by a verification sub-agent (high confidence).
- Numerical identity verified directly: Left actual-segment sum = 9403.037mm = ResolvedDimensionLine span; Right actual-segment sum = 8869.942mm vs span 8869.941mm (0.001mm serialization rounding, under 0.1mm audit tolerance). Every rounded segment matches both screenshots.
- BUG-11 broadened and reclassified as per-instance (Window 379475 + Doors 379472/379471 swapped; same-type Window 379484 not swapped); orientation/flip remains an unproven hypothesis because those flags are not logged.
- No live-reference owner loss; BUG-10 not exercised (no fallback).
- Durable records updated for smoke #3: this handoff, roadmap runtime record + status block + next-session prompt, `CLAUDE.md` header/BUG-11/status, `Memory/MEMORY.md`, and ADR-2026-07-30A.
- No Revit model opened, no Revit MCP invoked, no smoke run, and no codebase-memory re-index performed by Claude during this closure.

See [[project_qd_midrun_smoke_evidence]] and [[feedback_smoke_test_single_session_close]].
