---
name: project_qd_lt_aggregation_research
description: Session 2.7 research conclusion — one-axis mixed L/T station-aggregation contract for Quick Dimension; no-code decision and the mid-run gap that keeps Wall Spike open.
type: project
---

Session 2.7 (2026-07-17) researched the one-axis station-aggregation contract for one selected straight wall and produced a research-only conclusion. NO code was written; Wall Spike stays OPEN (ADR-2026-07-17C).

**Executive verdict:** Aggregation model is "conditionally plausible" but NOT proven. The 12/12 smoke pass proves only two end anchors per side pick. It does NOT prove collection of mid-run T-joint stations.

**Core gap (highest risk):** `QuickDimensionWallReferenceProbeService.RunWallReferenceProbe` only resolves the two ends (base Start/Finish from longest side-run + directional full-height resolve). It has NO loop that scans walls cutting into the MIDDLE of the selected side-line. `LocationCurve.get_ElementsAtJoin(0/1)` only returns joins at the two curve ENDS (revitapidocs 2026), so a separate mid-run collection contract is required — end-anchor resolver cannot discover mid-run T-joints.

**Decisive Revit API constraint (verified):** `NewDimension` requires all references mutually PARALLEL and PERPENDICULAR to the dimension line, geometric `Reference` only, and geometry VISIBLE in the target view (Autodesk Dimensions & Constraints guide; rvtdocs NewDimension). For a wall-axis chain the dimension line is perpendicular to the axis, so EVERY station's reference must be a vertical face/edge whose normal is ALONG the wall axis. This confirms the spike invariant "full-height threshold only from candidates with Reference != null" must extend to mid-run, and forbids horizontal side-run endpoints (no usable axis-normal reference) from the dimension-eligible tier — they may only fix the `t` station.

**Three-tier candidate classification (must not merge):**
- Geometry-only: projects to axis but fails side-line or has no reference → not in ReferenceArray.
- Diagnostics-only (`UnsupportedReference`): plausible corner, Reference == null or wrong normal → log only, never in ReferenceArray (no silent point-only fallback, per ADR-2026-07-17C).
- Dimension-eligible: CanonicalStation with axis-normal `Reference` → into ReferenceArray.

**Ordering/dedupe policy:** order by `t` (ProjectParameter). Reversed axis only flips labels/order, never the physical station set. Dedupe cluster key must add reference/owner identity (not `t` alone) so two genuinely distinct near-coincident architectural references are logged, not blindly collapsed. Keep existing tolerances unchanged for smoke comparability: `DuplicateTolerance≈1e-4 ft`, `sideLineTol=5mm`, `fullHeightTol=10mm`, `joinExtensionMargin=500mm`. Do not invent new numeric tolerances from theory — derive from smoke.

**Counterexample verdicts:** L-L/T-T/L-T/T-L/reversed-axis/coincident/join-cleanup/compound-shell = PASS or handled. FAIL/REFINE: mixed-one-axis with mid-run T (case 5, core gap), mid-run vs end-join distinction (case 6, needs experiment), reversed side pick (case 8, shell-specific not symmetric), near-coincident-distinct (case 10, key needs ref/owner), nonjoined proximity (case 14, must test real join not distance).

**Diagnostics contract:** extend XML log with `classification`, `ownership(selected|joined)`, `hasReference`, `referenceNormalAlongAxis`, `distanceToSideLineMm`, `clusterId`, add a `<ChainStations>` block, KEEP `<Corners>` intact for regression vs the 12 old logs.

**Acceptance gates before ANY code:** (1) port spike resolver into production `CollectSelectedWallEndAnchors` is a SEPARATE gate, smoke 12 old cases with no regress; (2) a fixture with one L end + one T end + ≥1 mid-run T on one axis must smoke with extended log; (3) log shows 2 end anchors + expected mid-run stations, ascending, distinct, all accepted stations `referenceNormalAlongAxis=true`; (4) case 14 shows no phantom station; (5) reversed axis + reversed side pick give invariant physical set; (6) operator accepts before a code session opens.

**No-code decision:** code prohibited (spike-code and production-port both) until the Section 11 log-only experiment runs in Revit and passes the gates. Next valid step = extend Wall Spike into a log-only mode on the mid-run + nonjoined-proximity fixture; NOT writing aggregation.

**Production drift note:** production `QuickDimensionWallCandidateCollector.CollectSelectedWallEndAnchors` still uses the superseded min/max planar-face wall-end model (ADR-2026-07-12), not the spike side-face directional resolver — so production has nothing to aggregate yet.
