---
name: project_qd_phase26_smoke_fix
description: Superseded Phase 2.6 smoke-test findings; preserve only chain-readiness and mm-formatting rules while Wall Spike replaces the wall-end resolver.
type: project
---

Phase 2.6 first real Revit 2026 smoke run of the [[project_qd_projection_pivot]] wall-axis projection path (2026-07-12) validated Door/Window collection but exposed three defects that are now fixed in-tree pending re-smoke.

**Supersession — 2026-07-17:**
- The wall-end min/max planar-face rule below is superseded for active Wall Spike and future production porting by the accepted side-face boundary + directional full-height resolver in [[project_qd_wall_spike_handoff]] / ADR-2026-07-17B. Do not use this memory to reintroduce min/max planar-face wall-end selection.
- The remaining chain-readiness and millimeter-formatting rules stay active, but they do not prove mixed L-joint/T-joint aggregation. ADR-2026-07-17C requires one-axis aggregation research and self-critique before code.

**Historical locked outcomes (wall-anchor part superseded):**
- Wall-end anchor collection must use physical wall solid end caps, not `LocationCurve` centerline endpoints. Collect wall-direction-aligned planar faces with direct `QuickDimensionLineContext.ProjectParameter`, then pick the min and max projected stations as the two solid caps. Face normal alignment alone is insufficient because opening reveal/jamb faces also point along the wall axis, but those opening faces sit between the physical solid caps.
- `QuickDimensionOptions.WallEndStationTolerance` remains useful only to reject degenerate min/max caps that are not distinct enough; it is no longer a "must be near station 0/Length" filter because wall joins can put the visible wall corner away from the LocationCurve endpoint by half the joining wall thickness.
- `QuickDimensionReadOnlyResult.CanCreateChainDimension` requires ≥2 candidates AND distinct projected stations within `DuplicateTolerance`. The engine performs a global projected-station dedupe after source-aware dedupe; each collision emits a `QuickDimensionRejectedReason.DuplicateStation` diagnostic naming the kept candidate.
- `QuickDimensionReadOnlySummaryCommand` renders wall-axis length and each ordered candidate `t` in millimeters via `UnitUtils.ConvertFromInternalUnits(value, UnitTypeId.Millimeters)`. Internal math (`ProjectParameter`, `Length`, tolerances) remains in Revit internal feet.

**Why:** the first smoke report labelled opening jamb faces as `Wall [Start End]` and `Wall [Finish End]`; the next smoke showed wall end references were missing because real physical caps at joined walls can be offset from LocationCurve station 0/Length (e.g. by half the joining wall width). `CandidateCount >= 2` also trusted collided stations, and feet output blocked user validation. These defects would have propagated into Phase 3 chain-dimension creation.

**How to apply:** in future edits, do not anchor wall ends to `LocationCurve` endpoints and do not treat "face normal aligned with wall direction" as sufficient wall-end evidence. For straight non-curtain wall MVP, use min/max projected station among wall-direction-aligned planar faces as the physical solid caps. Do not treat `CandidateCount >= 2` as chain readiness — enforce distinct stations. Render any Quick Dimension user-visible measurement in millimeters. See ADR-2026-07-12 entries in `codebase-memory-mcp` for the locked decision text.
