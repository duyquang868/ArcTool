---
name: project_qd_projection_pivot
description: Quick Dimension pivot from cross-cutting intersection to wall-axis projection model — scope, decisions, and edit plan
type: project
---

Quick Dimension main flow pivots from picked-two-point cross-cutting INTERSECTION to WALL-AXIS PROJECTION (decided 2026-06-11, ADR-2026-06-11).

**Model:** user selects ONE host Wall (its straight LocationCurve = dimension axis, even if skewed) + picks a side (left/right). Engine gathers references ONLY from that wall: its 2 end edges + every hosted Door/Window opening, each opening giving BOTH left+right jambs. Project each reference onto the wall axis (`QuickDimensionLineContext.ProjectParameter`), keep when within `[0, Length]`. No drawn line. Jamb points built along WALL direction (mixing wall dir with drawn-line dir was the window=0 root cause).

**Why:** 2026-06-11 Revit smoke test showed intersection model returns structurally wrong candidates for "dimension along a wall" intent (Window 17→0, Door 1 jamb, parallel walls rejected).

**How to apply (edit plan, decided with user):**
- Product scope is permanently manual and reviewable: exactly one operator-selected straight wall plus one side pick per invocation produces exactly one chain on that wall axis. Never add bulk/automatic multi-wall dimension creation; high-volume geometry-dependent output cannot be validated reliably by a human.
- Wall Spike remains open after its isolated 100% L/T pass: per-joint left/right anchor correctness is only an input invariant. Before code or production porting, define and self-critique a mixed L/T one-axis aggregation contract (two end anchors plus relevant intermediate joint stations; ascending order; explicit duplicate-station diagnostics) against L-L, T-T, L-T, T-L, reversed-axis, and coincident-station counterexamples.
- Scope THIS session = contract + engine layer ONLY. Do NOT edit `QuickDimensionReadOnlySummaryCommand` yet.
- Grid collector + intersection helpers (`TryIntersectSegmentWithDimensionLine2D`, `IsNearlyParallel`/`ParallelToDimensionLine` guard): KEEP in source, remove from main flow. Do not delete.
- Be ADDITIVE: add new APIs (e.g. wall-axis context factory, projection-based collect methods) alongside existing signatures so the unchanged command still compiles. Old picked-2-point `QuickDimensionLineContext.Create` stays.
- Phase 1 spike files (`QuickDimension*ReferenceProbe*`, `*ReferenceSpikeCommand`) are UNTOUCHED — they have own models + `CreateDimensionLine`, no dependency on `QuickDimensionLineContext`.
- `QuickDimensionCandidate.ParameterOnDimensionLine` reinterpreted as projected coord on wall axis.
- Build verification: `dotnet` unavailable in this shell; build/load/smoke must run in Windows/Revit dev env.

**Dependency map (production chain that changes):** QuickDimensionContract.cs → QuickDimensionGeometryService.cs → Wall+DoorWindow collectors → QuickDimensionReadOnlyEngine.cs → QuickDimensionReadOnlySummaryCommand.cs (command NOT edited this session).
