# ArcTool — Locked Technical Decisions

Last updated: 2026-08-07
Status: Durable decision register. Extracted verbatim from `CLAUDE.md` section 5 during the change-5 thin-core split; `CLAUDE.md` now keeps only a pointer.

**Read this file when** you are about to change platform-level conventions, Excel-to-Revit sync behavior, the Coordinate pipeline, or Quick Dimension collectors/creation/audit. Skip it for local single-file edits with known blast radius.

Every bullet below is a locked decision: do not silently reverse one. If a decision must change, state the reversal explicitly to the user, update this file, and record the reason.

---

## General / platform

- `ElementId.Value` comparisons use `long`; quick filters run before LINQ/slow filters.
- `TransactionGroup` for ArrangeDim = single undo record.
- `InstanceBinding` for shared params when binding many categories.
- `RevitView` alias whenever WinForms is in scope (avoids `CS0104`).
- Model enums carry a prefix (`ExcelViewType`, …) to avoid Revit namespace collisions.
- Multi-agent work package is the default execution model for 3+ file, runtime-investigation, roadmap-phase, audit, and regression work; small local tasks stay direct.

---

## Excel to Revit

Feature dossier: `.Dossier/Detailed Technical Dossier - Excel to Revit.md`.

- JSON settings are written atomically via `.tmp` + `File.Replace`/`File.Move` (same-volume assumption).
- `DateTime.Now` pairs with `File.GetLastWriteTime()` for one local-time basis.
- COM release order is child → parent; never `ReleaseComObject` after COM `Delete()`.
- Legends are created by duplicating an existing legend; Revit exposes no create path.
- Native runtime probing for `pdfium.dll` / `libSkiaSharp.dll` is mandatory; deployment must carry them.
- Mapping state mutates only after a successful commit, to prevent in-memory/JSON drift.
- Root guardrails kept short in `CLAUDE.md`: UI-only command entry, sync-engine-owned transactions, atomic JSON write, local-time file drift comparison, explicit COM release order, native dependency probing, post-commit mapping mutation only, legend duplication workaround, `RevitView` alias when WinForms is in scope.

---

## Coordinate

Feature dossier: `.Dossier/Detailed Technical Dossier - Coordinate Feature.md`.

### Locked decisions

- Scope expands only after deterministic extraction/writeback rules are locked per category.
- Extraction contract distinguishes `Vertical`, `Slanted`, and `Unsupported`; vertical point = `LocationPoint.Point`, slanted V1 point = `LocationCurve` start point.
- Extraction stays unit-neutral and returns raw internal feet; conversion is responsible for normalization/output.
- Shared params stay numeric only: `AT_CoordX`, `AT_CoordY`, `AT_CoordZ`.
- Conversion normalizes through millimeters, rounds canonical values to 3 decimals, then rounds final parameter writeback to 4 decimals.
- Conversion uses `ProjectLocation.GetProjectPosition(XYZ)` as the primary path.
- Axis mapping is explicit via `CoordAxisMapping`; persisted key `VN-2000`, code enum `VN2000`.
- Output unit lives in Project Information via `AT_CoordUnit`; values stay `SpecTypeId.Number` and are converted before `Set(double)`.
- Registered categories are the runtime source of truth for batch execution and updater trigger registration; trigger settings do not narrow execution to one active filter.
- Detail Items require both category registration and the RVT-adjacent JSON type-name allowlist.
- `ConvertedCoordinate` stays a sealed record in the conversion layer.
- Updater lifecycle is document-scoped; `UpdaterId` = stable `AddInId` + updater GUID.
- `App.cs` must capture `UIControlledApplication.ActiveAddInId` in `OnStartup()`; event-sender casting is not reliable.
- Triggers use one category-only `ElementCategoryFilter` per registered category; broader trigger hits are filtered again by extraction/writeback rules.
- `CoordinateUpdater.Execute()` never opens a transaction and re-reads Project Information settings every execution.

### Non-regression invariants (feature is closed; these must survive any future edit)

- Registration split: `Register Element Type` = `Structural Columns` + `Structural Foundations`; `Register Detail Type` = `Detail Items` + RVT-adjacent JSON type-name allowlist.
- Extraction: `LocationPoint.Point` for vertical instances; `LocationCurve` start point for slanted column/foundation instances; `LocationPoint` only for registered Detail Items.
- Project Information settings: `AT_CoordAxisMapping`, `AT_CoordUnit`, `AT_CoordTriggerFilter`; defaults `VN-2000`, `Meters`, `StructuralColumns`; updater re-reads them every execution.
- Registered categories, not one active UI filter, define runtime scope for `Write Coordinates` and Auto Update; writeback converts to `AT_CoordUnit`, rounds to 4 decimals, skips unchanged values.
- Auto Update is document-scoped with one `Element.GetChangeTypeGeometry()` trigger per registered category; `IUpdater.Execute()` opens no transaction and clears static `_isUpdating` in `finally`.
- `App.cs` must capture `application.ActiveAddInId` in `OnStartup()` and pass the stored field to event handlers; casting document-event `sender` to `ControlledApplication` silently breaks registration.
- Element binding stays in `CoordinateParameterBindingService`; Project Information binding stays in `CoordinateProjectSettingsService`.

---

## Quick Dimension

Roadmap: `.Dossier/Quick Dimension - Implementation Roadmap.md`. Deferred track: `.Dossier/Quick Dimension - Deferred Rollback Validation Task.md`. Audit evidence: `Memory/project_qd_chain_creation_audit_handoff.md`.

- Geometry service stays transaction-free and document-free; collectors own document access and diagnostics.
- Legacy collectors use true 2D segment/dimension-line intersection; Grid = `new Reference(grid)` with no arc-grid support; Wall = `HostObjectUtils.GetSideFaces()` with closest valid side face.
- Main flow is WALL-AXIS PROJECTION, not cross-cutting intersection: one selected straight host wall, one picked side, no bulk multi-wall mode.
- Axis = selected host wall `LocationCurve` (Line only); participation = `QuickDimensionLineContext.ProjectParameter` within `[0, Length]`.
- Each opening contributes BOTH left and right jambs along the wall direction; named `Left/Right` stationing must derive from that same reference's own geometry.
- `QuickDimensionCandidate.ParameterOnDimensionLine` means projected coordinate on the wall axis.
- Wall-end anchors are min/max projected stations among wall-direction-aligned planar faces, not raw `LocationCurve` endpoints.
- Chain readiness requires distinct projected stations; duplicates are dropped with `DuplicateStation` diagnostics.
- Read-only summary renders values in millimeters; internal math stays in feet.
- Wall Spike resolver remains directional per shell; full-height threshold counts only candidates with `Reference != null`; spike behavior is not yet ported to the production wall collector.
- Per-joint left/right anchor correctness is required input evidence but not sufficient chain acceptance evidence.
- Mid-run wall-joint detection uses side-line reference evidence, not join APIs.
- `NewDimension` must span the resolved final candidate range, not raw selected-wall `0..axisLength`.
- Post-commit audit is read-only, failure-isolated, and sequence-strict: append `<ChainCreationAudit>` atomically, accept only exact/complete-reversed stable-reference order, compare segment values to station deltas, and do not whitelist local pair swaps.
- Audit treats stable identity, live-reference owner preservation, and candidate metadata ownership as separate checks.
- Forced post-transaction rollback validation is a deferred verification track, not an acceptance gate for collector/candidate-metadata/audit changes.
