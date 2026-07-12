# Coordinate Feature — Detailed Technical Dossier

## 1. Scope

The Coordinate feature is a closed ArcTool subsystem for writing project/shared coordinates into raw numeric shared parameters on supported model and annotation elements.

Final supported scope:
- Structural Columns
- Structural Foundations
- Registered Detail Items

Final operator split:
- `Register Element Type` handles supported 3D categories.
- `Register Detail Type` handles Detail Item parameter binding plus RVT-adjacent JSON type-name registration.
- `Write Coordinates` performs manual batch processing.
- `Auto Update` toggles document-scoped real-time processing.

The feature was accepted by the user in Revit and closed on 2026-05-25.

## 2. Core architecture

### Command layer
- `RegisterCoordParamsCommand.cs`
- `RegisterDetailItemCoordTypeCommand.cs`
- `RunCoordBatchCommand.cs`
- `ToggleCoordUpdaterCommand.cs`

### Service layer
- `CoordinateParameterBindingService.cs`
- `CoordinateProjectSettingsService.cs`
- `CoordinateDetailItemRegistryService.cs`
- `CoordinateExtractionService.cs`
- `CoordinateConversionService.cs`
- `CoordinateBatchService.cs`
- `CoordinateUpdater.cs`
- `CoordinateUpdaterService.cs`
- `CoordinateLogService.cs`

### Host wiring
- `App.cs`
- `CoordSettingsDialog.xaml`
- `CoordSettingsDialog.xaml.cs`

## 3. Shared-parameter and settings contract

### Element parameters
The feature writes only these raw numeric shared parameters:
- `AT_CoordX`
- `AT_CoordY`
- `AT_CoordZ`

These remain `SpecTypeId.Number` values. Formatting concerns must stay outside the storage contract.

### Project Information settings
The feature persists model-owned settings in Project Information:
- `AT_CoordAxisMapping`
- `AT_CoordUnit`
- `AT_CoordTriggerFilter`

Default persisted values are:
- `VN-2000`
- `Meters`
- `StructuralColumns`

`AT_CoordTriggerFilter` remains part of the stored settings contract, but final runtime execution is no longer limited to one active trigger scope.

## 4. Registration model

### Element Type registration
`Register Element Type` is the 3D registration path.

Responsibilities:
- ensure the shared parameter file is available through Revit
- create or reuse the ArcTool coordinate definition group
- delegate element coordinate binding to `CoordinateParameterBindingService`
- bind `AT_CoordX/Y/Z` to `Structural Columns` and `Structural Foundations`
- delegate Project Information settings binding to `CoordinateProjectSettingsService`
- write accepted settings once inside the registration transaction
- keep registration idempotent
- refresh updater registration after a successful setup change

### Detail Type registration
`Register Detail Type` is the annotation registration path.

Responsibilities:
- require the RVT file to be saved first
- require the Revit shared parameter file to exist first
- ask the user to pick one representative Detail Item instance
- validate that the selected instance is a `FamilyInstance` in `OST_DetailComponents`
- validate that the selected instance uses `LocationPoint`
- bind `AT_CoordX/Y/Z` to `Detail Items`
- ensure Project Information coordinate settings exist
- persist the selected Detail Item type name into an RVT-adjacent JSON registry
- refresh updater registration after a successful setup change

The two registration commands are intentionally independent. Detail Item registration must not require the 3D registration command to be run first.

## 5. Runtime scope model

The final runtime source of truth is registration state, not one active UI trigger filter.

This affects both processing paths:
- `Write Coordinates`
- `Auto Update`

Final behavior:
- if Columns, Foundations, and Detail Items are all registered, both runtime paths process all three scopes together
- if only one or two scopes are registered, only those scopes participate
- Detail Items additionally require a matching type name in the RVT-adjacent JSON allowlist

This architecture replaced the earlier one-scope-at-a-time behavior, which caused poor UX because registering multiple scopes did not lead to cumulative processing.

## 6. Extraction rules locked for production

The critical design constraint was not coordinate math alone; it was defining one deterministic representative point rule per supported category and preserving that rule across registration, batch write, updater execution, and operator UX.

Locked rules:
- vertical supported point-based structural instances use `LocationPoint.Point`
- slanted supported structural instances use the `LocationCurve` start point
- registered Detail Items use `LocationPoint` only
- unsupported geometry must surface as unsupported, not guessed

The extraction contract distinguishes `Vertical`, `Slanted`, and `Unsupported` outcomes because downstream coordinate correctness depends on this boundary staying explicit.

## 7. Conversion and writeback model

Extraction remains unit-neutral and returns raw internal Revit feet.

Conversion rules:
- primary conversion path uses `ProjectLocation.GetProjectPosition(XYZ)`
- axis mapping is explicit and model-owned
- output unit comes from `AT_CoordUnit`
- canonical normalization still flows through millimeters before final output-unit conversion
- canonical rounding is 3 decimals in millimeter normalization
- final parameter writeback rounds to 4 decimals

Writeback rules:
- compare normalized values before `Set(double)`
- skip unchanged values to reduce churn
- keep `AT_CoordX/Y/Z` numeric only
- report unsupported or failed elements in batch summaries and logs

## 8. Batch workflow

`RunCoordBatchCommand.cs` is the operator entry point for manual execution.

Final manual workflow:
- resolve current coordinate settings from Project Information
- inspect document registration state
- fail clearly if no coordinate scope is registered
- extract all results across every registered coordinate scope
- convert and normalize coordinates through the shared conversion path
- write only changed values
- report updated, unchanged, unsupported, and failed counts

The command no longer runs against one selected trigger filter only.

## 9. Auto Update workflow

`ToggleCoordUpdaterCommand.cs` enables or disables document-scoped real-time execution.

`CoordinateUpdaterService.cs` owns updater registration.

Locked updater behavior:
- registration is document-scoped
- updater identity uses stable `AddInId` plus updater GUID
- one geometry trigger is added per registered coordinate category
- trigger registration is rebuilt when registration state changes
- runtime processing resolves all registered scopes together instead of one active scope only

The implementation intentionally avoids a combined logical trigger filter for multiple categories. Revit 2026 updater triggers should stay within supported category/parameter filter forms, so the production model is one category-only `ElementCategoryFilter` trigger per registered category. Runtime execution still filters through registered extraction/writeback rules, so broader category triggers do not define final write scope.

## 10. AddInId lifecycle decision

A key production fix was capturing `UIControlledApplication.ActiveAddInId` in `App.OnStartup()` and keeping it as the application-level source of truth.

Reason:
- journal testing proved that deriving `AddInId` from document-event `sender` was unreliable
- a null or unstable `AddInId` breaks updater registration and document lifecycle wiring

`App.cs` must therefore:
- capture `application.ActiveAddInId` during startup
- store it for later command and event-handler use
- register/unregister the updater during document open/create/closing through that stored value

## 11. Detail Item registry model

Detail Item support adds a second persistence boundary besides Project Information:
- an RVT-adjacent JSON registry keyed by Detail Item type name

Why this exists:
- category binding alone is too broad for annotation families
- the operator needs an explicit allowlist chosen from representative instances
- the registry must move with the model file to preserve project intent

Operational consequence:
- copying or moving the RVT without its JSON registry can silently reduce Detail Item processing scope
- deployment and QA must treat the sidecar JSON as part of the feature contract

## 12. UX constraints preserved in the final design

The feature intentionally stays compact in the ribbon and dialog surface.

Locked UX constraints:
- no placeholder settings fields without real backend support
- no forced overlap between 3D registration and Detail Item registration flows
- compact settings dialog with dropdown-only fields
- debug detail belongs in logs, not in an expanded primary UI

These constraints directly shaped the final split between `Register Element Type` and `Register Detail Type`.

## 13. Final acceptance state

Final acceptance evidence:
- registration commands behaved as expected
- manual batch writing behaved as expected
- Auto Update behaved as expected
- cumulative registered-scope behavior worked for Columns, Foundations, and registered Detail Items

The user confirmed in Revit that the full feature behaved as intended. The subsystem is therefore closed. Future work should be framed only as bug fixing, deployment hardening, packaging, release QA, or explicit scope expansion.

## 14. Post-closure cleanup record

A later cleanup pass removed stale logic that was no longer load-bearing after the final all-registered-scope model was accepted.

Cleanup results:
- command-layer duplicate coordinate binding helpers were removed from `RegisterCoordParamsCommand.cs`
- element coordinate binding remains centralized in `CoordinateParameterBindingService.cs`
- Project Information settings binding remains centralized in `CoordinateProjectSettingsService.cs`
- the duplicate standalone settings transaction was removed from `RegisterCoordParamsCommand.cs`
- `RunCoordBatchCommand.cs` no longer contains first-supported-instance probing
- `CoordinateUpdaterService.cs` no longer contains first-supported-instance probing
- updater registration now uses category-only trigger filters built from registered categories

Preserved intentionally:
- similar helper names inside `CoordinateParameterBindingService.cs` and `CoordinateProjectSettingsService.cs` are not duplicate dead code; they serve different binding targets
- `AT_CoordTriggerFilter` remains persisted for settings compatibility, but does not narrow runtime batch/updater scope
- broad category triggers are acceptable because final execution scope is enforced by extraction and registration rules