# ARCTOOL — TECHNICAL CONTEXT
Last updated: 2026-05-19 — Coordinate Phase D dynamic updater workflow is implemented, build-confirmed, and Revit-tested; next priority is Phase E compact operator UI and Phase F QA.

---

## Mandatory editing rules
- Preserve 100% of the file structure, numbering, and headings.
- Only add and update in place; never delete existing content.
- Never rewrite the file from scratch; edit only the exact lines that need changes.
- Keep updates clear, short, and information-dense to reduce token load while preserving full technical meaning.
- All content written inside `CLAUDE.md` must be in English.

## 1. Project snapshot

| Item | Value |
|---|---|
| Project | ArcTool |
| Main namespace | `ArcTool.Core` |
| Platform | Autodesk Revit 2026 API |
| Language | C# / .NET 8.0 |
| UI | WPF + limited WinForms |
| Units | `UnitTypeId` only; do not use deprecated `DisplayUnitType` |

---

## 2. Code map

```text
ArcTool/
├── ArcTool.Core/
│   ├── App.cs
│   ├── Commands/
│   │   ├── CreateVoidFromLinkCommand.cs
│   │   ├── MultiCutCommand.cs
│   │   ├── ArrangeDimensionCommand.cs
│   │   ├── FilterManagerCommand.cs
│   │   ├── ExcelToRevitCommand.cs
│   │   ├── RegisterCoordParamsCommand.cs
│   │   └── RunCoordBatchCommand.cs
│   ├── Services/
│   │   ├── ExcelInteropService.cs
│   │   ├── ArcToolSettingsService.cs
│   │   ├── ExcelSyncEngine.cs
│   │   ├── CoordinateExtractionService.cs
│   │   ├── CoordinateConversionService.cs
│   │   ├── CoordinateBatchService.cs
│   │   ├── CoordinateProjectSettingsService.cs
│   │   ├── CoordinateUpdater.cs
│   │   └── CoordinateUpdaterService.cs
│   ├── UI/
│   │   ├── FilterWindow.xaml
│   │   ├── FilterWindow.xaml.cs
│   │   ├── ExcelToRevitWindow.xaml
│   │   └── ExcelToRevitWindow.xaml.cs
│   ├── Models/
│   │   ├── CoordinateContract.cs
│   │   └── ExcelMapping.cs
│   ├── Utilities/
│   │   └── SelectionFilters.cs
│   └── Properties/
│       ├── Resources.resx
│       └── Resources.Designer.cs
```

---

## 3. Current technical state

### Stable features
- `App.cs`: Ribbon bootstrapping is stable; `Coordinate Tools` panel includes `RegisterCoordParamsCommand` and `RunCoordBatchCommand`; document open/create/closing events now register/unregister the coordinate updater using `UIControlledApplication.ActiveAddInId` captured in `OnStartup()`.
- `CreateVoidFromLinkCommand.cs`: stable, but still carries face-based host fragility on link geometry changes.
- `MultiCutCommand.cs`: stable broad-phase cut workflow using `BoundingBoxIntersectsFilter` + `InstanceVoidCutUtils`.
- `ArrangeDimensionCommand.cs`: stable spacing workflow using `TransactionGroup`.
- Excel to Revit stack is complete and considered closed for active development:
  - `ExcelToRevitCommand.cs`
  - `ExcelInteropService.cs`
  - `ArcToolSettingsService.cs`
  - `ExcelSyncEngine.cs`
  - `ExcelToRevitWindow.xaml/.cs`
  - `ExcelMapping.cs`
- Coordinate Phase A contract/bootstrap is implemented and validated in Revit:
  - `CoordinateContract.cs`
  - `RegisterCoordParamsCommand.cs`
  - Current command scope is registration/binding only; no extraction or writeback logic exists yet.
- Coordinate Phase B engine layer is implemented and build-confirmed in VS Code:
  - `CoordinateExtractionService.cs`
  - `CoordinateConversionService.cs`
  - Phase B has no UI, command, or parameter writeback; runtime validation depends on Phase C calling the services.
- Coordinate Phase C run-once batch layer is implemented, build-confirmed, and Revit-tested:
  - `CoordinateBatchService.cs`
  - `RunCoordBatchCommand.cs`
  - `CoordinateProjectSettingsService.cs`
  - Project Information settings: `AT_CoordAxisMapping`, `AT_CoordUnit`.
  - Default settings: `AT_CoordAxisMapping = VN-2000`, `AT_CoordUnit = Meters`.
  - Batch write applies axis mapping, converts all X/Y/Z values to `AT_CoordUnit`, rounds to 4 decimals, skips unchanged values, and reports unsupported/failed elements.
- Coordinate Phase D dynamic updater layer is implemented, build-confirmed, and Revit-tested:
  - `CoordinateUpdater.cs`
  - `CoordinateUpdaterService.cs`
  - `App.cs` document lifecycle wiring.
  - Document-scoped updater registration runs on `DocumentOpened` and `DocumentCreated`; unregister runs on `DocumentClosing`.
  - `App.cs` must capture `application.ActiveAddInId` in `OnStartup()` and pass the stored field to event handlers; journal test proved casting event `sender` to `ControlledApplication` returned null and skipped registration.
  - Updater processes only modified `OST_StructuralColumns` `FamilyInstance` elements, uses `Element.GetChangeTypeGeometry()`, applies the same output-unit conversion/rounding path as `CoordinateBatchService`, and writes no values when normalized values are unchanged.
  - `IUpdater.Execute()` opens no transaction and writes inside Revit's active updater transaction; `_isUpdating` is a static reentrance guard cleared in `finally`.
  - Runtime test evidence: after fixing `App.cs` AddInId handling, user build and in-Revit move/modify test succeeded; coordinates auto-updated without running the batch command.

### Incomplete feature
- `FilterManagerCommand.cs` + `FilterWindow.xaml/.cs`: UI skeleton exists; actual `ParameterFilterElement` copy/paste logic is not implemented.
- Coordinate feature after Phase D remains incomplete: operator UI, stress testing, undo/worksharing QA, deployment packaging, and optional schedule-validation workflow are not implemented yet.

---

## 4. Open bugs worth remembering

| ID | Area | Problem | Priority |
|---|---|---|---|
| BUG-06 | ArrangeDimension | Missing guard for `activeView.Scale == 0` / unsupported view contexts | Medium |
| BUG-07 | FilterManager | `Idling`-based refresh architecture does not scale on large models | Low |
| BUG-08 | CreateVoidFromLink | `SetParam("Height", -beamHeight)` is still a workaround, not a clean model | Low |

Only keep bug history here when it can affect future fixes. Resolved Excel bugs are intentionally removed from this top-level file.

---

## 5. Technical decisions already locked

| Decision | Reason | Trade-off |
|---|---|---|
| Use `long` for `ElementId.Value` comparisons | Avoid overflow bugs | Slightly noisier code |
| Quick filters before LINQ/slow filters | Revit collector performance | None |
| `TransactionGroup` for ArrangeDim | Single undo record | None |
| `InstanceBinding` for shared params when binding many categories | Batch setup is cheaper and cleaner | Requires scope discipline |
| `RevitView` alias when WinForms is in play | Avoid `CS0104` with `System.Windows.Forms.View` | Slightly longer type names |
| Enum prefix in models (`ExcelViewType`, etc.) | Avoid Revit namespace collisions | Longer identifiers |
| JSON atomic write via `.tmp` + `File.Replace` / `File.Move` | Prevent corrupt settings files | Same-volume assumption |
| `DateTime.Now` with `File.GetLastWriteTime()` | Same local-time basis | Timezone changes can skew comparisons |
| COM release order child → parent | Prevent RCW lifetime bugs | Requires discipline |
| Never `ReleaseComObject` after COM `Delete()` | Avoid undefined behavior | Must trust delete semantics |
| Legend creation via duplicate, not create | Revit API has no public create path for legends | Requires existing legend template |
| Excel native runtime probing (`pdfium.dll`, `libSkiaSharp.dll`) | Revit add-in loading is not normal app loading | Deployment must carry native libs |
| Excel mapping mutation after successful commit | Prevent in-memory/JSON drift when commit fails | Requires temporary locals |
| Coordinate V1 scope stays `OST_StructuralColumns` only | Prevent premature multi-category drift before deterministic rules exist | Later expansion needs an explicit V2 contract |
| Coordinate extraction contract distinguishes `Vertical`, `Slanted`, and `Unsupported` | Wrong point definition breaks all downstream coordinate workflows | Slightly more model ceremony up front |
| Vertical column point = `LocationPoint.Point` | Stable explicit rule for point-based structural columns | Ignores any future alternate placement interpretations |
| Slanted column V1 point = `LocationCurve` start point (index 0) | Locks one deterministic rule and avoids midpoint ambiguity | May need future project-specific revision |
| Coordinate shared params stay numeric only: `AT_CoordX`, `AT_CoordY`, `AT_CoordZ` | Formatting must not pollute the source of truth used by schedules/updaters | Debug visibility depends on external diagnostics |
| Coordinate conversion service still normalizes through millimeters before output-unit conversion | Phase B keeps one canonical conversion/rounding basis | Phase C converts to `AT_CoordUnit` before writing numeric parameters |
| Coordinate conversion rounding is 3 decimals in canonical millimeters, then parameter writeback rounds final output to 4 decimals | Prevent future updater chattering while matching operator schedule precision | Comparisons must normalize through the same conversion/writeback rounding path |
| Coordinate extraction stays unit-neutral and returns raw internal feet | Preserve a clean boundary between geometry extraction and storage conversion | Callers must not treat `CoordResult` values as millimeters |
| Coordinate conversion uses `ProjectLocation.GetProjectPosition(XYZ)` as primary path | Respect active project/shared coordinate setup instead of hand-rolled transform inversion | Phase C must validate project-location setup on real models |
| Coordinate axis mapping is explicit via `CoordAxisMapping` | International/project coordinate conventions are not universal core math | Project Information stores the setting key; code enum uses `VN2000` while persisted key is `VN-2000` |
| Coordinate parameter output unit is stored in Project Information via `AT_CoordUnit` | Unit convention must travel with the `.rvt`, not a sidecar JSON file | `AT_CoordX/Y/Z` remain `SpecTypeId.Number`; values are converted by code and rounded to 4 decimals before `Set(double)` |
| `ConvertedCoordinate` is a sealed record in the conversion layer | Immutable value semantics fit storage-ready conversion output | Axis-mapped labels reuse the same EastWest/NorthSouth property names for traceability |
| Coordinate updater registration is document-scoped | Prevent updater triggers from leaking across documents | Register on document open/create and unregister on document closing |
| Coordinate updater `UpdaterId` uses stable `AddInId` + updater GUID | Revit identifies updater persistence and trigger registration by this pair | GUID must remain stable after deployment |
| `App.cs` stores `UIControlledApplication.ActiveAddInId` during `OnStartup()` | Revit journal proved document event `sender` casting can produce null AddInId and skip registration | Keep `_addInId` as the event-handler source of truth |
| Coordinate updater trigger uses category + class `LogicalAndFilter` | Category-only triggers can include column type elements instead of only instances | Slightly more verbose trigger setup |
| `CoordinateUpdater.Execute()` never opens a transaction | Revit runs updater execution inside an active transaction | Parameter writes join the user's undo unit |
| Coordinate updater reads Project Information settings on every execution | Axis/unit settings can change during a Revit session | Small per-execution read cost |

---

## 6. Active roadmap

Priority is now shifted away from Excel to Revit. That feature is closed unless a real bug appears. New development priority is Coordinate Phase E compact operator UI and Phase F QA, then Filter Manager, then optional R&D.

### A. Coordinate feature — next active priority

#### Scope lock
V1 must target `Structural Columns` only. Do not start with “all 3D categories”. Generic expansion comes later.

#### Core principle
The real problem is not reading coordinates from Revit; it is defining which coordinate is correct for each supported element and project context. Wrong definition means the whole updater/UI/schedule chain is wrong.

#### Development phases

**Phase A — Scope & data contract**
- Status: completed and validated manually in Revit on 2026-05-17.
- Implemented files: `CoordinateContract.cs`, `RegisterCoordParamsCommand.cs`.
- Ribbon entry: `Coordinate Tools` → `Register Coord Params`.
- Session A1: locked V1 scope to `Structural Columns` only via `BuiltInCategory.OST_StructuralColumns`.
- Session A2: point rules are now locked.
  - Vertical column: `LocationPoint.Point`.
  - Slanted column V1: `LocationCurve` start point (index 0).
  - Unsupported geometry must return explicit unsupported state; never guess.
- Session A3: shared parameter contract is now locked.
  - Minimum and current set: `AT_CoordX`, `AT_CoordY`, `AT_CoordZ`.
  - No debug/meta parameter was added in V1 Phase A.
  - Registration command uses shared parameter group `ArcTool_Coordinates`.
  - Parameter type is `SpecTypeId.Number`; do not use `SpecTypeId.Length`.
  - Binding is `InstanceBinding` to `OST_StructuralColumns` only.
  - Registration command is idempotent; repeated runs must not create duplicate bindings.
- Session A4: storage form is now locked.
  - Persist raw numeric values only.
  - Phase B canonical conversion unit remains millimeters via `UnitTypeId.Millimeters`; Phase C converts to the Project Information `AT_CoordUnit` before parameter writeback.
  - Storage rounding policy: 3 decimals, canonicalized before future write/compare steps.
  - Formatting belongs to schedule/UI, not core math.
- Important boundary: Phase A does not implement extraction, coordinate conversion, VN-2000 mapping, or parameter writeback.

**Phase B — Coordinate engine**
- Status: implemented and build-confirmed in VS Code on 2026-05-17.
- Implemented files: `CoordinateExtractionService.cs`, `CoordinateConversionService.cs`.
- Session B1: built `CoordinateExtractionService`.
  - Supports `LocationPoint`.
  - Supports `LocationCurve`.
  - Returns explicit unsupported state instead of guessing.
  - `ExtractAll(Document)` uses quick filters: `OfClass(typeof(FamilyInstance))` then `OfCategory(BuiltInCategory.OST_StructuralColumns)`.
  - No unit conversion exists in extraction; `CoordResult` remains raw Revit internal feet.
- Session B2: built `CoordinateConversionService`.
  - Uses `ProjectLocation.GetProjectPosition(XYZ)` as the primary active-document coordinate conversion path.
  - Does not use `GetTotalTransform().Inverse` as the primary active-document algorithm.
  - Converts `ProjectPosition.EastWest`, `NorthSouth`, and `Elevation` from internal feet to millimeters via `UnitTypeId.Millimeters`.
- Session B3: unit conversion and rounding policy are implemented through `CoordStoragePolicy.RoundForStorage(...)`.
- Session B4: VN-2000 axis swap is implemented only as explicit `CoordAxisMapping.VN2000`; no core hardcoded X/Y swap exists.
- Session B5: full coordinate value validation is deferred to Phase C because Phase B has no visible command, UI, or parameter writeback.
- Revit smoke evidence: running `RegisterCoordParamsCommand` after the Phase B build did not crash or conflict; journal only validates Phase A command execution, not Phase B service runtime behavior.

**Phase C — Stable batch workflow**
- Status: completed, build-confirmed, and Revit-tested on 2026-05-19.
- Implemented files: `CoordinateBatchService.cs`, `RunCoordBatchCommand.cs`, `CoordinateProjectSettingsService.cs`.
- Session C1: deterministic `Write Coordinates` run-once command is implemented on the Coordinate Tools ribbon panel.
- Session C2: batch write stores coordinates into `AT_CoordX`, `AT_CoordY`, `AT_CoordZ` using `Parameter.Set(double)` on numeric instance parameters.
- Session C3: project-owned settings live in Project Information parameters `AT_CoordAxisMapping` and `AT_CoordUnit`; defaults are `VN-2000` and `Meters`.
- Session C4: skip-check normalizes existing values through the same unit conversion and 4-decimal writeback precision before comparing.
- Session C5: unsupported and failed elements are surfaced in the post-run TaskDialog with detail rows capped at 20 lines.
- Revit test evidence: user build and in-Revit run succeeded; expected output for VN-2000 + Meters is `CoordX = NorthSouth`, `CoordY = EastWest`, all X/Y/Z rounded to 4 decimals.

**Phase D — Dynamic update**
- Status: completed, build-confirmed, and Revit-tested on 2026-05-19.
- Implemented files: `CoordinateUpdater.cs`, `CoordinateUpdaterService.cs`; narrow diff in `App.cs`.
- Session D1: stable Phase C engine is wrapped by `IUpdater` without opening a nested transaction.
- Session D2: trigger scope is narrow: document-scoped `UpdaterRegistry.RegisterUpdater(..., doc, false)` plus `LogicalAndFilter` for `OST_StructuralColumns` and `FamilyInstance`, with `Element.GetChangeTypeGeometry()`.
- Session D3: re-entry/chattering is controlled by static `_isUpdating` cleared in `finally`, plus normalized same-path skip-check before `Set(double)`.
  - Delta check alone is not enough.
  - Canonicalize units/rounding before compare.
  - Write only when normalized value changes.
- Session D4: real Revit smoke test succeeded after fixing `App.cs` AddInId lifecycle.
  - Journal failure evidence: `AddInId is null — updater registration skipped` occurred when event handlers tried to cast `sender` to `ControlledApplication`.
  - Final fix: capture `application.ActiveAddInId` into `_addInId` during `OnStartup()` and pass `_addInId` to document lifecycle handlers.
  - Runtime behavior: after rebuild/copy and Revit reopen, moving/modifying Structural Columns auto-updates `AT_CoordX/Y/Z` without running `Write Coordinates` manually.
- Remaining Phase D-adjacent work: profile latency on large real models during Phase F stress testing.

**Phase E — Operator UI**
- Session E1: keep the UI minimal, not a large dashboard.
  - `Write Coordinates` remains the operator-facing coordinate control on the ribbon.
  - Main button click runs the existing manual batch write.
  - A small ON/OFF toggle attached to the same control enables or disables auto update for the current document.
  - UI should show no success/failure popups; operational detail belongs in the ArcTool log file.
- Session E2: `Register Coord Params` must open a small settings dialog instead of running as a silent one-shot setup command.
  - Dialog fields: `Axis Mapping`, `Output Unit`, `Trigger Filter`.
  - All three fields must be dropdown lists, not free-text inputs.
  - Each option must preserve three layers: user-facing label, code-safe enum/key, and persisted Project Information/config key.
  - User-facing labels and persisted keys should keep internationally recognizable strings such as `VN-2000`; code enums must keep the codebase-safe form such as `VN2000`.
  - Dialog actions stay minimal: `OK` and `Cancel` only.
- Session E3: do not expose arbitrary free-form trigger strings.
  - `Trigger Filter` must be limited to predefined supported scopes only.
  - A scope can appear in the dropdown only after its extraction/writeback rules are implemented and validated.
  - This prevents the UI from implying multi-category support before backend rules are real.
- Session E4: move detailed diagnostics out of the UI and into a dedicated ArcTool log file.
  - Log should carry updater registration state, AddInId acquisition, selected settings, trigger scope, processed/written/skipped/unsupported/failed counts, and first meaningful exception text.
  - The UI is only for compact control and state indication; the log is the support/debug source of truth.

**Phase F — QA & deployment**
- Session F1: stress test at 100 / 1,000 / 5,000 / 10,000 elements.
- Session F2: test undo behavior.
- Session F3: test worksharing / local-central scenarios if relevant.
- Session F4: package installer.

#### Explicit constraints
- Do not rebuild updater logic before preserving the Revit-tested Phase D behavior.
- Do not build generic multi-category support in V1.
- Do not hardcode X/Y swap as core math.
- Do not store formatted display strings as the main source of truth.

### B. Filter Manager — secondary priority
- Replace skeleton behavior with real `ParameterFilterElement` copy/paste logic.
- Remove or redesign `Idling` refresh if it becomes the source of model-scale lag.
- Only keep MVVM complexity that directly supports filtering workflow.

### C. Optional R&D
- Quick Dim / `ReferenceArray` extraction from Wall / Column / Beam.
- No production commitment until a deterministic geometry strategy exists.

---

## 7. Closed technical dossier — recent closure record

This section is reserved only for features that were closed recently enough to remain useful in top-level context. Older closed features must live in dedicated detailed dossier files under `.Dossier` and should not stay here indefinitely.

`.Dossier` stores only deep technical dossiers for individual features or subsystems. Do not put temporary notes, short-lived TODOs, or chat-session analysis reports there. Every dossier file name and every dossier file body must be written in English.

When a feature is closed:
- create or update one dedicated dossier file under `.Dossier`;
- keep one dossier per feature or clearly bounded subsystem;
- leave only a short summary and pointer here;
- remove the section-7 summary later when the closure is no longer recent.

### Current recent closure
- Excel to Revit
  - Detailed dossier: `.Dossier/Detailed Technical Dossier - Excel to Revit.md`
  - Components: `ExcelToRevitCommand.cs`, `ExcelInteropService.cs`, `ArcToolSettingsService.cs`, `ExcelSyncEngine.cs`, `ExcelToRevitWindow.xaml/.cs`, `ExcelMapping.cs`
  - Keep in mind: command opens UI only; transaction work stays inside sync engine; JSON settings live next to `.rvt`; JSON write is atomic; `LastModified` uses local time; COM wrappers like `Sheets` / `Names` must be released explicitly; native runtime probing for `pdfium.dll` and `libSkiaSharp.dll` is mandatory; `ExcelSyncEngine` mutates mapping only after commit succeeds; `using RevitView = Autodesk.Revit.DB.View` remains mandatory in the sync engine path; legend creation still depends on duplicating an existing legend template; `_suppressRowEvents` must be set before programmatic row mutation in the WPF window.

When a closed feature is no longer recent, remove its summary from this section and keep only its dedicated dossier file.

---

## 8. Coding rules

```csharp
// 1. Every command stays manual.
[Transaction(TransactionMode.Manual)]

// 2. Alias Revit UI / View types when WinForms namespaces are present.
using RevitTaskDialog = Autodesk.Revit.UI.TaskDialog;
using RevitView = Autodesk.Revit.DB.View;

// 3. Never cast ElementId.Value to int.
(long)BuiltInCategory.OST_Walls

// 4. Quick filters before slow filters.
new FilteredElementCollector(doc)
    .OfClass(typeof(Wall))
    .OfCategory(BuiltInCategory.OST_Walls)
    .Where(...);

// 5. Read mutable element state before delete.
if (existingInst?.IsValidObject == true)
{
    double w = existingInst.Width;
    doc.Delete(existingInst.Id);
}

// 6. JSON settings must use the atomic service, not raw File.WriteAllText.
ArcToolSettingsService.SaveMappings(doc, mappings);

// 7. Time comparisons use local time consistently.
mapping.LastModified = DateTime.Now;

// 8. Dispose Excel interop immediately.
using (var svc = new ExcelInteropService())
{
    svc.OpenFile(path);
    var names = svc.GetSheetNames();
}

// 9. COM release order is child -> parent.
Marshal.ReleaseComObject(child);
Marshal.ReleaseComObject(parent);

// 10. Do not hold long-lived COM references to ActiveSheet-like objects unless necessary.

// 11. For updater-style features, compare normalized values before Set().

// 12. For new model enums, prefix names to avoid DB namespace collisions.
```

---

## 9. API references worth remembering

| API | Why it matters |
|---|---|
| `FilteredElementCollector` | Primary collection pipeline; performance depends on filter order |
| `LocationPoint` | Point-based placement for supported elements |
| `LocationCurve` | Curve-based placement for slanted/linear elements |
| `ProjectLocation.GetProjectPosition(...)` | Safer path for project/shared coordinate conversion |
| `IUpdater` | Dynamic coordinate updater execution; `Execute()` runs inside Revit's active transaction |
| `UpdaterRegistry.RegisterUpdater(...)` / `AddTrigger(...)` / `UnregisterUpdater(...)` | Document-scoped coordinate updater lifecycle |
| `Element.GetChangeTypeGeometry()` | Geometry/location trigger used by the coordinate updater |
| `UIControlledApplication.ActiveAddInId` | Reliable source for updater `AddInId`; do not derive it from document event sender |
| `BasePoint.GetSurveyPoint(Document)` | Survey coordinate context |
| `BasePoint.GetProjectBasePoint(Document)` | Project base point context |
| `ParameterFilterElement` | Filter Manager implementation target |
| `Application.SharedParametersFilename` | Guard shared-parameter registration before trying to open or create definitions |
| `Application.OpenSharedParameterFile()` | Required entry point for coordinate shared-parameter registration |
| `ExternalDefinitionCreationOptions` + `SpecTypeId.Number` | Coordinate parameters must stay raw numeric values, not Revit length specs |
| `DefinitionBindingMapIterator` / `ForwardIterator()` | Required binding-map inspection path for idempotent shared-parameter registration |
| `BindingMap.Insert(...)` / `ReInsert(...)` | Coordinate parameter binding and repair path |
| `ImageType.Create(...)` / `ImageInstance.Create(...)` | Excel image import pipeline |
| `ViewDrafting.Create(...)` | Drafting view creation |
| `View.Duplicate(ViewDuplicateOption.WithDetailing)` | Legend-view workaround |
| `Marshal.ReleaseComObject(...)` | Excel interop hygiene |

Invariant reminder: always use internet search tools to look up Revit API and any other relevant technical documentation before answering or fixing a bug. Eliminate guesswork completely; every technical claim should have clear documentation or evidence. If no reliable information can be found, report that to the user and request human help.

Revit API reference of record: https://www.revitapidocs.com/2026/

---

## 10. Editing policy for this file
This file is a technical operating document, not a narrative report.
Keep only:
- active roadmap;
- open bugs that still matter;
- locked technical decisions;
- code-level invariants useful for implementation or bug fixing.

Remove:
- long historical walkthroughs;
- UI prose that does not affect code;
- completed implementation diaries;
- business-language explanations that do not help debug or extend the code.
