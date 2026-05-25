# ARCTOOL — TECHNICAL CONTEXT
Last updated: 2026-05-25 — Coordinate registration is split into Element Type and Detail Type pipelines; Write Coordinates and Auto Update now process all registered coordinate scopes instead of one active trigger.

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
│   │   ├── RegisterDetailItemCoordTypeCommand.cs
│   │   ├── RunCoordBatchCommand.cs
│   │   └── ToggleCoordUpdaterCommand.cs
│   ├── Services/
│   │   ├── ExcelInteropService.cs
│   │   ├── ArcToolSettingsService.cs
│   │   ├── ExcelSyncEngine.cs
│   │   ├── CoordinateExtractionService.cs
│   │   ├── CoordinateConversionService.cs
│   │   ├── CoordinateBatchService.cs
│   │   ├── CoordinateProjectSettingsService.cs
│   │   ├── CoordinateUpdater.cs
│   │   ├── CoordinateUpdaterService.cs
│   │   ├── CoordinateLogService.cs
│   │   ├── CoordinateDetailItemRegistryService.cs
│   │   └── CoordinateParameterBindingService.cs
│   ├── UI/
│   │   ├── FilterWindow.xaml
│   │   ├── FilterWindow.xaml.cs
│   │   ├── ExcelToRevitWindow.xaml
│   │   ├── ExcelToRevitWindow.xaml.cs
│   │   └── CoordSettingsDialog.xaml
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
- `App.cs`: Ribbon bootstrapping is stable; `Coordinate Tools` panel includes `RegisterCoordParamsCommand`, `RegisterDetailItemCoordTypeCommand`, `RunCoordBatchCommand`, and `ToggleCoordUpdaterCommand`; document open/create/closing events register/unregister the coordinate updater using `UIControlledApplication.ActiveAddInId` captured in `OnStartup()`.
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
- Coordinate feature is complete, Revit-tested, user-accepted, and closed for active development:
  - Core files: `CoordinateContract.cs`, `CoordinateExtractionService.cs`, `CoordinateConversionService.cs`, `CoordinateBatchService.cs`, `CoordinateProjectSettingsService.cs`, `CoordinateUpdater.cs`, `CoordinateUpdaterService.cs`, `CoordinateLogService.cs`, `CoordinateDetailItemRegistryService.cs`, `CoordinateParameterBindingService.cs`.
  - Command/UI files: `RegisterCoordParamsCommand.cs`, `RegisterDetailItemCoordTypeCommand.cs`, `RunCoordBatchCommand.cs`, `ToggleCoordUpdaterCommand.cs`, `CoordSettingsDialog.xaml`.
  - Registration pipeline is split: `Register Element Type` covers `Structural Columns` and `Structural Foundations`; `Register Detail Type` covers `Detail Items` plus the RVT-adjacent JSON type-name allowlist.
  - Extraction rules are locked: `LocationPoint.Point` for vertical instances, `LocationCurve` start point for slanted column/foundation instances, and `LocationPoint` only for registered Detail Items.
  - Project Information settings: `AT_CoordAxisMapping`, `AT_CoordUnit`, `AT_CoordTriggerFilter`; defaults are `VN-2000`, `Meters`, and `StructuralColumns`.
  - `Write Coordinates` applies axis mapping, converts all X/Y/Z values to `AT_CoordUnit`, rounds to 4 decimals, skips unchanged values, reports unsupported/failed elements, and processes every registered coordinate scope together.
  - Auto Update is document-scoped, uses `Element.GetChangeTypeGeometry()`, registers one trigger per registered coordinate category, and processes every registered scope together instead of one active UI filter.
  - `App.cs` must capture `application.ActiveAddInId` in `OnStartup()` and pass the stored field to event handlers; journal test proved casting event `sender` to `ControlledApplication` returned null and skipped registration.
  - `IUpdater.Execute()` opens no transaction and writes inside Revit's active updater transaction; `_isUpdating` is a static reentrance guard cleared in `finally`.
  - Runtime test evidence: user confirmed all registration, batch write, and auto-update flows now work as expected for Columns, Foundations, and registered Detail Items.

### Incomplete feature
- `FilterManagerCommand.cs` + `FilterWindow.xaml/.cs`: UI skeleton exists; actual `ParameterFilterElement` copy/paste logic is not implemented.
- Coordinate feature is functionally complete and closed; only future bug-fix maintenance or release-specific packaging/QA should reopen it.

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
| Coordinate scope expands only after deterministic extraction/writeback rules are explicitly locked per category | Prevent premature category drift and hidden rule mismatch | Each new category requires explicit backend work before registration is exposed |
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
| Registered coordinate categories are the runtime source of truth for batch execution and updater trigger registration | Runtime must remember all registered scopes instead of one active UI filter | Project Information trigger settings remain available, but do not narrow batch/updater execution to one scope |
| Detail Item processing requires both category registration and the RVT-adjacent JSON type-name allowlist | Category binding alone is too broad for annotation-family scope | Adds a sidecar persistence dependency that must move with the model |
| `ConvertedCoordinate` is a sealed record in the conversion layer | Immutable value semantics fit storage-ready conversion output | Axis-mapped labels reuse the same EastWest/NorthSouth property names for traceability |
| Coordinate updater registration is document-scoped | Prevent updater triggers from leaking across documents | Register on document open/create and unregister on document closing |
| Coordinate updater `UpdaterId` uses stable `AddInId` + updater GUID | Revit identifies updater persistence and trigger registration by this pair | GUID must remain stable after deployment |
| `App.cs` stores `UIControlledApplication.ActiveAddInId` during `OnStartup()` | Revit journal proved document event `sender` casting can produce null AddInId and skip registration | Keep `_addInId` as the event-handler source of truth |
| Coordinate updater trigger uses category + class `LogicalAndFilter` | Category-only triggers can include column type elements instead of only instances | Slightly more verbose trigger setup |
| `CoordinateUpdater.Execute()` never opens a transaction | Revit runs updater execution inside an active transaction | Parameter writes join the user's undo unit |
| Coordinate updater reads Project Information settings on every execution | Axis/unit settings can change during a Revit session | Small per-execution read cost |

---

## 6. Active roadmap

Excel to Revit is closed. Coordinate is also closed after end-to-end user validation. Active development priority is Filter Manager, then release-specific QA/packaging work, then optional R&D.

### A. Coordinate feature — closed implementation record

#### Closure state
The coordinate feature is operationally closed and should only be reopened for real bugs, release-specific hardening, deployment packaging, or an explicit scope expansion request.

#### Final shipped behavior
- `Register Element Type` handles supported 3D coordinate categories: `Structural Columns` and `Structural Foundations`.
- `Register Detail Type` handles Detail Item parameter binding plus RVT-adjacent JSON type-name registration in one independent operator flow.
- `Write Coordinates` processes every registered coordinate scope in the document together, not one active trigger only.
- Auto Update is document-scoped, registers one geometry trigger per registered coordinate category, and processes every registered coordinate scope together.
- `CoordSettingsDialog` remains intentionally compact: `Axis Mapping`, `Output Unit`, and `Trigger Filter`, all dropdown-only.
- `CoordinateLogService` remains the debug/support record; the ribbon UI stays compact.

#### Explicit constraints
- Preserve the locked representative-point rules per category.
- Preserve registration-state-as-scope behavior for batch and updater.
- Preserve numeric raw-parameter storage and the existing normalization/rounding path.
- Do not reintroduce a single active-trigger-only execution model.

### B. Filter Manager — active priority
- Replace skeleton behavior with real `ParameterFilterElement` copy/paste logic.
- Remove or redesign `Idling` refresh if it becomes the source of model-scale lag.
- Only keep MVVM complexity that directly supports filtering workflow.

### C. Release-specific QA / packaging
- Verify coordinate command labels, tooltips, and updater wiring in deployment builds.
- Verify RVT-adjacent JSON sidecar behavior survives real project file moves/copies.
- Verify shared-parameter registration guidance remains clear for first-run operator UX.

### D. Optional R&D
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
- Coordinate feature
  - Detailed dossier: `.Dossier/Detailed Technical Dossier - Coordinate Feature.md`
  - Components: `App.cs`, `RegisterCoordParamsCommand.cs`, `RegisterDetailItemCoordTypeCommand.cs`, `RunCoordBatchCommand.cs`, `ToggleCoordUpdaterCommand.cs`, `CoordinateBatchService.cs`, `CoordinateUpdater.cs`, `CoordinateUpdaterService.cs`, `CoordinateExtractionService.cs`, `CoordinateConversionService.cs`, `CoordinateProjectSettingsService.cs`, `CoordinateParameterBindingService.cs`, `CoordinateDetailItemRegistryService.cs`, `CoordinateLogService.cs`, `CoordSettingsDialog.xaml`
  - Keep in mind: runtime scope follows registered categories, not one active trigger only; updater registration is document-scoped and adds one geometry trigger per registered supported category; Detail Items require both category registration and the RVT-adjacent JSON type-name allowlist; `App.cs` must capture `UIControlledApplication.ActiveAddInId` during `OnStartup()`.

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
