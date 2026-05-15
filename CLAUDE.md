# ARCTOOL — TECHNICAL CONTEXT
Last updated: 2026-05-15 — roadmap refactor, Excel to Revit dossier closed, coordinate feature promoted to next priority.

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
│   │   └── ExcelToRevitCommand.cs
│   ├── Services/
│   │   ├── ExcelInteropService.cs
│   │   ├── ArcToolSettingsService.cs
│   │   └── ExcelSyncEngine.cs
│   ├── UI/
│   │   ├── FilterWindow.xaml
│   │   ├── FilterWindow.xaml.cs
│   │   ├── ExcelToRevitWindow.xaml
│   │   └── ExcelToRevitWindow.xaml.cs
│   ├── Models/
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
- `App.cs`: Ribbon bootstrapping is stable.
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

### Incomplete feature
- `FilterManagerCommand.cs` + `FilterWindow.xaml/.cs`: UI skeleton exists; actual `ParameterFilterElement` copy/paste logic is not implemented.

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

---

## 6. Active roadmap

Priority is now shifted away from Excel to Revit. That feature is closed unless a real bug appears. New development priority is the coordinate feature, then Filter Manager, then optional R&D.

### A. Coordinate feature — next active priority

#### Scope lock
V1 must target `Structural Columns` only. Do not start with “all 3D categories”. Generic expansion comes later.

#### Core principle
The real problem is not reading coordinates from Revit; it is defining which coordinate is correct for each supported element and project context. Wrong definition means the whole updater/UI/schedule chain is wrong.

#### Development phases

**Phase A — Scope & data contract**
- Session A1: lock V1 scope to `Structural Columns`.
- Session A2: define point rule for vertical vs slanted columns.
  - Vertical column: `LocationPoint.Point`.
  - Slanted column V1: use one explicit rule only, preferably base point / start point.
- Session A3: define shared parameter contract.
  - Minimum: `AT_CoordX`, `AT_CoordY`, `AT_CoordZ`.
  - Add debug/meta parameter only if it helps support or bug fixing.
- Session A4: decide storage form.
  - Prefer numeric parameters.
  - Lock unit convention once: meters or millimeters.
  - Formatting belongs to schedule/UI, not core math.

**Phase B — Coordinate engine**
- Session B1: build `CoordinateExtractionService`.
  - Support `LocationPoint`.
  - Support `LocationCurve`.
  - Return explicit unsupported state instead of guessing.
- Session B2: build `CoordinateConversionService`.
  - Base it on `ProjectLocation`, `BasePoint`, `SharedPosition`, `GetProjectPosition(...)`.
  - Do not use `GetTotalTransform().Inverse` as the primary active-document algorithm.
- Session B3: lock unit conversion and rounding policy.
- Session B4: treat VN-2000 axis swap as project mapping rule, not universal truth.
- Session B5: verify across different project location conditions before moving on.

**Phase C — Stable batch workflow**
- Session C1: create a deterministic `Run Once` command.
- Session C2: write coordinates into shared parameters.
- Session C3: validate schedule output.
- Session C4: optimize collectors and skip unchanged writes.
- Session C5: surface unsupported-element cases explicitly.

**Phase D — Dynamic update**
- Session D1: wrap the stable engine with `IUpdater`.
- Session D2: add narrow trigger scope.
- Session D3: prevent re-entry/chattering.
  - Delta check alone is not enough.
  - Canonicalize units/rounding before compare.
  - Write only when normalized value changes.
- Session D4: profile latency on real models.

**Phase E — Operator UI**
- Session E1: WPF modeless dashboard for enable/disable updater and run-once actions.
- Session E2: expose only settings that matter for support/debugging.

**Phase F — QA & deployment**
- Session F1: stress test at 100 / 1,000 / 5,000 / 10,000 elements.
- Session F2: test undo behavior.
- Session F3: test worksharing / local-central scenarios if relevant.
- Session F4: package installer.

#### Explicit constraints
- Do not build updater before batch command is correct.
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
| `BasePoint.GetSurveyPoint(Document)` | Survey coordinate context |
| `BasePoint.GetProjectBasePoint(Document)` | Project base point context |
| `ParameterFilterElement` | Filter Manager implementation target |
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
