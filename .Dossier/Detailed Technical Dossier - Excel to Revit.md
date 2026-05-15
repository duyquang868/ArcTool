# Detailed Technical Dossier — Excel to Revit

## 1. Purpose
This document is the dedicated technical dossier for the **Excel to Revit** feature. Its purpose is to preserve the architecture, data contracts, technical decisions, execution flow, useful bug history, API constraints, and the invariants that must remain intact when fixing bugs or reopening work on this feature later.

This document prioritizes **debug value** over narrative. Every section should help with at least one of these tasks: reading the code faster, rebuilding the mental model, narrowing down bugs, avoiding regression to old bugs, or identifying which parts are safe versus unsafe to change.

---

## 2. Closed scope
Excel to Revit is **closed** and is no longer part of the active roadmap. It should only be reopened for real bugs or explicit maintenance work.

The finalized feature scope is:
- Manage multiple mappings between an **Excel source** and a **Revit view target**.
- Export an Excel data region to PNG through a PDF-based pipeline.
- Import or refresh the image inside a Drafting View or Legend View.
- Detect Excel file changes via timestamps.
- Preserve image size after the user resizes it in Revit.
- Auto-update when the dialog opens if `AutoSync = true`.

Out of scope:
- No synchronization of cell data into Revit parameters.
- No reverse write-back from Revit into Excel.
- No continuous background file watcher.
- No external image links; images are **imported** into Revit.

---

## 3. Main components

### 3.1. Files
- `ArcTool.Core/Commands/ExcelToRevitCommand.cs`
- `ArcTool.Core/Services/ExcelInteropService.cs`
- `ArcTool.Core/Services/ArcToolSettingsService.cs`
- `ArcTool.Core/Services/ExcelSyncEngine.cs`
- `ArcTool.Core/UI/ExcelToRevitWindow.xaml`
- `ArcTool.Core/UI/ExcelToRevitWindow.xaml.cs`
- `ArcTool.Core/Models/ExcelMapping.cs`

### 3.2. Responsibility map
- `ExcelToRevitCommand`: entry point, guards `doc.PathName`, opens the modal WPF window, and does **not** open business transactions.
- `ExcelToRevitWindow`: UI layer, row load/save, sheet/range lookup, update triggering, and AutoSync orchestration.
- `ExcelMapping`: JSON contract plus per-row runtime helper.
- `ArcToolSettingsService`: JSON persistence next to `.rvt`, atomic write, and file status helpers.
- `ExcelInteropService`: Excel COM plus the Excel → PDF → PNG → crop export pipeline.
- `ExcelSyncEngine`: core Excel → Revit sync pipeline, transaction boundary, image create/update, and mapping state persistence.

---

## 4. Overall architecture

```text
ExcelToRevitCommand
  -> ExcelToRevitWindow
     -> ArcToolSettingsService.LoadMappings()
     -> ExcelInteropService.GetSheetNames()/GetNamedRanges()
     -> ExcelSyncEngine.CheckForChanges()
     -> ExcelSyncEngine.ExecuteUpdate()
        -> ExcelInteropService.ExportRegion()
        -> Revit Transaction 1: create image
        -> Revit Transaction 2: resize image
        -> ArcToolSettingsService.SaveMappings()
```

### 4.1. Boundary rules
- The command boundary only opens the UI and validates document state.
- The UI boundary must not contain complex Revit transaction logic.
- Core sync logic belongs in `ExcelSyncEngine`.
- Persistence logic belongs in `ArcToolSettingsService`.
- Excel COM lifetime must stay short and scoped through `using`.

### 4.2. Architecture evolution
The original pipeline used `CopyPicture`, clipboard operations, and chart workarounds. That pipeline was removed because hidden Excel mode plus the virtual device context made exports unreliable and truncated content.

The final production pipeline is:

```text
Excel sheet/range
  -> ExportAsFixedFormat(PDF)
  -> PDFtoImage (PDFium, 300 DPI)
  -> PNG
  -> SkiaSharp crop white margins
  -> ImageType.Create
  -> ImageInstance.Create
```

When reading old bugs or old notes, always prioritize the current code. Any note that still refers to `CopyPicture`, `_activeSheet` swapping inside `ExportRegion()`, or `FIXED_SCALE_FACTOR 35x` is historical data and is no longer a production invariant.

---

## 5. Data contract — ExcelMapping

`ExcelMapping` is the central contract between the UI, JSON persistence, and the sync engine.

### 5.1. Field inventory
- `Id`: stable row identifier generated from `Guid.NewGuid().ToString()`.
- `ViewName`: target Revit view name.
- `ViewType`: `DraftingView` or `LegendView`.
- `FilePath`: absolute Excel file path.
- `WorkSheet`: source sheet name.
- `Region`: Named Range name, or `null`.
- `RegionType`: `NamedRange`, `PrintArea`, `UsedRange`.
- `AutoSync`: auto-update when the dialog opens.
- `LastModified`: last successful sync time, using local time.
- `ImageInstanceId`: `ElementId.Value` of the current Revit image.
- `StoredWidth`, `StoredHeight`: image size stored in JSON.

### 5.2. Sentinel values
- `ImageInstanceId = 0`: never imported.
- `StoredWidth = 0.0`, `StoredHeight = 0.0`: no persisted size yet.
- `LastModified = DateTime.MinValue`: on first dialog open, the file is always treated as changed.
- `Region = null`: no Named Range selected; the engine falls back to PrintArea or UsedRange.

### 5.3. Final unit invariant
**Current production invariant:** `StoredWidth` and `StoredHeight` are stored in **millimeters**, not internal feet.

This matters because some older notes described feet in a few places. When fixing Smart Scale bugs, trust the current code:
- read from `ImageInstance.Width/Height` -> convert **from internal units to millimeters**;
- write back to the instance -> convert **to internal units from millimeters**.

### 5.4. View name convention
- `PrintArea` / `UsedRange` -> `ViewName = WorkSheet`
- `NamedRange` -> `ViewName = WorkSheet_Region`

### 5.5. Enum naming invariant
`ExcelViewType` and `ExcelRegionType` use the `Excel` prefix to avoid collisions with Revit namespaces, especially `Autodesk.Revit.DB.ViewType`.

---

## 6. Persistence contract — ArcToolSettingsService

### 6.1. File location
JSON settings always live next to the currently open `.rvt` file:
- filename: `ArcTool_ExcelSync.json`
- root rule: `doc.PathName` must be valid before load/save

### 6.2. Atomic write invariant
JSON must never be written directly with `File.WriteAllText(finalPath, ...)`.

Correct pattern:
```text
serialize -> write temp .tmp -> File.Replace() or File.Move()
```

Reason:
- a crash in the middle must not corrupt the original JSON;
- a leftover temp file is acceptable, but the primary state must remain safe.

### 6.3. Deserialize behavior
- File missing -> return an empty `List<ExcelMapping>`.
- JSON corrupt -> debug log + `.corrupt_[timestamp]` backup + return empty list.
- Ordinary IO error -> debug log + return empty list.

### 6.4. Timestamp invariant
`HasFileChanged(mapping)` uses:
```text
File.GetLastWriteTime(mapping.FilePath) > mapping.LastModified
```

Both sides must share the same **local time basis**.

Invariant rules:
- set `LastModified` with `DateTime.Now`;
- do not use `DateTime.UtcNow` anywhere in this flow.

### 6.5. Known limitations
- If `SaveMappings()` fails after the Revit transaction has already committed, the model state has changed but the JSON state has not.
- `JsonStringEnumConverter` is not friendly to very old JSON that stored enums as numeric values.
- Atomic write assumes the temp file and target file stay on the same volume.

---

## 7. Command boundary — ExcelToRevitCommand

### 7.1. Current behavior
The command only does four things:
- acquire `Document` from `ExternalCommandData`;
- guard against a missing document;
- guard against an unsaved `.rvt` file;
- open `ExcelToRevitWindow` with `ShowDialog()`.

### 7.2. Important invariant
The command must **not** open transactions for import/update work. All business transactions belong inside `ExcelSyncEngine.ExecuteUpdate()`.

### 7.3. Why the modal window was kept
The window uses `ShowDialog()` instead of modeless UI plus external events because:
- it is simpler for this flow;
- it preserves API context while the command is still active;
- it avoids overcomplicating a feature that is already stable enough.

If this feature is ever refactored to modeless UI plus external events, that is a major architectural change, not a small bugfix.

---

## 8. UI contract — ExcelToRevitWindow

### 8.1. DataGrid layout
The column order is locked as:
1. Select
2. Status Dot
3. View Name
4. Auto Sync
5. Last Modified
6. WorkSheet
7. Region
8. View Type
9. File Path
10. Update

Toolbar:
- `+`
- `−`
- `Update All`

### 8.2. Window load flow
```text
Window_Loaded
  -> LoadMappingsIntoRows()
  -> RefreshAllStatuses()
  -> RunAutoSyncRows()
  -> RefreshAllStatuses()
```

### 8.3. Status semantics
- Green: file exists, no changes.
- Red: file exists, newer than `LastModified`.
- Yellow: file path missing, file moved, or file deleted.

### 8.4. Lookup flow
- Browse file -> `GetSheetNames()` -> populate the WorkSheet dropdown.
- WorkSheet selected -> `GetNamedRanges(sheet)` -> build Region options.
- Region options always begin with `Print Area`.
- `Used Range` is not shown by default; it is only added when the current row already uses `UsedRange`, so legacy state is preserved.

### 8.5. Event suppression invariant
`_suppressRowEvents` is a critical invariant for the entire UI.

Reason:
- the code-behind performs programmatic property mutations;
- those properties trigger `Row_PropertyChanged`;
- without correct suppression, the code can produce event cascades, double-loads, or logic loops.

### 8.6. BUG-P3-01 lesson learned
When browsing for a file, the correct order is:
```text
set _suppressRowEvents = true
-> row.FilePath = ...
-> LoadLookupData(...)
-> restore the flag in finally
```

If `FilePath` is assigned first and suppression happens later, the event already fired and `LoadLookupData()` will run twice.

### 8.7. AutoSync contract
AutoSync runs **once when the dialog opens** for rows that satisfy all conditions:
- `AutoSync == true`
- `FileExists == true`
- `HasChanges == true`

This feature is not a background watcher.

---

## 9. Excel COM contract — ExcelInteropService

### 9.1. Lifetime model
`ExcelInteropService` is disposable and must be used in short scope:
```csharp
using (var svc = new ExcelInteropService())
{
    svc.OpenFile(path);
    ...
}
```

Do not keep the service alive across multiple user actions unless that is truly unavoidable.

### 9.2. COM release invariants
- Release in child -> parent order.
- Release wrapper COM objects like `Sheets` and `Names` explicitly.
- Do not call `ReleaseComObject()` after a COM object has already been `Delete()`-ed.
- Do not keep long-lived `ActiveSheet` references unless absolutely necessary.

### 9.3. Current public API
- `OpenFile(filePath)`
- `GetActiveSheetName()`
- `ExportPrintAreaAsHighResImage(outputPath)`
- `GetSheetNames()`
- `GetNamedRanges(sheetName)`
- `ExportRegion(sheetName, regionName, outputPath)`
- `Dispose()`

### 9.4. Current ExportRegion behavior
Resolve order:
```text
NamedRange -> PrintArea -> UsedRange
```

Current production note:
- `ExportRegion()` **no longer** uses the `_activeSheet` swap pattern.
- `ExportRangeInternal()` now receives `Worksheet ws` and `Range range` directly.
- Any older note that says `_activeSheet` must be restored before releasing `ws` is now historical only.

### 9.5. Why old docs can mislead bugfix work
This feature evolved significantly across multiple sessions. Three categories of historical notes are especially dangerous:
- notes about the `CopyPicture` pipeline;
- notes about `_activeSheet` swapping;
- notes that describe Smart Scale storage in feet.

When old notes conflict with current code, treat **the current code plus this dossier** as the source of truth.

---

## 10. Export pipeline in detail

### 10.1. Final pipeline
```text
Excel Worksheet/Range
  -> PageSetup normalize
  -> ExportAsFixedFormat(PDF)
  -> EnsurePdfiumLoaded()
  -> PDFtoImage render page 0 at 300 DPI
  -> EnsureSkiaSharpLoaded()
  -> crop white margins if Skia runtime is available
  -> PNG output
```

### 10.2. PageSetup normalization
Before exporting PDF, the code forces:
- `PrintArea = range.Address[false, false]`
- `Zoom = false`
- `FitToPagesWide = 1`
- `FitToPagesTall = 1`
- margins = 0
- prefer `xlPaperEsheet`, fallback to `A3`, fallback again to the current paper size

The goal is to force the entire export region onto a single PDF page so the PDF render stays predictable.

### 10.3. PDF render invariant
- engine: `PDFtoImage`
- native backend: PDFium
- current DPI: `300`
- current page index: `0`
- annotations: `false`

### 10.4. Crop behavior
If `libSkiaSharp.dll` loads successfully:
- scan white margins on top, bottom, left, and right;
- crop using `ExtractSubset`;
- overwrite the PNG.

If it does not load:
- **do not fail the whole command**;
- keep the raw PNG, even if white margins remain.

### 10.5. Native dependency invariant
Deployment must include:
- `pdfium.dll`
- `libSkiaSharp.dll`

Search order:
1. next to `ArcTool.Core.dll`
2. `native\...`
3. `runtimes\win-x64|x86|arm64\native\...`

### 10.6. Failure policy
- Missing `pdfium.dll` -> export fails softly at the service level, and the sync engine converts it into a user-facing error.
- Missing `libSkiaSharp.dll` -> skip crop and continue.

### 10.7. Closed root cause
The most important historical export root cause was hidden Excel plus the virtual device context making `CopyPicture` unreliable. That is why the production pipeline must not go back to clipboard/chart hacks unless there is very strong technical evidence.

---

## 11. Core sync contract — ExcelSyncEngine

### 11.1. Role
`ExcelSyncEngine` is the orchestration core between Excel export, Revit image updates, and JSON persistence.

### 11.2. Public API
- `CheckForChanges(IEnumerable<ExcelMapping>)`
- `ExecuteUpdate(ExcelMapping, Document, List<ExcelMapping>)`
- `GetOrCreateView(string, ExcelViewType, Document)`

### 11.3. CheckForChanges contract
- read filesystem timestamps only;
- do not open Excel;
- do not create Revit transactions;
- return a dictionary keyed by `mapping.Id`.

If a Status Dot bug appears, this is the first place to inspect.

### 11.4. ExecuteUpdate full flow
```text
1. Validate mapping.ViewName
2. Export Excel -> temp PNG
3. Read old image width/height before delete
4. Transaction 1: delete old image, create/get view, create ImageType, create ImageInstance
5. Read natural size after Tx1
6. Transaction 2: resize image
7. Mutate mapping after commit
8. Save JSON
9. Delete temp PNG in finally
```

### 11.5. Why two transactions exist
`ImageInstance` resize behavior is not stable enough if Width/Height is assigned in the same transaction that creates the instance. The production flow therefore uses two transactions:
- Tx1 to create the image;
- Tx2 to resize the image.

If Tx2 fails, the image still exists at natural size. That is an acceptable soft degradation.

### 11.6. Smart Scale invariant
Before deleting the old instance, the engine must read its real width and height if the instance is still valid. It must not trust JSON blindly, because the user may have resized the image directly in Revit after the previous sync.

### 11.7. Mapping mutation invariant
Only mutate these fields **after** the create/resize transaction flow has completed successfully at an acceptable level:
- `ImageInstanceId`
- `StoredWidth`
- `StoredHeight`
- `LastModified`

Do not mutate them early inside a transaction and then try to commit later, because a commit failure would create state drift between the Revit model and JSON.

### 11.8. Error policy
Current behavior:
- configuration errors or unusable export results -> `InvalidOperationException` with a clear user-facing message;
- IO errors during JSON save -> propagate;
- unknown errors -> debug log + propagate.

Reason: an older silent `return false` policy at the UI boundary made users think the command simply did nothing.

---

## 12. Revit view creation contract

### 12.1. Drafting View
- Uses `ViewDrafting.Create(...)`.
- Can create a new view if no matching one exists.
- If the name conflicts in a race or undo edge case, the code falls back to a timestamp suffix.

### 12.2. Legend View
Production invariant:
- Revit API **does not provide** a usable public create path for legend views from scratch.
- The only production approach is to duplicate an existing legend.

Logic:
- if the target view name already exists -> reuse it;
- otherwise -> try `ArcTool_LegendTemplate` first;
- if missing -> fall back to any non-template Legend View;
- if no legend exists at all -> throw an error instructing the user to create a blank legend manually.

### 12.3. Bugfix implication
When a Legend View bug appears, first determine which layer is failing:
- the project does not contain a valid legend source;
- the duplicate/create path is broken;
- the name conflicts after duplication.

Do not waste time searching for a production-grade `Legend.Create()` API. The absence of a suitable public create path has already been verified multiple times.

---

## 13. Remaining known limitations
- The feature still depends on Excel COM; it is not a pure OpenXML pipeline.
- There is no background file watcher; changes are only checked when the dialog opens.
- If JSON save fails after the Revit model has committed, sync state can drift until the next successful update.
- Cropping only trims simple white margins; it does not understand semantic content.
- `GetNamedRanges()` skips workbook-level or cross-sheet ranges that cannot be resolved as `RefersToRange.Worksheet.Name == sheetName`.
- `UsedRange` is the final fallback and may be broader than intended.

---

## 14. Bug history worth remembering
Only keep bugs that provide real regression-prevention value.

### BUG-E6 — `View` ambiguous
Root cause: `UseWindowsForms=true` introduced `System.Windows.Forms.View`, which conflicted with `Autodesk.Revit.DB.View`.

Invariant learned:
```csharp
using RevitView = Autodesk.Revit.DB.View;
```

### BUG-P3-01 — double-call `LoadLookupData()`
Root cause: event cascade when `FilePath` was set before suppression.

Invariant learned:
- set `_suppressRowEvents = true` before any programmatic mutation that may trigger a `PropertyChanged` loop.

### BUG-E9 — long-lived `_activeSheet` COM state
Root cause: keeping COM references alive too long caused invalid RCW state or made cleanup unreliable.

Invariant learned:
- prefer local scope plus immediate release.

### BUG-E10 — missing `pdfium.dll`
Root cause: a Revit add-in does not resolve native libraries the same way a normal .NET application does.

Invariant learned:
- native runtime probing must be explicit.
- missing PDFium must never be allowed to crash the AppDomain.

### BUG-E11 — missing `libSkiaSharp.dll`
Root cause: the crop step depends on a separate native runtime.

Invariant learned:
- crop is an optional enhancement; the main export pipeline must not die because crop failed.

### BUG-E12 — export failed but the UI stayed silent
Root cause: the old boundary swallowed errors with `return false`.

Invariant learned:
- every sufficiently serious user-facing error must be surfaced clearly through an exception or explicit message.

---

## 15. Troubleshooting map

### 15.1. Wrong Status Dot state
Check in this order:
1. `ArcToolSettingsService.FileExists()`
2. `ArcToolSettingsService.HasFileChanged()`
3. whether `LastModified` still uses local time
4. whether `mapping.Id` remains stable
5. whether `RefreshAllStatuses()` runs after update

### 15.2. Excel process does not exit
Check:
- whether `Sheets` / `Names` wrappers were released
- whether local `Worksheet` / `Range` COM objects were released
- whether `Workbook.Close(false)` and `Application.Quit()` were called
- whether newer code accidentally keeps COM references alive beyond short scopes

### 15.3. Exported PNG is blank or the wrong region
Check:
- whether `sheetName` is correct
- whether `regionName` really belongs to that sheet
- whether `PrintArea` is valid
- whether fallback dropped to `UsedRange`
- whether `PageSetup` was blocked by a protected sheet
- whether the PDFium runtime actually loaded

### 15.4. Legend View cannot be created
Check:
- whether the project contains any legend view at all
- whether `ArcTool_LegendTemplate` exists
- whether duplication returned a valid `ElementId`
- whether the new view name conflicts

### 15.5. Smart Scale resets
Check:
- whether `existingInst.Width/Height` was read before `doc.Delete(...)`
- whether the mm <-> internal units conversion is still correct
- whether Tx2 rolled back
- whether `StoredWidth/StoredHeight` was saved back into JSON after update

### 15.6. JSON state goes missing
Check:
- whether `SaveMappings()` threw after commit
- whether `.tmp` or `.corrupt_*` files appeared
- whether the `.rvt` file was moved or copied without the JSON file

---

## 16. Minimum regression test matrix for bugfixes
When fixing a bug in this feature, rerun at least these test groups.

### A. Persistence
- `.rvt` saved / not saved
- normal JSON / corrupt JSON
- Excel file exists / missing

### B. Lookup UI
- Excel file with multiple sheets
- sheet with Named Range
- sheet without Named Range
- Print Area fallback
- UsedRange fallback

### C. Revit view target
- Drafting View first import
- Drafting View second update
- Legend View with existing template
- Legend View with no template available

### D. Smart Scale
- first import
- manual resize inside Revit
- update after Excel changes
- verify that width/height do not reset

### E. Runtime dependency
- deployment with `pdfium.dll`
- deployment without `pdfium.dll`
- deployment without `libSkiaSharp.dll`

### F. AutoSync
- `AutoSync = false`
- `AutoSync = true` + changed file
- `AutoSync = true` + missing file

---

## 17. Future bugfix rules
- Do not refactor this feature broadly just to make the code look nicer.
- Every change must preserve the invariants locked in this dossier.
- If you touch `ExcelInteropService`, think first about COM lifetime and native runtime behavior.
- If you touch `ExcelSyncEngine`, think first about transaction boundaries and state drift between the Revit model and JSON.
- If you touch `ExcelToRevitWindow`, think first about event cascades.
- If expected behavior conflicts with old notes, inspect the current production code first and only then compare it with this dossier.

---

## 18. Usage rule for this dossier
- `CLAUDE.md` should keep only short summaries and high-level invariants.
- All deep technical details for Excel to Revit belong in this file.
- When another feature is closed, create a similar dedicated dossier instead of expanding `CLAUDE.md` again.
