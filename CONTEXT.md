# ARCTOOL — AI SESSION CONTEXT
<<<<<<< HEAD
> Paste file này vào ĐẦU mỗi session chat mới với AI.
> Cập nhật sau mỗi session làm việc.
> Last updated: 2026-04-16 — Session 5.7.1: MERGE ArcTool_Instructions.md - Inherit old content + add SESSION 5 updates
=======
> Last updated: 2026-04-19 — Session 5.12: ExcelToRevit V3.0 — Full Rebuild
>>>>>>> 175ed814420ed8c34d98e53dd28aad238ef98aa5

---

## 1. THÔNG TIN DỰ ÁN

| Mục | Chi tiết |
|---|---|
| Tên dự án | ArcTool |
| Namespace chính | `ArcTool.Core` |
| Nền tảng | Autodesk Revit 2026 (API 2026) |
| Ngôn ngữ | C# / .NET 8.0 |
| IDE | Visual Studio Enterprise 2026 |
| UI Framework | WPF (modeless window) + WinForms (dialog chọn Family) |
| Icon/Resource | `Properties/Resources.resx` (Access Modifier: Public) |
| Revit Unit System | `UnitTypeId` (ForgeTypeId) — KHÔNG dùng `DisplayUnitType` cũ |

---

## 2. CẤU TRÚC THƯ MỤC

```
<<<<<<< HEAD
ArcTool/
├── ArcTool.slnx
├── CONTEXT.md                          ← Project status & roadmap
├── ArcTool_Instructions.md             ← Development guide (NEW - SESSION 5.6) ✨
├── SKILL.md                            ← Coding patterns & techniques (NEW - SESSION 5.6) ✨
├── ArcTool.Core/
│   ├── App.cs                          ← Ribbon UI, entry point
│   ├── App.config
│   ├── ArcTool.Core.csproj
│   ├── Commands/
│   │   ├── CreateVoidFromLinkCommand.cs
│   │   ├── MultiCutCommand.cs
│   │   ├── ArrangeDimensionCommand.cs
│   │   ├── FilterManagerCommand.cs
│   │   └── ExcelToRevitCommand.cs      ← Phase 3: Unified Excel → Revit Image Pipeline
│   ├── Services/
│   │   └── ExcelInteropService.cs
│   ├── UI/
│   │   ├── FilterWindow.xaml
│   │   └── FilterWindow.xaml.cs
│   ├── Utilities/
│   │   └── SelectionFilters.cs
│   ├── Models/
│   │   └── (placeholder for future data models)
│   ├── Resources/
│   │   ├── icon_create_16.jpg
│   │   ├── icon_create_32.jpg
│   │   ├── icon_cut_16.png
│   │   └── icon_cut_32.png
│   └── Properties/
│       ├── Resources.resx
│       └── Resources.Designer.cs
└── ArcTool.TestConsole/
    └── TestConsole/
        └── TestConsole/
            ├── Program.cs              ← Test Excel export độc lập
            └── ArcTool.TestConsole.csproj
=======
ArcTool.Core/
├── App.cs
├── Commands/
│   ├── CreateVoidFromLinkCommand.cs
│   ├── MultiCutCommand.cs
│   ├── ArrangeDimensionCommand.cs
│   ├── FilterManagerCommand.cs
│   └── ExcelToRevitCommand.cs          ← V3.0: REBUILD HOÀN TOÀN
├── Services/
│   ├── ExcelInteropService.cs          ← V5.2: ổn định, không sửa
│   └── ArcToolSettings.cs             ← V1.0: ổn định, không sửa
├── UI/
│   ├── FilterWindow.xaml + .cs
│   ├── SyncStatusWindow.xaml           ← V1.0 XAML ổn định, không sửa
│   └── SyncStatusWindow.xaml.cs       ← V2.0: Thêm SetStatus() + constructor mới
└── Utilities/
    └── SelectionFilters.cs
>>>>>>> 175ed814420ed8c34d98e53dd28aad238ef98aa5
```

---

## 3. TRẠNG THÁI HIỆN TẠI — BUG REGISTRY (ExcelToRevit ecosystem)

<<<<<<< HEAD
### A. Ribbon UI — `App.cs` (V5.1)
- Tab: `ArcTool`
- Panel 1: `Void Tools` → SplitButton "Void Manager"
  - Nút chính: **Create Void** → `CreateVoidFromLinkCommand`
  - Nút phụ (Separator): **Multi-Cut** → `MultiCutCommand`
- Panel 2: `Annotation Tools`
  - Nút: **Arrange Dimensions** → `ArrangeDimensionCommand`
- Helper: `ConvertToImageSource(Bitmap)` chuyển Resource sang WPF ImageSource

### B. Create Void — `CreateVoidFromLinkCommand.cs` (V4.0)
- Tự động tạo Generic Model Void tại **TẤT CẢ** dầm (Structural Framing) trong file Link
- User chọn: (1) Family Void qua WinForms dialog, (2) File Link qua PickObject
- Logic vị trí: midpoint của `LocationCurve`, áp dụng `linkTransform`
- Logic kích thước: Width/Height từ tham số dầm, Length = chiều dài dầm
- Tạo instance kiểu **Face-Based** (`NewFamilyInstance` với `linkedFaceRef`)
- **BIẾT RỦI RO:** Face-Based sẽ bị "unhosted" nếu Link reload với geometry thay đổi

### C. Multi-Cut — `MultiCutCommand.cs` (V2.0)
- Cắt Tường (Walls) + Cột Kết cấu + Cột Kiến trúc bằng các Void đã tạo
- Broad Phase: `BoundingBoxIntersectsFilter` để lọc sơ bộ (tránh O(n²))
- Dùng `InstanceVoidCutUtils.AddInstanceVoidCut()`
- **TODO:** Thêm Narrow Phase bằng Solid Intersection

### D. Arrange Dimensions — `ArrangeDimensionCommand.cs` (V1.0)
- Pick đường Dim gốc (baseline), sau đó pick liên tục các Dim tiếp theo
- Tự động tịnh tiến Dim cách đều nhau theo `Snap Distance × View Scale`
- Dùng `TransactionGroup` để gộp toàn bộ thao tác vào 1 lần Undo
- Filter: `LinearDimensionSelectionFilter` (chỉ cho phép chọn Linear Dim)

### E. Excel Export Engine — `ExcelInteropService.cs` (V5.1)
- Đọc file Excel (hidden mode), export Print Area hoặc UsedRange thành PNG
- Scale factor: 35x
- Dùng COM Interop, có `IDisposable`
- ✅ **BUG FIXED SESSION 4:** Tất cả 4 bugs COM management đã fix (BUG-E1, E2, E3, E4)

### F. Filter Manager — `FilterManagerCommand.cs` + `FilterWindow.xaml`
- UI WPF modeless đã xong (FilterWindow với 2 DataGrid: Filters + Views)
- Command skeleton đã xong, dùng `Idling` event để real-time update
- **TODO:** Implement logic Copy/Paste Filter thực sự bằng `ParameterFilterElement` API

### G. Excel to Revit Command — `ExcelToRevitCommand.cs` (V1.0 - Final) — **PHASE 3 REFACTOR**
- **🎯 Unified Pipeline:** Gộp ExcelInteropService + ImageType.Create API thành 1 lệnh duy nhất
- **✨ Workflow:**
  1. User chọn file Excel
  2. ExcelInteropService xuất Print Area → Temp PNG (hidden Excel)
  3. ImageType.Create(doc, ImageTypeOptions) — tạo ImageType từ file PNG
  4. Dialog chỉnh Scale (%) — modern TableLayoutPanel design
  5. ImageInstance.Create() đặt ảnh tại **TÂM VIEW** tự động
  6. Xoá Temp PNG sau commit
- **🔧 Technical Details:**
  - ✅ **API Confirmation:** Revit 2026 API **CÓ HỖ TRỢ** `ImageType.Create()` công khai
  - ✅ **SESSION 5 FIX:** Giải quyết ambiguous reference `TaskDialog` và `TextBox`
    - Added alias: `using RevitTaskDialog = Autodesk.Revit.UI.TaskDialog;`
    - Changed all `TaskDialog.Show()` → `RevitTaskDialog.Show()` (9 occurrences)
    - Changed `new TextBox` → `new System.Windows.Forms.TextBox` (explicit qualification)
  - Temp file path: unique GUID (`ArcTool_Excel_{GUID}.png`)
  - Resolution: 300 DPI (ImageTypeOptions)
  - ImagePlacementOptions: `BoxPlacement.Center` → center-aligned mặc định
  - Transaction: Manual mode, RollBack nếu lỗi, Commit nếu thành công
  - Error handling: Chi tiết từng bước (Excel open, export, ImageType create, ImageInstance create, scale apply)
- **🏆 Build Status:** ✅ **BUILD SUCCESSFUL** — 0 errors, 0 warnings
- **Ribbon:** Tab "ArcTool" → Panel "Excel Tools" → Button "Excel to Revit"

### Former Features (Deprecated/Merged):
- ❌ **ImageImportCommand.cs** — **REMOVED SESSION 5** (chức năng merged vào ExcelToRevitCommand.cs)
  - Lý do: ImageType.Create() được Revit 2026 hỗ trợ công khai, không cần workaround
  - ExcelToRevitCommand cung cấp complete pipeline: Excel → PNG → ImageType → ImageInstance

---

## 4. BUG ĐÃ PHÁT HIỆN — TRẠNG THÁI ĐẦY ĐỦ

> Cập nhật trạng thái: [ ] Chưa fix / [x] Đã fix

### 🔴 BUG NGHIÊM TRỌNG

- [x] **MultiCutCommand** — `(int)elem.Category.Id.Value` gây Integer Overflow
  - ✅ Fix: Cast về `(long)BuiltInCategory.OST_Walls` thay vì `(int)`
  - File: `VoidSelectionFilter.AllowElement()` + `CutTargetSelectionFilter.AllowElement()`

- [x] **CreateVoidFromLinkCommand** — `GetParamValue` chỉ tìm trên `Symbol`, bỏ sót `Instance`
  - ✅ Fix: Tìm Instance parameters trước, fallback về Symbol nếu không tìm thấy
  - File: dòng 106-113

- [ ] **CreateVoidFromLinkCommand** — `SetParam(voidInst, "Height", -beamHeight)` gán giá trị âm
  - ⏳ TODO: Cần thiết kế lại: dùng Mirror hoặc đổi hướng Family

- [x] **FilterManagerCommand** — `_lastUpdate` là instance field, reset mỗi lần chạy lệnh
  - ✅ Fix: Đổi thành `private static DateTime _lastUpdate`

- [x] **ExcelToRevitCommand (SESSION 5)** — Ambiguous reference `TaskDialog` (System.Windows.Forms vs Autodesk.Revit.UI)
  - ✅ Fix: Added alias `using RevitTaskDialog = Autodesk.Revit.UI.TaskDialog;`
  - ✅ Fix: Changed all 9 `TaskDialog.Show()` calls to `RevitTaskDialog.Show()`
  - ✅ Fix: Explicit qualification `new System.Windows.Forms.TextBox` for TextBox control
  - File: `ExcelToRevitCommand.cs` lines 8, 49, 56, 76, 88, 126, 138, 161, 171, 190, 252

### 🟠 RỦI RO / CẦN CẢI THIỆN

- [x] **ArrangeDimensionCommand** — baseline cập nhật dù Dim lỗi (Line == null)
  - ✅ Fix: `MoveDimensionToMatchSnap()` trả về `bool`, chỉ update baseline nếu `moved == true`

- [ ] **ArrangeDimensionCommand** — không kiểm tra `activeView.Scale == 0` (3D view)
  - ⏳ TODO: Thêm guard clause `if (activeView.Scale == 0) return Result.Failed;` trước khi pick

- [x] **CreateVoidFromLinkCommand** — `doc.Regenerate()` thừa trước `t.Commit()`
  - ✅ Fix: Đã xóa

- [ ] **FilterManagerCommand** — `Idling` event là anti-pattern cho model lớn
  - ⏳ TODO: Nâng cấp lên `IExternalEventHandler` + `ExternalEvent.Raise()`

### 🔴 BUG MỚI PHÁT HIỆN SESSION 3 — ExcelInteropService.cs

- [ ] **[BUG-E1] ExcelInteropService** — `ReleaseObject` có `finally { obj = null; }` VÔ NGHĨA
  - **Mức độ:** 🔴 Nghiêm trọng — COM leak tiềm ẩn
  - **Nguyên nhân:** `obj` là local parameter (pass-by-value). Gán `obj = null` chỉ null biến local trong stack frame, KHÔNG null field gốc `_excelApp`, `_workbook` ở caller. COM object không bao giờ được set null đúng cách.
  - **Fix:**
    ```csharp
    // Xóa finally { obj = null; } trong ReleaseObject()
    private void ReleaseObject(object obj)
    {
        if (obj == null) return;
        try { Marshal.ReleaseComObject(obj); }
        catch { }
    }

    // Null field GỐC trong Dispose() mới có tác dụng
    public void Dispose()
    {
        if (_activeSheet != null) { ReleaseObject(_activeSheet); _activeSheet = null; }
        if (_workbook != null)
        {
            try { _workbook.Close(false); } catch { }
            ReleaseObject(_workbook);
            _workbook = null;
        }
        if (_excelApp != null)
        {
            try { _excelApp.Quit(); } catch { }
            ReleaseObject(_excelApp);
            _excelApp = null;
        }
        GC.Collect();
        GC.WaitForPendingFinalizers();
    }
    ```

- [ ] **[BUG-E2] ExcelInteropService** — Comment sai: class summary ghi "50x Scale" nhưng constant là 35.0
  - **Mức độ:** 🟡 Technical Debt — gây nhầm lẫn khi maintain
  - **Fix:** Sửa dòng summary thành:
    ```csharp
    /// Updated V5: Synchronous Execution - No Sleep, No Retry, Hardcode 35x Scale.
    private const double FIXED_SCALE_FACTOR = 35.0;
    ```

- [ ] **[BUG-E3] ExcelInteropService** — `chartObj` bị `ReleaseObject()` SAU KHI đã `Delete()` — thứ tự sai
  - **Mức độ:** 🟠 Rủi ro — có thể gây `InvalidComObjectException` bị swallow âm thầm
  - **Nguyên nhân:** Sau `chartObj.Delete()`, COM object đã bị revoke. Gọi `Marshal.ReleaseComObject(chartObj)` tiếp tục là undefined behavior với COM. Thêm nữa, `chart` (child) phải được release trước `chartObj` (parent).
  - **Fix — thứ tự release đúng (child → parent):**
    ```csharp
    finally
    {
        // 1. Release child trước
        ReleaseObject(chart);
        // 2. Delete và KHÔNG release thêm (Delete đã dọn COM)
        if (chartObj != null)
        {
            try { chartObj.Delete(); } catch { }
            // KHÔNG gọi ReleaseObject(chartObj) ở đây
        }
        // 3. Release container cuối
        ReleaseObject(chartObjects);
    }
    ```

- [ ] **[BUG-E4] ExcelInteropService** — `_activeSheet` KHÔNG được `ReleaseComObject` trong `Dispose()`
  - **Mức độ:** 🔴 COM memory leak — Excel process sẽ không thoát sạch
  - **Nguyên nhân:** `_activeSheet` là COM Worksheet object, nhưng `Dispose()` chỉ release `_workbook` và `_excelApp`. `_activeSheet` bị bỏ quên hoàn toàn.
  - **Fix:** Thêm vào đầu `Dispose()` (trước khi close workbook):
    ```csharp
    if (_activeSheet != null) { ReleaseObject(_activeSheet); _activeSheet = null; }
    ```

---

## 5. QUYẾT ĐỊNH KỸ THUẬT ĐÃ CHỐT

| Quyết định | Lý do | Trade-off đã chấp nhận |
|---|---|---|
| Face-Based Void approach | Đảm bảo Void bám theo mặt phẳng dầm | Unhosted nếu Link reload |
| BoundingBox broad-phase trong MultiCut | Tránh O(n²) khi dự án lớn | Chưa có narrow phase |
| Height âm để đảo chiều Void | Fix lỗi vị trí Bottom tạm thời | Cần refactor về sau |
| `TransactionGroup` trong ArrangeDim | Gộp thành 1 lần Undo | — |
| WinForms cho Family selection dialog | Đơn giản, không cần MVVM | UI không đồng nhất với WPF |
| `Idling` event cho FilterManager | Prototype nhanh | Cần đổi sang ExternalEvent |
| COM release: child trước parent | Tránh InvalidComObjectException | — |

---

## 6. QUY TẮC LẬP TRÌNH (BẮT BUỘC TUÂN THEO)
=======
> ⚠️ **CẢNH BÁO:** `ExcelToRevitCommand.cs` trên repo hiện tại là **STUB chưa hoàn thiện**.
> Toàn bộ pipeline import Excel → Revit KHÔNG hoạt động.
> Cần rebuild hoàn toàn theo spec dưới đây.

### 🔴 CRITICAL — Lệnh không chạy được

| ID | File | Mô tả bug | Root cause |
|---|---|---|---|
| ETR-01 | `ExcelToRevitCommand.cs` | File picker dùng `RevitTaskDialog` — không mở được file browser | `TaskDialog` là dialog thông báo, không phải file picker. Phải dùng `OpenFileDialog` từ WinForms |
| ETR-02 | `ExcelToRevitCommand.cs` | Toàn bộ import pipeline là stub — hardcode `C:\Temp\sample.xlsx` | `ExcelInteropService`, `ImageType.Create()`, `ImageInstance.Create()` chưa được gọi |
| ETR-03 | `ExcelToRevitCommand.cs` | `ExcelRefreshHandler` class không tồn tại | CONTEXT.md mô tả handler có `StoredWidth/Height/TargetViewId/ImageInstanceId` nhưng chỉ có `ReopenHandler` stub |
| ETR-04 | `SyncStatusWindow.xaml.cs` | `SetStatus(bool)` method không tồn tại | ExcelToRevitCommand gọi `_statusWindow?.SetStatus(false)` → compile error |

### 🟠 HIGH — Crash / Race condition

| ID | File | Mô tả bug | Root cause |
|---|---|---|---|
| ETR-05 | `ExcelToRevitCommand.cs` | Race condition trong `ScheduleToast()` | Thiếu `lock(_debounceLock)` và thiếu field `private static readonly object _debounceLock` |
| ETR-06 | `ExcelToRevitCommand.cs` | `ReopenHandler` giữ stale `UIDocument` | Constructor capture `uidoc` từ `Execute()` lần đầu → crash nếu document thay đổi. Phải dùng `app.ActiveUIDocument` trong `Execute(UIApplication)` |
| ETR-07 | `ExcelToRevitCommand.cs` | `_reopenEvent` không được reset khi chạy lại lệnh | `StopWatcher()` không null `_reopenEvent` → lần chạy thứ 2 vẫn dùng handler cũ với document cũ |

### 🟡 MEDIUM — Feature missing

| ID | File | Mô tả bug | Root cause |
|---|---|---|---|
| ETR-08 | `ExcelToRevitCommand.cs` | `SyncStatusWindow` không được hiện sau import | `ShowStatusWindow()` method không tồn tại |
| ETR-09 | `ExcelToRevitCommand.cs` | Temp PNG không được dọn dẹp | `CleanupTempFile()` không tồn tại |
| ETR-10 | `ExcelToRevitCommand.cs` | Settings không được load/save | `ArcToolSettings.Load()`/`.Save()` không được gọi → `LastExcelFile`, `LastScale` không persist |
| ETR-11 | `ExcelToRevitCommand.cs` | Import options dialog không tồn tại | Không có UI để user nhập Scale% và chọn CreateNewView |
| ETR-12 | `SyncStatusWindow.xaml.cs` | Constructor cũ nhận `changedFilePath` — không khớp với design persistent window | Window nên được tạo 1 lần sau import, nhận `watchedFilePath`. `SetStatus()` cập nhật state sau đó |

---

## 4. KIẾN TRÚC ĐÚNG — ExcelToRevit V3.0

### 4.1 Luồng thực thi chính (Execute)

```
[ExcelToRevitCommand.Execute()]
  │
  ├─ 1. StopWatcher()               — dừng watcher + event cũ hoàn toàn
  │
  ├─ 2. OpenFileDialog              — WinForms dialog, filter "*.xlsx;*.xls"
  │      → nếu cancel → return Cancelled
  │
  ├─ 3. ArcToolSettings.Load()      — lấy LastScale, LastExcelFile
  │
  ├─ 4. ImportOptionsDialog         — WinForms dialog: Scale%, CreateNewView bool
  │      → nếu cancel → return Cancelled
  │
  ├─ 5. ExcelInteropService
  │      ├─ OpenFile(excelPath)
  │      ├─ GetActiveSheetName()    → sheetName (dùng cho view name)
  │      ├─ ExportPrintAreaAsHighResImage(tempPng)
  │      └─ Dispose()
  │
  ├─ 6. Transaction("ArcTool: Import Excel Image")
  │      ├─ GetOrCreateDraftingView(sheetName)   → targetView
  │      ├─ ImageType.Create(tempPng)            → imageType
  │      ├─ ImageInstance.Create(targetView, ...) → inst
  │      ├─ inst.Width  = scaledWidth
  │      ├─ inst.Height = scaledHeight
  │      └─ Commit()
  │
  ├─ 7. ArcToolSettings { LastExcelFile=, LastScale= }.Save()
  │
  ├─ 8. Khởi tạo ExcelRefreshHandler
  │      { StoredWidth=inst.Width, StoredHeight=inst.Height,
  │        TargetViewId=targetView.Id, ImageInstanceId=inst.Id }
  │
  ├─ 9. _updateEvent = ExternalEvent.Create(handler)
  │
  ├─ 10. SetupWatcher(excelPath)     — FileSystemWatcher bắt đầu theo dõi
  │
  ├─ 11. ShowStatusWindow(excelPath) — SyncStatusWindow hiện góc dưới phải
  │
  └─ 12. File.Delete(tempPng)        — dọn temp file
```

### 4.2 Luồng khi Excel thay đổi (Manual Sync)

```
FileSystemWatcher.Changed / Renamed (background thread)
  └─ lock(_debounceLock): Stop + Dispose timer cũ, tạo timer mới 2500ms
       └─ Timer.Elapsed (background thread)
            └─ _statusWindow?.SetStatus(hasChanges: true)
               [SetStatus() tự Dispatcher.Invoke → safe]
               → StatusDot đỏ, BtnApply enabled

User nhấn "Cập nhật" trong SyncStatusWindow (UI thread)
  └─ _updateEvent.Raise()   — non-blocking
       └─ Revit main thread: ExcelRefreshHandler.Execute(UIApplication app)
            ├─ doc = app.ActiveUIDocument.Document
            ├─ ExcelInteropService.Export → tempPng
            ├─ Lấy existingInst = doc.GetElement(ImageInstanceId) as ImageInstance
            ├─ Guard: if (existingInst?.IsValidObject == true)
            │    StoredWidth  = existingInst.Width   ← TRƯỚC KHI xóa
            │    StoredHeight = existingInst.Height
            ├─ Transaction
            │    doc.Delete(existingInst.Id)
            │    ImageType.Create(tempPng) → newType
            │    ImageInstance.Create(targetView, ...) → newInst
            │    newInst.Width  = StoredWidth
            │    newInst.Height = StoredHeight
            │    ImageInstanceId = newInst.Id     ← cập nhật reference
            │    Commit()
            ├─ File.Delete(tempPng)
            └─ _statusWindow?.SetStatus(hasChanges: false)  ← tick xanh
```

---

## 5. CODE ĐẦY ĐỦ — CÁC FILE CẦN SỬA/TẠO MỚI

### 5.1 SyncStatusWindow.xaml.cs — V2.0 (THAY THẾ TOÀN BỘ)

> **Vấn đề cần sửa:** Thêm `SetStatus(bool)` thread-safe. Sửa constructor để nhận `watchedFilePath` (không phải `changedFilePath` — window là persistent, không phải toast).
>>>>>>> 175ed814420ed8c34d98e53dd28aad238ef98aa5

```csharp
using System;
using System.IO;
using System.Windows;
using System.Windows.Media;

<<<<<<< HEAD
// 2. Mọi logic chính phải có try-catch
try { ... }
catch (Autodesk.Revit.Exceptions.OperationCanceledException) { return Result.Cancelled; }
catch (Exception ex) { message = ex.Message; return Result.Failed; }

// 3. Thông báo kết quả
TaskDialog.Show("ArcTool", "...");
uidoc.Application.Application.WriteJournalComment("...", true);

// 4. Revit 2026: dùng long, không dùng int cho ElementId
elem.Category.Id.Value == (long)BuiltInCategory.OST_Walls  // ĐÚNG
(int)elem.Category.Id.Value == (int)BuiltInCategory.OST_Walls  // SAI

// 5. Đơn vị: Revit dùng feet nội bộ
// 1 foot = 304.8mm. Convert: UnitUtils.ConvertToInternalUnits(mm, UnitTypeId.Millimeters)

// 6. COM Interop: LUÔN release theo thứ tự child → parent
// LUÔN null field gốc sau ReleaseObject, KHÔNG null trong ReleaseObject()
// KHÔNG ReleaseComObject sau Delete() — COM đã tự cleanup
=======
namespace ArcTool.Core.UI
{
    /// <summary>
    /// Persistent floating window — hiện sau khi import thành công, tồn tại
    /// suốt thời gian FileSystemWatcher đang chạy.
    ///
    /// Lifecycle:
    ///   Show()         → sau khi import lần đầu thành công
    ///   SetStatus(true) → khi watcher phát hiện file thay đổi (tick đỏ)
    ///   SetStatus(false)→ sau khi refresh thành công (tick xanh)
    ///   Close()        → khi user nhấn ✕ HOẶC khi StopWatcher() được gọi
    /// </summary>
    public partial class SyncStatusWindow : Window
    {
        // Màu sắc — static để tránh tạo lại mỗi lần SetStatus
        private static readonly SolidColorBrush GreenBrush =
            new SolidColorBrush(Color.FromRgb(34, 197, 94));   // #22C55E
        private static readonly SolidColorBrush RedBrush =
            new SolidColorBrush(Color.FromRgb(239, 68, 68));   // #EF4444

        // Callback: user nhấn "Cập nhật" → raise ExternalEvent
        private readonly Action _onUpdateClicked;

        /// <param name="watchedFilePath">File Excel đang được theo dõi</param>
        /// <param name="onUpdateClicked">Action gọi _updateEvent.Raise()</param>
        public SyncStatusWindow(string watchedFilePath, Action onUpdateClicked)
        {
            InitializeComponent();
            _onUpdateClicked = onUpdateClicked
                ?? throw new ArgumentNullException(nameof(onUpdateClicked));

            // Hiển thị tên file ngắn gọn
            TxtFileName.Text = Path.GetFileName(watchedFilePath);

            // State ban đầu: xanh (vừa import xong = đã đồng bộ)
            SetStatus(hasChanges: false);

            // Cho phép kéo window
            MouseLeftButtonDown += (s, e) => DragMove();
        }

        private void Window_Loaded(object sender, RoutedEventArgs e)
        {
            // Đặt ở góc dưới phải, cách viền 20px, loại trừ Taskbar
            var area = SystemParameters.WorkArea;
            Left = area.Right  - Width  - 20;
            Top  = area.Bottom - Height - 20;
        }

        /// <summary>
        /// Cập nhật trạng thái tick xanh/đỏ. Thread-safe — có thể gọi từ bất kỳ thread nào.
        /// </summary>
        public void SetStatus(bool hasChanges)
        {
            // Guard cross-thread: không invoke nếu window đã đóng
            if (!IsLoaded) return;

            if (!Dispatcher.CheckAccess())
            {
                Dispatcher.Invoke(() => SetStatus(hasChanges));
                return;
            }

            StatusDot.Fill     = hasChanges ? RedBrush : GreenBrush;
            TxtStatus.Text     = hasChanges ? "File Excel đã thay đổi" : "Đã đồng bộ";
            TxtStatus.Foreground = hasChanges ? RedBrush : GreenBrush;
            BtnApply.IsEnabled = hasChanges;
        }

        private void BtnApply_Click(object sender, RoutedEventArgs e)
        {
            // Disable ngay để tránh double-click trong khi handler đang chạy
            BtnApply.IsEnabled = false;
            _onUpdateClicked.Invoke();
            // Tick sẽ chuyển xanh khi ExcelRefreshHandler.Execute() hoàn tất
        }

        private void BtnClose_Click(object sender, RoutedEventArgs e)
        {
            Close(); // ExcelToRevitCommand đăng ký Closed event để StopWatcher()
        }
    }
}
```

### 5.2 ExcelToRevitCommand.cs — V3.0 (THAY THẾ TOÀN BỘ)

```csharp
using System;
using System.IO;
using System.Windows.Forms;  // OpenFileDialog, Form, Label, NumericUpDown, Button, CheckBox
using Autodesk.Revit.Attributes;
using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using Autodesk.Revit.UI.Events;
using ArcTool.Core.Services;
using ArcTool.Core.UI;

// Tránh ambiguity: Revit TaskDialog vs WinForms
using RevitTaskDialog = Autodesk.Revit.UI.TaskDialog;
using Timer = System.Timers.Timer;

namespace ArcTool.Core.Commands
{
    [Transaction(TransactionMode.Manual)]
    public class ExcelToRevitCommand : IExternalCommand
    {
        // ── Static state: tồn tại sau khi Execute() kết thúc ─────────────────
        private static FileSystemWatcher    _watcher;
        private static Timer                _debounceTimer;
        private static readonly object      _debounceLock = new object();
        private static ExcelRefreshHandler  _handler;
        private static ExternalEvent        _updateEvent;
        private static SyncStatusWindow     _statusWindow;

        // ── Entry point ───────────────────────────────────────────────────────
        public Result Execute(ExternalCommandData commandData, ref string message, ElementSet elements)
        {
            UIApplication uiApp = commandData.Application;
            UIDocument    uidoc = uiApp.ActiveUIDocument;
            Document      doc   = uidoc.Document;

            // Dừng session cũ hoàn toàn trước khi bắt đầu session mới
            StopWatcher();

            try
            {
                // ── BƯỚC 1: Chọn file Excel ──────────────────────────────────
                string excelPath = PickExcelFile();
                if (string.IsNullOrEmpty(excelPath)) return Result.Cancelled;

                // ── BƯỚC 2: Load settings + hỏi options ─────────────────────
                var settings   = ArcToolSettings.Load();
                var options    = ShowImportOptions(settings.LastScale);
                if (options == null) return Result.Cancelled;

                // ── BƯỚC 3: Export Excel → PNG ───────────────────────────────
                string tempPng = Path.Combine(Path.GetTempPath(),
                    $"ArcTool_Excel_{Guid.NewGuid():N}.png");

                string sheetName;
                using (var svc = new ExcelInteropService())
                {
                    if (!svc.OpenFile(excelPath))
                    {
                        RevitTaskDialog.Show("ArcTool Error",
                            "Không thể mở file Excel. Kiểm tra file có đang mở không.");
                        return Result.Failed;
                    }

                    sheetName = svc.GetActiveSheetName();

                    if (!svc.ExportPrintAreaAsHighResImage(tempPng))
                    {
                        RevitTaskDialog.Show("ArcTool Error",
                            "Không thể export ảnh từ Excel. Kiểm tra Print Area đã được thiết lập chưa.");
                        return Result.Failed;
                    }
                }

                // ── BƯỚC 4: Import ảnh vào Revit ─────────────────────────────
                ImageInstance inst = null;
                Autodesk.Revit.DB.View targetView = null;

                using (var tx = new Transaction(doc, "ArcTool: Import Excel Image"))
                {
                    tx.Start();

                    // Lấy hoặc tạo Drafting View
                    string viewName = string.IsNullOrWhiteSpace(sheetName)
                        ? "Excel Import"
                        : sheetName;
                    targetView = options.CreateNewView
                        ? GetOrCreateDraftingView(doc, viewName)
                        : doc.ActiveView;

                    if (targetView == null)
                    {
                        tx.RollBack();
                        RevitTaskDialog.Show("ArcTool Error", "Không tạo được Drafting View.");
                        return Result.Failed;
                    }

                    // Tạo ImageType từ file PNG
                    var imgOptions = new ImageTypeOptions(tempPng,
                        useRelativePath: false,
                        importMethodEnum: ImageTypeImportMethod.Import);
                    ImageType imageType = ImageType.Create(doc, imgOptions);

                    // Đặt ảnh vào View
                    var instOptions = new ImageInstancePlacementOptions
                    {
                        PlacementPoint = BoxPlacement.TopLeft
                    };
                    inst = ImageInstance.Create(doc, targetView, imageType.Id, instOptions);

                    // Áp dụng scale từ dialog
                    double scaleFactor = options.ScalePercent / 100.0;
                    inst.Width  = inst.Width  * scaleFactor;
                    inst.Height = inst.Height * scaleFactor;

                    tx.Commit();
                }

                // ── BƯỚC 5: Lưu settings ─────────────────────────────────────
                settings.LastExcelFile = excelPath;
                settings.LastScale     = options.ScalePercent;
                settings.Save();

                // ── BƯỚC 6: Khởi tạo ExcelRefreshHandler ─────────────────────
                _handler = new ExcelRefreshHandler
                {
                    // Smart Scale: lưu kích thước thực từ Revit làm baseline
                    StoredWidth      = inst.Width,
                    StoredHeight     = inst.Height,
                    TargetViewId     = targetView.Id,
                    ImageInstanceId  = inst.Id,
                    ExcelPath        = excelPath
                };

                _updateEvent = ExternalEvent.Create(_handler);

                // Callback: sau khi refresh xong → đổi tick về xanh
                _handler.OnRefreshComplete = () =>
                    _statusWindow?.SetStatus(hasChanges: false);

                // ── BƯỚC 7: Bắt đầu theo dõi file ────────────────────────────
                SetupWatcher(excelPath);

                // ── BƯỚC 8: Hiện SyncStatusWindow ────────────────────────────
                ShowStatusWindow(excelPath);

                // ── BƯỚC 9: Dọn temp file ─────────────────────────────────────
                TryDeleteFile(tempPng);

                return Result.Succeeded;
            }
            catch (Autodesk.Revit.Exceptions.OperationCanceledException)
            {
                return Result.Cancelled;
            }
            catch (Exception ex)
            {
                message = ex.Message;
                return Result.Failed;
            }
        }

        // ── Helper: Chọn file Excel ───────────────────────────────────────────
        private static string PickExcelFile()
        {
            using var dlg = new OpenFileDialog
            {
                Title  = "ArcTool — Chọn file Excel",
                Filter = "Excel Files (*.xlsx;*.xls)|*.xlsx;*.xls",
                CheckFileExists = true
            };
            return dlg.ShowDialog() == DialogResult.OK ? dlg.FileName : null;
        }

        // ── Helper: Import Options dialog ─────────────────────────────────────
        private static ImportOptions ShowImportOptions(double defaultScale)
        {
            using var form = new ImportOptionsForm(defaultScale);
            if (form.ShowDialog() != DialogResult.OK) return null;
            return new ImportOptions
            {
                ScalePercent  = form.ScalePercent,
                CreateNewView = form.CreateNewView
            };
        }

        // ── Helper: GetOrCreateDraftingView ───────────────────────────────────
        /// <summary>Phải gọi trong Transaction đang active.</summary>
        private static ViewDrafting GetOrCreateDraftingView(Document doc, string name)
        {
            var existing = new FilteredElementCollector(doc)
                .OfClass(typeof(ViewDrafting))
                .Cast<ViewDrafting>()
                .FirstOrDefault(v =>
                    string.Equals(v.Name, name, StringComparison.OrdinalIgnoreCase));

            if (existing != null) return existing;

            var familyType = new FilteredElementCollector(doc)
                .OfClass(typeof(ViewFamilyType))
                .Cast<ViewFamilyType>()
                .FirstOrDefault(t => t.ViewFamily == ViewFamily.Drafting);

            if (familyType == null) return null;

            var view = ViewDrafting.Create(doc, familyType.Id);
            try   { view.Name = name; }
            catch { view.Name = $"{name}_{DateTime.Now:HHmmss}"; }

            return view;
        }

        // ── FileSystemWatcher ─────────────────────────────────────────────────
        private static void SetupWatcher(string excelPath)
        {
            string dir  = Path.GetDirectoryName(excelPath);
            string file = Path.GetFileName(excelPath);
            if (string.IsNullOrEmpty(dir)) return;

            _watcher = new FileSystemWatcher(dir, file)
            {
                // FileName bắt buộc để Renamed event (Office save pattern) hoạt động
                NotifyFilter        = NotifyFilters.LastWrite
                                    | NotifyFilters.Size
                                    | NotifyFilters.FileName,
                EnableRaisingEvents = true
            };

            _watcher.Changed += (s, e) => ScheduleStatusUpdate(e.FullPath);
            _watcher.Renamed += (s, e) => ScheduleStatusUpdate(e.FullPath);
        }

        private static void ScheduleStatusUpdate(string changedFilePath)
        {
            // lock bắt buộc: Changed/Renamed có thể fire đồng thời từ nhiều thread
            lock (_debounceLock)
            {
                _debounceTimer?.Stop();
                _debounceTimer?.Dispose();
                _debounceTimer = new Timer(2500) { AutoReset = false };
                _debounceTimer.Elapsed += (s, args) =>
                {
                    // CHỈ đổi màu tick → đỏ. KHÔNG tự refresh ảnh.
                    // User chủ động nhấn "Cập nhật" để trigger ExternalEvent.
                    _statusWindow?.SetStatus(hasChanges: true);
                };
                _debounceTimer.Start();
            }
        }

        // ── SyncStatusWindow ──────────────────────────────────────────────────
        private static void ShowStatusWindow(string excelPath)
        {
            // SyncStatusWindow là WPF window → phải tạo trên UI thread.
            // Execute() của IExternalCommand chạy trên Revit main thread (= WPF UI thread)
            // → Tạo trực tiếp, không cần Dispatcher.Invoke.
            _statusWindow = new SyncStatusWindow(
                watchedFilePath: excelPath,
                onUpdateClicked: () => _updateEvent?.Raise()
            );

            _statusWindow.Closed += (s, e) =>
            {
                // User nhấn ✕ → dừng watcher và cleanup
                StopWatcher();
            };

            _statusWindow.Show();
        }

        // ── Cleanup ───────────────────────────────────────────────────────────
        private static void StopWatcher()
        {
            // Watcher
            _watcher?.Dispose();
            _watcher = null;

            // Debounce timer
            lock (_debounceLock)
            {
                _debounceTimer?.Stop();
                _debounceTimer?.Dispose();
                _debounceTimer = null;
            }

            // External event — phải null để Execute() lần sau tạo lại với handler mới
            _updateEvent = null;
            _handler     = null;

            // Status window
            if (_statusWindow != null)
            {
                try { _statusWindow.Close(); } catch { }
                _statusWindow = null;
            }
        }

        private static void TryDeleteFile(string path)
        {
            try { if (File.Exists(path)) File.Delete(path); } catch { }
        }

        // ── Inner types ───────────────────────────────────────────────────────

        private class ImportOptions
        {
            public double ScalePercent  { get; set; }
            public bool   CreateNewView { get; set; }
        }
    }

    // ═══════════════════════════════════════════════════════════════════════
    //  ExcelRefreshHandler — IExternalEventHandler
    //  Chạy trên Revit main thread khi user nhấn "Cập nhật"
    // ═══════════════════════════════════════════════════════════════════════
    public class ExcelRefreshHandler : IExternalEventHandler
    {
        // Smart Scale state — đọc từ Revit instance, không phải ScaleFactor ban đầu
        public double  StoredWidth     { get; set; }
        public double  StoredHeight    { get; set; }
        public ElementId TargetViewId  { get; set; }
        public ElementId ImageInstanceId { get; set; }
        public string  ExcelPath       { get; set; }

        // Callback sau khi refresh hoàn tất → đổi tick xanh
        public Action OnRefreshComplete { get; set; }

        public string GetName() => "ArcTool: ExcelRefreshHandler";

        public void Execute(UIApplication app)
        {
            Document doc = app.ActiveUIDocument?.Document;
            if (doc == null) return;

            string tempPng = Path.Combine(Path.GetTempPath(),
                $"ArcTool_Refresh_{Guid.NewGuid():N}.png");

            try
            {
                // ── Export Excel → PNG ────────────────────────────────────────
                using (var svc = new ExcelInteropService())
                {
                    if (!svc.OpenFile(ExcelPath) ||
                        !svc.ExportPrintAreaAsHighResImage(tempPng))
                    {
                        RevitTaskDialog.Show("ArcTool Error",
                            "Không thể refresh: lỗi đọc file Excel.");
                        return;
                    }
                }

                // ── Smart Scale: đọc kích thước thực TRƯỚC KHI xóa ───────────
                using (var tx = new Transaction(doc, "ArcTool: Refresh Excel Image"))
                {
                    tx.Start();

                    var existingInst = doc.GetElement(ImageInstanceId) as ImageInstance;

                    if (existingInst != null && existingInst.IsValidObject)
                    {
                        // Phản ánh mọi resize thủ công của user sau lần import
                        StoredWidth  = existingInst.Width;
                        StoredHeight = existingInst.Height;

                        doc.Delete(existingInst.Id);
                        // KHÔNG ReleaseComObject — đây là Revit managed object, không phải COM
                    }

                    // ── Tạo ImageType + Instance mới ─────────────────────────
                    var imgOptions = new ImageTypeOptions(tempPng,
                        useRelativePath: false,
                        importMethodEnum: ImageTypeImportMethod.Import);
                    ImageType newType = ImageType.Create(doc, imgOptions);

                    var view = doc.GetElement(TargetViewId) as Autodesk.Revit.DB.View;
                    if (view == null)
                    {
                        tx.RollBack();
                        RevitTaskDialog.Show("ArcTool Error", "Target View không còn tồn tại.");
                        return;
                    }

                    var instOptions = new ImageInstancePlacementOptions
                    {
                        PlacementPoint = BoxPlacement.TopLeft
                    };
                    ImageInstance newInst = ImageInstance.Create(
                        doc, view, newType.Id, instOptions);

                    // Áp dụng lại kích thước user đã set (smart scale)
                    if (StoredWidth > 0 && StoredHeight > 0)
                    {
                        newInst.Width  = StoredWidth;
                        newInst.Height = StoredHeight;
                    }

                    // Cập nhật reference cho lần refresh tiếp theo
                    ImageInstanceId = newInst.Id;

                    tx.Commit();
                }

                // ── Thông báo hoàn tất → tick xanh ───────────────────────────
                OnRefreshComplete?.Invoke();
            }
            catch (Exception ex)
            {
                RevitTaskDialog.Show("ArcTool Error",
                    $"[ExcelRefreshHandler] Refresh thất bại: {ex.Message}");
            }
            finally
            {
                TryDeleteFile(tempPng);
            }
        }

        private static void TryDeleteFile(string path)
        {
            try { if (File.Exists(path)) File.Delete(path); } catch { }
        }
    }

    // ═══════════════════════════════════════════════════════════════════════
    //  ImportOptionsForm — WinForms dialog
    //  Scale% + CreateNewView checkbox
    // ═══════════════════════════════════════════════════════════════════════
    internal class ImportOptionsForm : System.Windows.Forms.Form
    {
        public double ScalePercent  { get; private set; }
        public bool   CreateNewView { get; private set; }

        private NumericUpDown _nudScale;
        private CheckBox      _chkNewView;

        public ImportOptionsForm(double defaultScale)
        {
            Text            = "ArcTool — Import Options";
            Size            = new System.Drawing.Size(320, 200);
            StartPosition   = FormStartPosition.CenterScreen;
            FormBorderStyle = FormBorderStyle.FixedDialog;
            MaximizeBox     = false;
            MinimizeBox     = false;

            var lblScale = new Label
            {
                Text     = "Scale (%):",
                Location = new System.Drawing.Point(20, 20),
                AutoSize = true
            };

            _nudScale = new NumericUpDown
            {
                Location  = new System.Drawing.Point(20, 45),
                Size      = new System.Drawing.Size(120, 25),
                Minimum   = 10,
                Maximum   = 500,
                Value     = (decimal)Math.Clamp(defaultScale, 10, 500),
                DecimalPlaces = 0
            };

            _chkNewView = new CheckBox
            {
                Text     = "Tạo Drafting View mới theo tên sheet",
                Location = new System.Drawing.Point(20, 85),
                AutoSize = true,
                Checked  = true
            };

            var btnOk = new Button
            {
                Text         = "OK",
                DialogResult = DialogResult.OK,
                Location     = new System.Drawing.Point(120, 130),
                Width        = 80
            };
            btnOk.Click += (s, e) =>
            {
                ScalePercent  = (double)_nudScale.Value;
                CreateNewView = _chkNewView.Checked;
            };

            var btnCancel = new Button
            {
                Text         = "Hủy",
                DialogResult = DialogResult.Cancel,
                Location     = new System.Drawing.Point(210, 130),
                Width        = 80
            };

            Controls.AddRange(new Control[]
                { lblScale, _nudScale, _chkNewView, btnOk, btnCancel });

            AcceptButton = btnOk;
            CancelButton = btnCancel;
        }
    }
}
>>>>>>> 175ed814420ed8c34d98e53dd28aad238ef98aa5
```

---

## 6. CHECKLIST SỬA TỪNG FILE

<<<<<<< HEAD
### Giai đoạn 1 — Trả nợ kỹ thuật (ƯU TIÊN CAO) — 5/5 HOÀN THÀNH ✅
- [x] Fix bug `long` cast trong `MultiCutCommand` — ✅ DONE
- [x] Fix `GetParamValue` tìm Instance trong `CreateVoidFromLinkCommand` — ✅ DONE
- [x] Fix `_lastUpdate` thành `static` trong `FilterManagerCommand` — ✅ DONE
- [x] Fix baseline update logic trong `ArrangeDimensionCommand` — ✅ DONE
- [x] Xóa `doc.Regenerate()` thừa trong `CreateVoidFromLinkCommand` — ✅ DONE

### Giai đoạn 1B — Cải thiện thêm (ƯU TIÊN TRUNG BÌNH) — 4/6 HOÀN THÀNH
- [x] **[BUG-E1]** Fix `ReleaseObject` + null field gốc trong `Dispose()` — ✅ DONE
- [x] **[BUG-E2]** Sửa comment sai "50x" → "35x" — ✅ DONE
- [x] **[BUG-E3]** Fix thứ tự release COM trong `finally` block — ✅ DONE
- [x] **[BUG-E4]** Thêm `_activeSheet` vào `Dispose()` — ✅ DONE
- [ ] Fix `activeView.Scale == 0` check trong `ArrangeDimensionCommand`
- [ ] Refactor `Idling` → `ExternalEvent` pattern trong `FilterManagerCommand`
=======
### File 1: `SyncStatusWindow.xaml.cs` — Thay thế toàn bộ
- [x] Sửa constructor: `watchedFilePath` (không phải `changedFilePath`)
- [x] Thêm `SetStatus(bool hasChanges)` với `Dispatcher.CheckAccess()` guard
- [x] Thêm `GreenBrush`, `RedBrush` static fields
- [x] Xóa `_onUpdateClicked` invoke ra khỏi `BtnApply_Click` (chỉ gọi action, không Close)
- [x] Giữ nguyên `Window_Loaded` (vị trí góc dưới phải)

### File 2: `ExcelToRevitCommand.cs` — Thay thế toàn bộ
- [x] Thêm `private static readonly object _debounceLock`
- [x] Sửa `StopWatcher()`: null `_updateEvent`, null `_handler`
- [x] Thêm `PickExcelFile()` dùng `OpenFileDialog`
- [x] Thêm `ShowImportOptions()` + `ImportOptionsForm` inner class
- [x] Implement pipeline đầy đủ trong `Execute()`
- [x] Thêm `ExcelRefreshHandler` class (hoặc tách ra file riêng)
- [x] Thêm `ImportOptionsForm` inner class
- [x] Thêm `ShowStatusWindow()` method
- [x] Thêm `TryDeleteFile()` method
- [x] Xóa `ReopenHandler` stub

### File 3: `SyncStatusWindow.xaml` — KHÔNG SỬA
> XAML hiện tại (có `StatusDot`, `TxtStatus`, `BtnApply IsEnabled=False`) đã đúng.

### File 4: `ExcelInteropService.cs` — KHÔNG SỬA
> V5.2 ổn định. COM bugs E1–E4 đã fix.

### File 5: `ArcToolSettings.cs` — KHÔNG SỬA
> JSON persistence hoạt động đúng.

---

## 7. COMPILE ERRORS SẼ GẶP VÀ CÁCH XỬ LÝ

| Error | Nguyên nhân | Cách sửa |
|---|---|---|
| `ImageTypeOptions` không tìm thấy | Revit API 2024+ | Thêm `using Autodesk.Revit.DB;` |
| `ImageTypeImportMethod` không tìm thấy | Enum mới | Check `revitapidocs.com/2026` |
| `ImageInstancePlacementOptions` không tìm thấy | Revit API 2024+ | Check `revitapidocs.com/2026` |
| `BoxPlacement` không tìm thấy | Enum mới | Check `revitapidocs.com/2026` |
| `FirstOrDefault` không có | Thiếu `using System.Linq` | Thêm `using System.Linq;` |
| CS0104 ambiguous | WinForms + Revit.UI | Dùng alias `using RevitTaskDialog = Autodesk.Revit.UI.TaskDialog;` |

> ⚠️ **Quan trọng:** Nếu `ImageTypeOptions`, `ImageInstancePlacementOptions` không tồn tại trong Revit 2026 API,
> tra cứu tại https://www.revitapidocs.com/2026/ để tìm API tương đương.
> Revit API thay đổi giữa các phiên bản — luôn verify trước khi code.

---

## 8. QUY TẮC LẬP TRÌNH (BẤT BIẾN)

```csharp
// 1. Transaction attribute bắt buộc
[Transaction(TransactionMode.Manual)]

// 2. Namespace alias tránh conflict
using RevitTaskDialog = Autodesk.Revit.UI.TaskDialog;

// 3. ElementId: long (không int)
elem.Category.Id.Value == (long)BuiltInCategory.OST_Walls

// 4. COM release: child → parent, null field gốc, KHÔNG release sau Delete()
// (xem Pattern 5 trong SKILL.md)

// 5. Background thread + Revit API: ExternalEvent bắt buộc
// FileSystemWatcher.Changed → ScheduleStatusUpdate() → SetStatus() [UI only]
// User click → _updateEvent.Raise() → ExcelRefreshHandler.Execute() [Revit API]

// 6. Smart Scale: đọc WIDTH/HEIGHT từ instance TRƯỚC KHI xóa
StoredWidth  = existingInst.Width;
StoredHeight = existingInst.Height;
// SAU ĐÓ mới doc.Delete(existingInst.Id)

// 7. WPF cross-thread: Dispatcher.CheckAccess() trước update
if (!Dispatcher.CheckAccess()) { Dispatcher.Invoke(() => SetStatus(x)); return; }

// 8. StopWatcher() luôn null _updateEvent + _handler
// để Execute() lần sau tạo handler mới với document đúng

// 9. SyncStatusWindow là persistent window (không phải toast)
// Tạo 1 lần sau import thành công, SetStatus() cập nhật state
// Close() khi user nhấn ✕ hoặc StopWatcher() được gọi

// 10. Debounce timer phải bọc trong lock(_debounceLock)
// FileSystemWatcher có thể fire từ nhiều thread pool threads
```

---

## 9. TRẠNG THÁI CÁC TÍNH NĂNG KHÁC (không thay đổi)

| Feature | File | Trạng thái |
|---|---|---|
| Create Void | `CreateVoidFromLinkCommand.cs` | ✅ V4.0 ổn định |
| Multi-Cut | `MultiCutCommand.cs` | ✅ V2.0 ổn định |
| Arrange Dimensions | `ArrangeDimensionCommand.cs` | ✅ V1.0 ổn định |
| Filter Manager | `FilterManagerCommand.cs` | ⚠️ V1.0 — thiếu Ribbon Button (Giai đoạn 2) |
| Excel Export Engine | `ExcelInteropService.cs` | ✅ V5.2 ổn định |
| Settings | `ArcToolSettings.cs` | ✅ V1.0 ổn định |
| Ribbon UI | `App.cs` | ⚠️ V5.2 — Filter Manager chưa có button |

---

## 10. ROADMAP

### Giai đoạn 3 — Excel to Revit — 🔧 IN PROGRESS (Session 5.12)
- [x] Core pipeline: Excel → PNG → ImageType → ImageInstance
- [x] Persist Scale (JSON)
- [x] Auto-Create Drafting View theo tên sheet
- [x] Manual Sync: FileSystemWatcher + SyncStatusWindow (tick đỏ/xanh)
- [x] Smart Scale: đọc kích thước thực từ Revit trước khi refresh
- [ ] **[Session 5.12]** Rebuild ExcelToRevitCommand V3.0 (fix ETR-01 → ETR-12)
>>>>>>> 175ed814420ed8c34d98e53dd28aad238ef98aa5

### Giai đoạn 2 — Filter Manager
- [ ] Implement logic Copy Filter: đọc `ParameterFilterElement` từ View nguồn
- [ ] Implement logic Paste Filter: `view.AddFilter()`, set Visibility/Override
- [ ] Hoàn thiện MVVM binding cho `FilterWindow`

<<<<<<< HEAD
### Giai đoạn 3 — Excel to Revit Image — **REFACTORED SESSION 5 — MERGED INTO ExcelToRevitCommand** ✅
- [x] Viết `ImageImportCommand.cs` (V1.0-1.3.1) — ✅ **DEPRECATED SESSION 5**
  - Nguyên nhân: Xóa ImageImportCommand (lỗi nhận định về API limitation)
  - ImageType.Create() **CÓ HỖ TRỢ** công khai trong Revit 2026 API
  - Chức năng merged vào ExcelToRevitCommand.cs (unified pipeline)
- [x] Refactor thành `ExcelToRevitCommand.cs` (V1.0 - Final) — ✅ **DONE SESSION 5**
  - ✅ Unified pipeline: Excel → PNG → ImageType → ImageInstance
  - ✅ Modern dialog (TableLayoutPanel, Segoe UI, flat buttons)
  - ✅ Auto-center at View center
  - ✅ Flexible scale adjustment (%)
  - ✅ Complete error handling per step
- [x] **Fix ambiguous reference** (TaskDialog + TextBox)
  - ✅ Added alias `RevitTaskDialog`
  - ✅ Explicit qualification for TextBox
  - ✅ 9 TaskDialog.Show() replacements
- [x] **Build successful + Production ready**
  - ✅ BUILD SUCCESSFUL — 0 errors, 0 warnings

=======
>>>>>>> 175ed814420ed8c34d98e53dd28aad238ef98aa5
### Giai đoạn 4 — Quick Dim (R&D)
- [ ] Nghiên cứu trích xuất `ReferenceArray` từ Face/Edge của Wall, Column, Beam
- [ ] Revit Dim qua Reference (khác AutoCAD qua điểm tọa độ)
- [ ] Xây dựng hàm `GetDimensionableReferences(Element)` trước khi gọi `NewDimension()`

---

<<<<<<< HEAD
## 8. API REFERENCES QUAN TRỌNG

| Class/Method | Ghi chú |
|---|---|
| `InstanceVoidCutUtils.AddInstanceVoidCut()` | Cắt element bằng Void |
| `InstanceVoidCutUtils.CanBeCutWithVoid()` | Kiểm tra trước khi cắt |
| `ElementTransformUtils.MoveElement()` | Di chuyển element |
| `RevitLinkInstance.GetTotalTransform()` | Transform matrix của Link |
| `Reference.CreateLinkReference()` | Chuyển Reference sang Linked Reference |
| `ParameterFilterElement` | API cho Filter Manager |
| `View.AddFilter()` / `View.GetFilters()` | Copy/Paste filter vào View |
| `ImageInstance.Create()` | Import ảnh vào Revit (Giai đoạn 3) |
| `doc.Create.NewDimension()` | Tạo Dimension (Giai đoạn 4) |
| `UnitUtils.ConvertToInternalUnits()` | Convert đơn vị |
| `Marshal.ReleaseComObject()` | Giải phóng COM object — release theo thứ tự child → parent |

> Tra cứu API bắt buộc tại: https://www.revitapidocs.com/2026/

---

## 9. PROJECT DOCUMENTATION FILES

### A. CONTEXT.md (file này) — Project Status & Roadmap
- **Mục đích:** Track dự án state, bug list, features completed, decisions made
- **Cập nhật:** Cuối mỗi session
- **Dùng khi:** Session start (read history), planning features, tracking progress
- **Sections:**
  - 1-2: Project info + folder structure
  - 3: Features completed (detailed status)
  - 4-6: Bug list + technical decisions
  - 7: Roadmap + phase completion status
  - 8: API references
  - 9: This section

### B. ArcTool_Instructions.md (NEW - SESSION 5.6) — Development Guide
- **Mục đích:** Hands-on reference để setup, debug, test, deploy plugin
- **Cập nhật:** Khi có thay đổi workflow hoặc setup instructions
- **Dùng khi:** First setup, debugging session, manual testing, deployment
- **Sections:**
  - Debug setup (VS config, Revit debugger)
  - Testing workflow (4 test procedures with assertions)
  - Ribbon UI layout diagram
  - Build & deployment checklist
  - Common development tasks (add command, add icon, modify ExcelToRevit)
  - Troubleshooting table (build errors + runtime issues)
  - References & links
  - Session workflow

### C. SKILL.md (NEW - SESSION 5.6) — Coding Patterns & Standards
- **Mục đích:** Reference guide cho coding patterns, best practices, standards dùng trong project
- **Cập nhật:** Khi discover new pattern hoặc improve existing ones
- **Dùng khi:** Writing new code, code review, refactoring, maintenance
- **Sections:**
  - 15 major coding patterns (Transaction, COM, Filters, UI, etc.)
  - Each pattern: code example + explanation + when to use
  - DO's & DON'Ts quick reference
  - Future recommendations (logging, DI, etc.)

---

## 10. WORKFLOW: CÁCH DÙNG 3 FILES NÀY

### Session Start ✅
1. **Mở CONTEXT.md** → read timestamp + latest updates
2. **Skim SKILL.md** → refresh memory về patterns (nếu chưa làm lâu)
3. **Refer to ArcTool_Instructions.md** → setup/debug as needed

### During Coding 💻
1. **Viết code:** Refer SKILL.md patterns
2. **Test:** Follow testing workflow từ ArcTool_Instructions.md
3. **Debug:** Check troubleshooting table

### Session End ✅
1. **Update CONTEXT.md:** Bug status, decisions made, timestamp
2. **Add to SKILL.md:** Any new patterns discovered
3. **Update ArcTool_Instructions.md:** If workflow changed
4. **Commit all 3 files** + source code

---

## 11. CÁCH PASTE VÀO CHAT AI

**Đầu mỗi session mới:**
1. Copy toàn bộ **CONTEXT.md** (chỉ file này)
2. Paste vào đầu chat với prompt
3. Ghi rõ: "Session [N]: Tôi muốn [task cụ thể]"
4. **Ghi chú:** Bạn (AI) sẽ tự refer SKILL.md + ArcTool_Instructions.md từ workspace

**Example:**
```
# COPILOTWORKSPACE CONTEXT
- Projects targeting: .NET 8
- IDE: Visual Studio Enterprise 2026

[PASTE CONTEXT.md HERE]

---

**Session 6 Task:** Tôi muốn implement Filter Copy/Paste feature theo SKILL.md pattern #2 (Transaction Group)
```

---

## 12. FILE UPDATE CHECKLIST

### Khi add New Feature → Update cái gì?
- [ ] **CONTEXT.md** Section 3 (thêm feature description)
- [ ] **CONTEXT.md** Section 7 (update phase status)
- [ ] **SKILL.md** (thêm pattern nếu là pattern mới)
- [ ] **ArcTool_Instructions.md** Section 5 (add to common tasks nếu relevant)

### Khi fix Bug → Update cái gì?
- [ ] **CONTEXT.md** Section 4 (change [ ] → [x], remove from BUG list)
- [ ] **CONTEXT.md** Section 7 (update phase completion %)
- [ ] **SKILL.md** (update pattern explanation nếu bug liên quan)

### Khi discover New Pattern → Update cái gì?
- [ ] **SKILL.md** (thêm section mới, numbered incrementally)
- [ ] **SKILL.md** DO's/DON'Ts (thêm relevant points)
- [ ] **ArcTool_Instructions.md** Section 5 (thêm task example nếu applicable)

### Khi change Setup/Workflow → Update cái gì?
- [ ] **ArcTool_Instructions.md** (relevant section)
- [ ] **CONTEXT.md** Section 2 (folder structure nếu thay đổi)

---

*ArcTool © 2026 — Internal development documentation*

=======
## 11. API REFERENCES

| Class/Method | Ghi chú |
|---|---|
| `ImageType.Create()` | Tạo ImageType từ PNG — verify signature tại revitapidocs.com/2026 |
| `ImageTypeOptions` | Constructor params — verify tại revitapidocs.com/2026 |
| `ImageInstance.Create()` | Đặt ảnh vào View |
| `ImageInstance.Width/Height` | Đọc TRƯỚC khi xóa (Smart Scale) |
| `ImageInstance.IsValidObject` | Guard trước khi đọc/xóa trong handler |
| `ViewDrafting.Create(doc, typeId)` | Tạo Drafting View mới |
| `IExternalEventHandler + ExternalEvent` | Bridge background → Revit API |
| `FileSystemWatcher` | `NotifyFilters.FileName` bắt buộc cho Office save pattern |
| `Dispatcher.CheckAccess()` | Guard cross-thread WPF update |

> **Tra cứu API:** https://www.revitapidocs.com/2026/

---

*ArcTool © 2026 — Session 5.12: ExcelToRevitCommand V3.0 Rebuild*
>>>>>>> 175ed814420ed8c34d98e53dd28aad238ef98aa5
