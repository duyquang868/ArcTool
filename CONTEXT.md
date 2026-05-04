# ARCTOOL — AI SESSION CONTEXT
> Paste file này vào ĐẦU mỗi session chat mới với AI.
> Cập nhật sau mỗi session làm việc.
> Last updated: 2026-05-04 — Session 6.4: Phase 2 complete — Services/ExcelSyncEngine.cs V1.0 ✅ build success (BUG-E6 fixed)

---

## 1. THÔNG TIN DỰ ÁN

| Mục | Chi tiết |
|---|---|
| Tên dự án | ArcTool |
| Namespace chính | `ArcTool.Core` |
| Nền tảng | Autodesk Revit 2026 (API 2026) |
| Ngôn ngữ | C# / .NET 8.0 |
| IDE | Visual Studio Enterprise 2026 |
| UI Framework | WPF (modeless window) + WinForms (dialog) |
| Icon/Resource | `Properties/Resources.resx` (Access Modifier: Public) |
| Revit Unit System | `UnitTypeId` (ForgeTypeId) — KHÔNG dùng `DisplayUnitType` cũ |

---

## 2. CẤU TRÚC THƯ MỤC

```
ArcTool/
├── ArcTool.slnx
├── CONTEXT.md
├── ArcTool_Instructions.md
├── SKILL.md
├── ArcTool.Core/
│   ├── App.cs                          ← Ribbon UI, entry point
│   ├── App.config
│   ├── ArcTool.Core.csproj
│   ├── Commands/
│   │   ├── CreateVoidFromLinkCommand.cs
│   │   ├── MultiCutCommand.cs
│   │   ├── ArrangeDimensionCommand.cs
│   │   ├── FilterManagerCommand.cs
│   │   └── ExcelToRevitCommand.cs      ← V1.0 stable (V3.0 đang trong roadmap)
│   ├── Services/
│   │   ├── ExcelInteropService.cs      ← V5.3 ✅ STABLE — thêm GetSheetNames/GetNamedRanges/ExportRegion
│   │   ├── ArcToolSettingsService.cs   ← V1.0 ✅ STABLE — Load/Save JSON, atomic write
│   │   └── ExcelSyncEngine.cs         ← V1.0 ✅ STABLE — CheckForChanges/ExecuteUpdate/GetOrCreateView
│   ├── UI/
│   │   ├── FilterWindow.xaml
│   │   └── FilterWindow.xaml.cs
│   ├── Utilities/
│   │   └── SelectionFilters.cs
│   ├── Models/
│   │   └── ExcelMapping.cs             ← V1.0 ✅ STABLE — POCO + enums ExcelRegionType/ExcelViewType
│   ├── Resources/
│   │   ├── icon_create_16.jpg
│   │   ├── icon_create_32.jpg
│   │   ├── icon_cut_16.png
│   │   └── icon_cut_32.png
│   └── Properties/
│       ├── Resources.resx
│       └── Resources.Designer.cs
```

---

## 3. TRẠNG THÁI HIỆN TẠI — CÁC TÍNH NĂNG

### A. Ribbon UI — `App.cs` (V5.1 — STABLE)
- Tab: `ArcTool`
- Panel 1: `Void Tools` → SplitButton "Void Manager"
  - Nút chính: **Create Void** → `CreateVoidFromLinkCommand`
  - Nút phụ (Separator): **Multi-Cut** → `MultiCutCommand`
- Panel 2: `Annotation Tools`
  - Nút: **Arrange Dimensions** → `ArrangeDimensionCommand`
- Panel 3: `Excel Tools`
  - Nút: **Excel to Revit** → `ExcelToRevitCommand`
- Helper: `ConvertToImageSource(Bitmap)` chuyển Resource sang WPF ImageSource

### B. Create Void — `CreateVoidFromLinkCommand.cs` (V4.0 — STABLE)
- Tự động tạo Generic Model Void tại TẤT CẢ dầm (Structural Framing) trong file Link
- User chọn: (1) Family Void qua WinForms dialog, (2) File Link qua PickObject
- Logic vị trí: midpoint của LocationCurve, áp dụng linkTransform
- Logic kích thước: Width/Height từ tham số dầm, Length = chiều dài dầm
- Tạo instance kiểu Face-Based (NewFamilyInstance với linkedFaceRef)
- ⚠️ BIẾT RỦI RO: Face-Based sẽ bị "unhosted" nếu Link reload với geometry thay đổi

### C. Multi-Cut — `MultiCutCommand.cs` (V2.0 — STABLE)
- Cắt Tường (Walls) + Cột Kết cấu + Cột Kiến trúc bằng Void đã tạo
- Broad Phase: BoundingBoxIntersectsFilter để lọc sơ bộ (tránh O(n²))
- Dùng InstanceVoidCutUtils.AddInstanceVoidCut()
- ⏳ TODO: Thêm Narrow Phase bằng Solid Intersection

### D. Arrange Dimensions — `ArrangeDimensionCommand.cs` (V1.0 — STABLE)
- Pick đường Dim gốc (baseline), sau đó pick liên tục các Dim tiếp theo
- Tự động tịnh tiến Dim cách đều nhau theo Snap Distance × View Scale
- Dùng TransactionGroup để gộp toàn bộ thao tác vào 1 lần Undo
- Filter: LinearDimensionSelectionFilter (chỉ cho phép chọn Linear Dim)

### E. Excel Export Engine — `ExcelInteropService.cs` (V5.3 ✅ STABLE — Session 6.3)
- Đọc file Excel (hidden mode), export Print Area hoặc UsedRange thành PNG
- Scale factor: 35x cố định
- Có IDisposable, COM release đúng thứ tự child → parent
- **Public API đầy đủ:**
  - `OpenFile(filePath)` — mở file Excel, hidden mode
  - `GetActiveSheetName()` — tên sheet đang active
  - `ExportPrintAreaAsHighResImage(outputPath)` — export Print Area / UsedRange của active sheet
  - `GetSheetNames()` ✅ V5.3 — list tất cả tên sheet trong workbook
  - `GetNamedRanges(sheetName)` ✅ V5.3 — list Named Ranges thuộc 1 sheet cụ thể
  - `ExportRegion(sheetName, regionName, outputPath)` ✅ V5.3 — export theo sheet + region (Named Range → Print Area → UsedRange fallback)
- **Quyết định thiết kế V5.3 đã chốt:**
  - `Sheets` và `Names` COM wrapper phải release riêng sau khi duyệt xong (không chỉ release từng item)
  - `ExportRegion()` swap `_activeSheet` tạm thời → gọi `ExportRangeInternal()` → restore trong `finally`
  - Restore `_activeSheet` **TRƯỚC KHI** release `ws` local — tránh trỏ vào COM đã revoke
  - Named Range lỗi (formula, deleted, cross-sheet) bị bỏ qua trong try-catch per item, không dừng iteration

### F. Excel to Revit — `ExcelToRevitCommand.cs` (V1.0 — STABLE, chờ V3.0)
- Pipeline hiện tại: Excel → PNG → ImageType.Create() → ImageInstance.Create()
- User chọn file Excel, nhập scale %, ảnh đặt tại tâm view
- Build thành công, 0 errors, 0 warnings
- ⏳ CẦN NÂNG CẤP lên V3.0 theo roadmap bên dưới

### G. Filter Manager — `FilterManagerCommand.cs` + `FilterWindow.xaml` (SKELETON)
- UI WPF modeless đã xong (FilterWindow với 2 DataGrid: Filters + Views)
- Command skeleton đã xong, dùng Idling event để real-time update
- ⏳ TODO: Implement logic Copy/Paste Filter thực sự bằng ParameterFilterElement API

### H. ExcelMapping Model — `Models/ExcelMapping.cs` (V1.0 ✅ STABLE — Session 6.1)
- POCO class, JsonSerializable với `System.Text.Json` — không phụ thuộc Revit API
- Namespace: `ArcTool.Core.Models`
- Enum `ExcelRegionType`: NamedRange / PrintArea / UsedRange
- Enum `ExcelViewType`: DraftingView / LegendView
- ⚠️ Enum đặt tên có prefix `Excel` — tránh collision với `Autodesk.Revit.DB.ViewType`
- Computed helpers `[JsonIgnore]`: `IsFirstImport`, `HasStoredDimensions`, `BuildViewName()`
- Sentinel values đã chốt: `ImageInstanceId = 0`, `StoredWidth/Height = 0.0`, `LastModified = DateTime.MinValue`
- `Region = null` (không phải `""`) → "chưa chọn Named Range, dùng PrintArea/UsedRange"

### I. Settings Service — `Services/ArcToolSettingsService.cs` (V1.0 ✅ STABLE — Session 6.2)
- Static class, không phụ thuộc Revit API ngoài `Document` (chỉ đọc `doc.PathName`)
- **Atomic write pattern**: ghi vào `.tmp` cùng thư mục → `File.Replace()` / `File.Move()`
  - `File.Replace()` nếu JSON đích đã tồn tại (lần 2 trở đi)
  - `File.Move()` nếu JSON chưa tồn tại (lần đầu tiên)
  - Cả hai đều atomic trên cùng NTFS volume → không bao giờ corrupt JSON giữa chừng
- **JsonSerializerOptions** cached dưới dạng `static readonly` — tránh allocate mỗi call
  - `JsonStringEnumConverter`: serialize enum thành string (`"DraftingView"`) thay vì số (`0`)
  - `PropertyNameCaseInsensitive = true`: tolerate case mismatch khi đọc JSON cũ
  - `WriteIndented = true`: JSON dễ đọc/debug bằng text editor
- **LoadMappings()**: trả về `List<ExcelMapping>` rỗng nếu file không tồn tại hoặc corrupt
  - Bắt `JsonException` riêng → log + backup `.corrupt_[timestamp]` (tối đa 5 bản)
  - Bắt `Exception` (IOException, UnauthorizedAccess) → log + trả về List rỗng
- **SaveMappings()**: throw `IOException` nếu ghi thất bại — caller phải hiện dialog
- **GetSettingsPath()**: throw `InvalidOperationException` khi `doc.PathName` rỗng
- **FileExists(mapping)**: check `File.Exists(mapping.FilePath)` — guard cho Status Dot vàng
- **HasFileChanged(mapping)**: so sánh `File.GetLastWriteTime()` > `mapping.LastModified`
  - Dùng **local time** (không phải UTC) — phải nhất quán với `DateTime.Now` khi gán `LastModified`
  - Trả về `false` (không throw) nếu file không tồn tại — caller dùng `FileExists()` riêng
- ⚠️ KNOWN LIMITATION: `JsonStringEnumConverter` sẽ throw `DeserializeException` nếu JSON cũ
  chứa enum dạng số nguyên (`"viewType": 0`) — cần migration logic khi upgrade từ phiên bản cũ hơn
- ⚠️ KNOWN LIMITATION: Atomic write không đảm bảo nếu `.rvt` và `.tmp` nằm khác volume
  (edge case không thực tế trong môi trường làm việc bình thường)

### J. Excel Sync Engine — `Services/ExcelSyncEngine.cs` (V1.0 ✅ STABLE — Session 6.4)
- Static class, không có mutable state — mọi dependency truyền qua parameter
- PHẢI gọi trong Revit API context (Execute() của IExternalCommand) — dùng Transaction nội bộ
- **Supporting types định nghĩa trong cùng file:**
  - `MappingSyncStatus` (sealed class): `FileExists`, `HasChanges`, `DotColor` — immutable, UI chỉ đọc
  - `SyncDotColor` (enum): `Green` / `Red` / `Yellow` — dùng để bind WPF Ellipse.Fill
- **Public API:**
  - `CheckForChanges(IEnumerable<ExcelMapping>)` → `IReadOnlyDictionary<string, MappingSyncStatus>`
    - Chỉ so sánh filesystem timestamp, không mở Excel, không đọc Revit
    - Key = `mapping.Id`, null/empty Id bị bỏ qua
  - `ExecuteUpdate(ExcelMapping, Document, List<ExcelMapping>)` → `bool`
    - Tự mở Transaction — caller KHÔNG wrap thêm
    - Soft failures (file không mở được, export thất bại) → return false, không throw
    - Hard failures (Revit API lỗi, IOException) → throw, caller hiện dialog
  - `GetOrCreateView(string viewName, ExcelViewType, Document)` → `View`
    - PHẢI gọi trong Transaction đang active
    - Dispatcher sang `GetOrCreateDraftingView` hoặc `GetOrCreateLegendView`
- **Quyết định thiết kế V1.0 đã chốt:**
  - Mapping mutation xảy ra SAU Commit — capture `committedInstanceId/Width/Height` trước, mutate sau
    → Nếu Commit fail, mapping giữ nguyên state cũ, JSON không bị ghi
  - `using RevitView = Autodesk.Revit.DB.View` alias bắt buộc — tránh CS0104 với `System.Windows.Forms.View`
  - `GetOrCreateLegendView()` ưu tiên tên `ArcTool_LegendTemplate`; fallback bất kỳ Legend View nào
  - `CheckForChanges()` gọi `ArcToolSettingsService.FileExists()` và `HasFileChanged()` — không duplicate logic
- ⚠️ KNOWN LIMITATION: Nếu Commit thành công nhưng `SaveMappings()` throw IOException
  → mapping trong memory đã được mutate nhưng JSON chưa được ghi. Revit restart sẽ mất sync state.

---

## 4. BUG REGISTRY — TRẠNG THÁI

### ✅ ĐÃ FIX (không cần xem lại)

| ID | File | Mô tả |
|---|---|---|
| BUG-01 | MultiCutCommand | `(int)elem.Category.Id.Value` → Integer Overflow. Fix: cast `(long)` |
| BUG-02 | CreateVoidFromLink | `GetParamValue` chỉ tìm Symbol, bỏ sót Instance. Fix: tìm Instance trước |
| BUG-03 | FilterManagerCommand | `_lastUpdate` instance field, reset mỗi lần chạy. Fix: `static` |
| BUG-04 | ArrangeDimensionCommand | Baseline cập nhật dù Dim lỗi. Fix: `bool` return từ Move |
| BUG-05 | CreateVoidFromLink | `doc.Regenerate()` thừa trước Commit. Fix: xóa |
| BUG-E1 | ExcelInteropService | `ReleaseObject` null local param vô nghĩa. Fix: null field gốc ở caller |
| BUG-E2 | ExcelInteropService | Comment "50x" sai, constant là 35.0. Fix: sửa comment |
| BUG-E3 | ExcelInteropService | Release chartObj sau Delete() — undefined behavior. Fix: KHÔNG release sau Delete |
| BUG-E4 | ExcelInteropService | `_activeSheet` không được release trong Dispose(). Fix: thêm vào đầu Dispose |
| BUG-E5 | ExcelToRevitCommand | Ambiguous reference TaskDialog + TextBox. Fix: alias RevitTaskDialog |
| BUG-E6 | ExcelSyncEngine | CS0104: `View` ambiguous giữa `Autodesk.Revit.DB.View` và `System.Windows.Forms.View` (do `<UseWindowsForms>true</UseWindowsForms>` trong .csproj). Fix: `using RevitView = Autodesk.Revit.DB.View` |

### ⏳ CÒN TỒN TẠI

| ID | File | Mô tả | Priority |
|---|---|---|---|
| BUG-06 | ArrangeDimensionCommand | Không check `activeView.Scale == 0` (3D view) | Medium |
| BUG-07 | FilterManagerCommand | Idling event là anti-pattern cho model lớn | Low |
| BUG-08 | CreateVoidFromLink | `SetParam("Height", -beamHeight)` gán âm là workaround | Low |

---

## 5. QUYẾT ĐỊNH KỸ THUẬT ĐÃ CHỐT

| Quyết định | Lý do | Trade-off |
|---|---|---|
| Face-Based Void | Void bám theo mặt phẳng dầm | Unhosted nếu Link reload |
| BoundingBox broad-phase | Tránh O(n²) dự án lớn | Chưa có narrow phase |
| TransactionGroup trong ArrangeDim | Gộp 1 lần Undo | — |
| WinForms cho Family dialog | Đơn giản, không cần MVVM | UI không đồng nhất WPF |
| COM release: child trước parent | Tránh InvalidComObjectException | — |
| KHÔNG ReleaseComObject sau Delete() | Delete đã revoke COM handle | — |
| JSON lưu cạnh file .rvt | Setting đi theo project folder | Mất nếu copy .rvt mà không copy JSON |
| Legend View: Duplicate thay vì Create | Revit API 2026 không có method tạo Legend mới | User phải tạo thủ công 1 Legend View rỗng làm template lần đầu |
| Enum prefix `Excel` (ExcelViewType, ExcelRegionType) | Tránh `CS0104` collision với `Autodesk.Revit.DB.ViewType` | Tên dài hơn — bắt buộc, không phải tuỳ chọn |
| Alias `RevitView` trong ExcelSyncEngine | Tránh `CS0104` collision với `System.Windows.Forms.View` (UseWindowsForms=true) | Phải dùng `RevitView` thay vì `View` trong toàn bộ file |
| Mapping mutation SAU Commit trong ExecuteUpdate() | Nếu Commit fail, mapping giữ nguyên state cũ — JSON không bị ghi sai | Capture committed values vào locals trước Commit, mutate mapping sau |
| Enum prefix `Excel` (ExcelViewType, ExcelRegionType) | Tránh `CS0104` collision với `Autodesk.Revit.DB.ViewType` | Tên dài hơn — bắt buộc, không phải tuỳ chọn |
| Atomic write: `.tmp` → `File.Replace()`/`File.Move()` | Không để JSON corrupt nếu crash | `.tmp` không được dọn nếu crash ở bước 2; vô hại |
| `JsonStringEnumConverter` cho enum fields | Forward-compatible khi thêm enum value mới; JSON dễ đọc | `DeserializeException` nếu JSON cũ chứa enum dạng số |
| `DateTime` local time cho `LastModified` | `File.GetLastWriteTime()` trả về local time; nhất quán | Nếu user đổi timezone, so sánh timestamp có thể sai |
| `JsonSerializerOptions` là `static readonly` | Tránh allocate object mỗi lần call Load/Save | Không thread-safe nếu có code sửa options — nhưng options là immutable sau init |
| File corrupt → backup `.corrupt_[timestamp]`, tối đa 5 bản | Giữ lại để debug, không để disk đầy | Nếu corrupt liên tục, vẫn tích lũy 5 file |
| `ExportRegion()` swap `_activeSheet` thay vì truyền ws vào `ExportRangeInternal()` | Không sửa code cũ đã stable; ExportRangeInternal dùng _activeSheet | Tạm thời thay đổi state của instance — được vì Revit single-thread |
| Release `Sheets`/`Names` COM wrapper sau forEach | COM wrapper là object riêng, không tự release khi GC | Pattern bổ sung so với Pattern 10 gốc trong SKILL.md |

---

## 6. ROADMAP — EXCEL TO REVIT V3.0

> **Đây là tính năng ưu tiên cao nhất của Session 6+**
> V1.0 stable đang hoạt động — V3.0 bổ sung 3 tính năng mới hoàn toàn.

### 6.1 Tổng quan 3 tính năng mới

| # | Tính năng | Mô tả ngắn |
|---|---|---|
| T1 | Auto-Create View | Tự động tạo Drafting View hoặc Legend theo tên sheet Excel |
| T2 | Change Detection | Phát hiện Excel thay đổi khi dialog mở, hiển thị trạng thái xanh/đỏ per row |
| T3 | Smart Scale Persistence | Lưu Width/Height tuyệt đối của ImageInstance, áp lại khi refresh |

---

### 6.2 Data Model — `ExcelMapping` (JSON)

File JSON lưu tại: **cùng folder với file .rvt**, tên `ArcTool_ExcelSync.json`

```json
{
  "Mappings": [
    {
      "id": "guid-string",
      "viewName": "BudgetOverview",
      "autoSync": true,
      "lastModified": "2024-04-22T14:53:00",
      "workSheet": "Budget Overview",
      "region": "ChartTest",
      "regionType": "NamedRange",
      "viewType": "LegendView",
      "filePath": "C:\\Project\\Chart-Sample.xlsx",
      "imageInstanceId": 12345,
      "storedWidth": 2.5,
      "storedHeight": 1.8
    }
  ]
}
```

**Logic Region:**
- `RegionType = "NamedRange"`: gọi `ExcelInteropService.GetNamedRanges(sheetName)`
- `RegionType = "PrintArea"`: dùng `worksheet.PageSetup.PrintArea`
- `RegionType = "UsedRange"`: fallback — `worksheet.UsedRange`
- Ưu tiên khi export: NamedRange → PrintArea → UsedRange (ExportRegion() xử lý tự động)

**Logic ViewName:**
- Dùng Print Area / UsedRange → `ViewName = SheetName`
- Dùng Named Range "ChartTest" trên sheet "Budget Overview" → `ViewName = "Budget Overview_ChartTest"`

---

### 6.3 UI Spec — Bảng chính (WPF DataGrid)

**Columns theo thứ tự:**

| # | Column | Control | Ghi chú |
|---|---|---|---|
| 1 | Select | CheckBox | Chọn nhiều dòng để thực hiện batch action |
| 2 | Status Dot | Ellipse fill | Xanh = synced, Đỏ = file Excel mới hơn LastModified, Vàng = file không tìm thấy |
| 3 | View Name | TextBlock (read-only) | Auto-generated, không cho user sửa trực tiếp |
| 4 | Auto Sync | CheckBox | true → nút Update per-row bị disabled |
| 5 | Last Modified | TextBlock | Format: "dd/MM/yyyy HH:mm" |
| 6 | WorkSheet | ComboBox | Dropdown — load sheet names từ Excel file |
| 7 | Region | ComboBox | Dropdown — Print Areas + Named Ranges của sheet đang chọn |
| 8 | View Type | ComboBox | "Drafting View" / "Legend View" |
| 9 | File Path | TextBlock + Browse button | Hiện tên file ngắn, tooltip = full path |
| 10 | Update | Button (✅ icon) | Xanh = synced, Đỏ = có thay đổi; disabled nếu AutoSync=true |

**Toolbar phía trên bảng:**
- Nút `+` : Thêm dòng mới
- Nút `−` : Xóa dòng(s) được chọn
- Nút `Update All` : Update tất cả dòng có Status đỏ

---

### 6.4 Luồng khi Dialog Mở

```
[ExcelToRevitCommand.Execute()]
  │
  ├─ 1. ArcToolSettingsService.LoadMappings(doc)
  │      → Deserialize thành List<ExcelMapping>
  │      → Nếu doc.PathName rỗng → throw → hiện dialog yêu cầu save file
  │
  ├─ 2. Với mỗi mapping:
  │      ├─ ArcToolSettingsService.FileExists(mapping)
  │      │    └─ false → Status = FileNotFound (Status Dot vàng)
  │      ├─ ArcToolSettingsService.HasFileChanged(mapping)
  │      │    └─ true → HasChanges = true → Status Dot đỏ
  │      └─ Nếu AutoSync = true && HasChanges = true && FileExists = true
  │           → ExcelSyncEngine.ExecuteUpdate(mapping, doc) tự động
  │
  └─ 3. Show dialog với bảng đã populate
```

---

### 6.5 Luồng Khi User Nhấn `+` (Thêm Dòng Mới)

```
User nhấn "+"
  │
  ├─ Tạo dòng mới với giá trị mặc định:
  │    AutoSync = false, ViewType = "DraftingView", Region = null
  │
  ├─ FilePath column: user click Browse button
  │    └─ OpenFileDialog → chọn .xlsx / .xls
  │         └─ Sau khi chọn:
  │              ├─ using (var svc = new ExcelInteropService())
  │              │    svc.OpenFile() → svc.GetSheetNames()   ← V5.3
  │              │    svc.Dispose() ngay sau khi đọc xong
  │              └─ WorkSheet dropdown: populate sheet names
  │
  ├─ User chọn WorkSheet
  │    └─ using (var svc = new ExcelInteropService())
  │         svc.OpenFile() → svc.GetNamedRanges(sheetName)  ← V5.3
  │         svc.Dispose() ngay sau khi đọc xong
  │       → Region dropdown: populate (Print Area + Named Ranges)
  │
  ├─ ViewName: tự điền = mapping.BuildViewName()
  │    → Tự cập nhật khi WorkSheet hoặc Region thay đổi
  │
  └─ User nhấn OK / nhấn Update
```

---

### 6.6 Luồng Khi Nhấn Update (Per Row hoặc Update All)

```
ExcelSyncEngine.ExecuteUpdate(ExcelMapping mapping, Document doc)
  │
  ├─ 1. Export Excel → Temp PNG
  │      using (var svc = new ExcelInteropService())
  │        svc.OpenFile(mapping.FilePath)
  │        svc.ExportRegion(mapping.WorkSheet, mapping.Region, tempPng)  ← V5.3
  │        svc.Dispose()
  │
  ├─ 2. Đọc StoredWidth/StoredHeight TRƯỚC KHI xóa ảnh cũ (Smart Scale)
  │      var existingInst = doc.GetElement(new ElementId(mapping.ImageInstanceId)) as ImageInstance
  │      if (existingInst != null && existingInst.IsValidObject)
  │      {
  │          storedWidth  = existingInst.Width
  │          storedHeight = existingInst.Height
  │      }
  │      else { storedWidth = mapping.StoredWidth; storedHeight = mapping.StoredHeight; }
  │
  ├─ 3. Transaction("ArcTool: Refresh Excel Image")
  │      ├─ GetOrCreateView(mapping.ViewName, mapping.ViewType, doc) → targetView
  │      ├─ if (existingInst valid) → doc.Delete(existingInst.Id)
  │      ├─ ImageType.Create(doc, tempPng) → imageType
  │      ├─ ImageInstance.Create(doc, targetView, ...) → newInst
  │      ├─ if (storedWidth > 0 && storedHeight > 0)
  │      │    newInst.Width = storedWidth; newInst.Height = storedHeight
  │      └─ Commit()
  │
  ├─ 4. Cập nhật mapping:
  │      mapping.LastModified    = DateTime.Now   ← local time
  │      mapping.ImageInstanceId = newInst.Id.Value
  │      mapping.StoredWidth     = newInst.Width
  │      mapping.StoredHeight    = newInst.Height
  │      ArcToolSettingsService.SaveMappings(doc, allMappings)
  │
  ├─ 5. Cập nhật UI: Status Dot → xanh
  │
  └─ 6. TryDeleteFile(tempPng)
```

---

### 6.7 Xử Lý File Excel Không Tìm Thấy

```
ArcToolSettingsService.FileExists(mapping) == false:
  ├─ Status Dot = màu vàng
  ├─ Nút Update = disabled
  ├─ Tooltip: "File không tìm thấy. Click để chọn lại đường dẫn."
  └─ User click icon warning:
       └─ OpenFileDialog → chọn file mới
            ├─ mapping.FilePath = newPath
            ├─ ArcToolSettingsService.SaveMappings(doc, mappings)
            └─ Re-check: HasFileChanged(mapping) → cập nhật Status Dot
```

---

### 6.8 ExcelInteropService — Public API Đầy Đủ (V5.3 ✅)

```csharp
// Mở file, bắt buộc gọi trước mọi method khác
public bool OpenFile(string filePath)

// Lấy tên sheet active — dùng cho V1.0 pipeline
public string GetActiveSheetName()

// Export Print Area / UsedRange của active sheet — dùng cho V1.0 pipeline
public bool ExportPrintAreaAsHighResImage(string outputPath)

// [V5.3] Lấy tất cả tên sheet trong workbook
// Dùng để populate WorkSheet ComboBox khi user chọn file Excel
// Release Sheets wrapper + từng Worksheet COM ngay sau khi duyệt
public List<string> GetSheetNames()

// [V5.3] Lấy Named Ranges thuộc 1 sheet cụ thể
// Named Range lỗi (formula, deleted, cross-sheet) bị skip — không dừng iteration
// Release Names wrapper + từng Name + từng Range COM
public List<string> GetNamedRanges(string sheetName)

// [V5.3] Export theo sheet + region cụ thể
// regionName = null/empty → fallback Print Area → UsedRange tự động
// Swap _activeSheet tạm, restore trong finally TRƯỚC KHI release ws
public bool ExportRegion(string sheetName, string regionName, string outputPath)

// Dispose: sheet → workbook → app, GC.Collect() × 2
public void Dispose()
```

**Pattern sử dụng V5.3 trong ExcelSyncEngine:**
```csharp
// Mở 1 lần, gọi method cần thiết, Dispose ngay
using (var svc = new ExcelInteropService())
{
    if (!svc.OpenFile(mapping.FilePath)) return false;
    bool ok = svc.ExportRegion(mapping.WorkSheet, mapping.Region, tempPng);
    // Dispose() tự gọi khi ra khỏi using
}
```

---

### 6.9 Tạo View Revit — API Notes (ĐÃ VERIFY)

#### Drafting View — ✅ API đầy đủ, stable

```csharp
ViewFamilyType draftingType = new FilteredElementCollector(doc)
    .OfClass(typeof(ViewFamilyType))
    .Cast<ViewFamilyType>()
    .First(t => t.ViewFamily == ViewFamily.Drafting);

ViewDrafting view = ViewDrafting.Create(doc, draftingType.Id);
view.Name = viewName;
```

#### Legend View — ⚠️ KHÔNG CÓ API TẠO MỚI — Dùng Workaround Duplicate

Revit API 2026 **không có method tạo Legend View mới từ đầu**. Workaround: `view.Duplicate(ViewDuplicateOption.WithDetailing)`. Yêu cầu: project phải có sẵn ít nhất 1 Legend View rỗng. Pattern implement: xem SKILL.md Pattern 9.

---

### 6.10 JSON Service — ArcToolSettingsService (✅ IMPLEMENTED — Session 6.2)

Public API:

```csharp
public static string GetSettingsPath(Document doc)                          // throw nếu doc.PathName rỗng
public static List<ExcelMapping> LoadMappings(Document doc)                 // trả về [] nếu không có/corrupt
public static void SaveMappings(Document doc, List<ExcelMapping> mappings)  // atomic write
public static bool FileExists(ExcelMapping mapping)                         // check file Excel tồn tại
public static bool HasFileChanged(ExcelMapping mapping)                     // so sánh timestamp local time
```

---

### 6.11 Checklist Implementation V3.0

```
Phase 1 — Models & Services (không phụ thuộc UI)
  [x] Tạo Models/ExcelMapping.cs ✅ Session 6.1 — build success
        NOTE: Enum ExcelViewType / ExcelRegionType (tránh collision DB.ViewType)
        NOTE: Region = null (không phải "") cho "chưa chọn Named Range"
        NOTE: LastModified default = DateTime.MinValue (file luôn "changed" lần đầu)
  [x] Tạo Services/ArcToolSettingsService.cs ✅ Session 6.2 — build success
        NOTE: Atomic write pattern (.tmp → Replace/Move)
        NOTE: JsonStringEnumConverter (enum → string trong JSON)
        NOTE: DateTime local time cho HasFileChanged() và LastModified
        NOTE: File corrupt → backup .corrupt_[timestamp], tối đa 5 bản
  [x] Mở rộng Services/ExcelInteropService.cs → V5.3 ✅ Session 6.3 — build success
        NOTE: GetSheetNames() — release Sheets wrapper sau forEach
        NOTE: GetNamedRanges() — release Names wrapper + từng Name + Range; skip lỗi
        NOTE: ExportRegion() — swap _activeSheet, restore TRƯỚC release ws
        NOTE: regionName = null → PrintArea → UsedRange fallback tự động
  [x] Verify Legend View creation API tại revitapidocs.com/2026 ✅ (không có Create())

Phase 2 — Logic Layer (không phụ thuộc UI) ✅ COMPLETE — Session 6.4
  [x] Viết Services/ExcelSyncEngine.cs — build success
        NOTE: BUG-E6 fix — alias `using RevitView = Autodesk.Revit.DB.View` tránh CS0104
        NOTE: MappingSyncStatus (sealed) + SyncDotColor (enum) định nghĩa trong cùng file
        NOTE: CheckForChanges() — chỉ filesystem, không mở Excel, không đọc Revit
        NOTE: ExecuteUpdate() tự mở Transaction — caller KHÔNG wrap thêm
        NOTE: Mapping mutation SAU Commit — capture locals trước, mutate sau Commit thành công
        NOTE: GetOrCreateView() là dispatcher; GetOrCreate*View() là private helpers trong Transaction

Phase 3 — UI
  [ ] Thiết kế UI/ExcelToRevitWindow.xaml (WPF DataGrid theo UI Spec 6.3)
  [ ] UI/ExcelToRevitWindow.xaml.cs:
        [ ] Load data → DataGrid (dùng LoadMappings + HasFileChanged + FileExists)
        [ ] Handle "+" / "-" buttons
        [ ] Handle Update per-row button
        [ ] Handle "Update All" button
        [ ] Handle Browse file button + reload GetSheetNames() sau khi chọn file mới
        [ ] Handle WorkSheet ComboBox selection → reload GetNamedRanges()
        [ ] Handle File Not Found warning click

Phase 4 — Integration
  [ ] Sửa ExcelToRevitCommand.cs để mở ExcelToRevitWindow thay vì pipeline cũ
  [ ] Test với file Excel có: nhiều sheet, Named Ranges, Print Areas
  [ ] Test Smart Scale: import → resize trong Revit → Update → kiểm tra kích thước giữ nguyên
```

---

## 7. ROADMAP TỔNG THỂ

### Giai đoạn 1 — Technical Debt ✅ HOÀN THÀNH
Tất cả 5 bug nghiêm trọng + 4 COM bug đã fix.

### Giai đoạn 2 — Filter Manager ⏳ CHỜ
- Implement Copy/Paste Filter bằng ParameterFilterElement API
- Hoàn thiện MVVM binding

### Giai đoạn 3 — Excel to Revit V3.0 🔧 ĐANG TIẾN HÀNH
- Phase 1A ✅: ExcelMapping.cs (Session 6.1)
- Phase 1B ✅: ArcToolSettingsService.cs (Session 6.2)
- Phase 1C ✅: ExcelInteropService.cs V5.3 (Session 6.3)
- Phase 2 ✅: ExcelSyncEngine.cs V1.0 (Session 6.4)
- Phase 3 ⏳: ExcelToRevitWindow.xaml + .cs ← **NEXT SESSION**

### Giai đoạn 4 — Quick Dim (R&D) 📋 TƯƠNG LAI
- Nghiên cứu ReferenceArray extraction từ Wall, Column, Beam

---

## 8. QUY TẮC LẬP TRÌNH (BẮT BUỘC)

```csharp
// 1. Transaction attribute
[Transaction(TransactionMode.Manual)]

// 2. Namespace alias tránh conflict
using RevitTaskDialog = Autodesk.Revit.UI.TaskDialog;  // tránh conflict với WinForms
using RevitView = Autodesk.Revit.DB.View;               // tránh conflict với System.Windows.Forms.View
                                                         // BẮT BUỘC trong ExcelSyncEngine và các file
                                                         // import cả Autodesk.Revit.DB lẫn UseWindowsForms

// 3. ElementId: luôn dùng long
elem.Category.Id.Value == (long)BuiltInCategory.OST_Walls  // ĐÚNG
(int)elem.Category.Id.Value                                 // SAI

// 4. COM release: child → parent, null field gốc, KHÔNG release sau Delete()
if (_activeSheet != null) { ReleaseObject(_activeSheet); _activeSheet = null; }

// 5. COM wrapper (Sheets, Names): phải release riêng sau forEach
Sheets sheets = _workbook.Worksheets;
// ... duyệt xong ...
Marshal.ReleaseComObject(sheets); // KHÔNG bỏ qua bước này

// 6. Quick filter trước slow filter
new FilteredElementCollector(doc)
    .OfClass(typeof(Wall))          // quick — index-based
    .OfCategory(...)                // quick
    .Where(w => ...)                // slow — sau cùng

// 7. Smart Scale pattern — đọc kích thước TRƯỚC KHI xóa
if (existingInst?.IsValidObject == true)
{
    storedWidth  = existingInst.Width;
    storedHeight = existingInst.Height;
    doc.Delete(existingInst.Id);
}

// 8. JSON: atomic write — KHÔNG dùng File.WriteAllText() trực tiếp
// Dùng ArcToolSettingsService.SaveMappings() — đã xử lý atomic write

// 9. DateTime: LUÔN dùng DateTime.Now (local) cho LastModified
// HasFileChanged() dùng File.GetLastWriteTime() cũng trả về local

// 10. Enum trong Models PHẢI có prefix để tránh collision với Revit API
// Sai:  public enum ViewType   → CS0104
// Đúng: public enum ExcelViewType

// 11. ExcelInteropService: Dispose ngay sau khi đọc xong
using (var svc = new ExcelInteropService())
{
    svc.OpenFile(path);
    var sheets = svc.GetSheetNames();
} // Dispose() tự gọi — không giữ Excel mở lâu hơn cần
```

---

## 9. API REFERENCES QUAN TRỌNG

| Class/Method | Ghi chú |
|---|---|
| `ImageType.Create(doc, ImageTypeOptions)` | Tạo ImageType từ PNG — V2026 stable |
| `ImageInstance.Create(doc, view, typeId, opts)` | Đặt ảnh vào View |
| `ImageInstance.Width` / `.Height` | Đọc kích thước thực (feet) — Smart Scale |
| `ImageInstance.IsValidObject` | Guard trước khi đọc/xóa |
| `ViewDrafting.Create(doc, typeId)` | Tạo Drafting View — API đầy đủ, stable |
| `view.Duplicate(ViewDuplicateOption.WithDetailing)` | Tạo Legend View mới — workaround bắt buộc |
| `ViewType.Legend` | Dùng để filter/đọc Legend View hiện có, không có Create() |
| `ParameterFilterElement` | API cho Filter Manager |
| `View.AddFilter()` / `View.GetFilters()` | Copy/Paste filter |
| `Marshal.ReleaseComObject()` | COM release — child trước parent |
| `File.GetLastWriteTime(path)` | Check timestamp Excel file — trả về local time |
| `File.Replace(source, dest, null)` | Atomic rename — NTFS, cùng volume |
| `ArcToolSettingsService.LoadMappings(doc)` | Load JSON — không throw nếu corrupt |
| `ArcToolSettingsService.SaveMappings(doc, list)` | Save JSON — atomic write |
| `ArcToolSettingsService.HasFileChanged(mapping)` | So sánh timestamp local time |
| `ExcelInteropService.GetSheetNames()` | [V5.3] List tên sheet trong workbook |
| `ExcelInteropService.GetNamedRanges(sheetName)` | [V5.3] List Named Ranges của 1 sheet |
| `ExcelInteropService.ExportRegion(sheet, region, path)` | [V5.3] Export vùng cụ thể → PNG |

> **Tra cứu API bắt buộc:** https://www.revitapidocs.com/2026/

---

*ArcTool © 2026 — Internal development documentation*
*Session 6.4: Phase 2 — ExcelSyncEngine.cs V1.0 ✅ build success + BUG-E6 fixed (RevitView alias)*
