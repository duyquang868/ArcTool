# ARCTOOL — AI SESSION CONTEXT
> Paste file này vào ĐẦU mỗi session chat mới với AI.
> Cập nhật sau mỗi session làm việc.
> Last updated: 2026-04-28 — Session 6.1: Phase 1A complete — Models/ExcelMapping.cs ✅ build success

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
│   │   └── ExcelInteropService.cs      ← V5.2 stable
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

### E. Excel Export Engine — `ExcelInteropService.cs` (V5.2 — STABLE)
- Đọc file Excel (hidden mode), export Print Area hoặc UsedRange thành PNG
- Scale factor: 35x cố định
- Có IDisposable, COM release đúng thứ tự child → parent
- Có GetActiveSheetName() để lấy tên sheet đang active

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
| JSON lưu cạnh file .rvt | Setting đi theo project folder | Mất nếu copy .rvt sang thư mục khác mà không copy JSON |
| Legend View: Duplicate thay vì Create | Revit API 2026 không có method tạo Legend mới | User phải tạo thủ công 1 Legend View rỗng làm template lần đầu |
| Enum prefix `Excel` (ExcelViewType, ExcelRegionType) | Tránh `CS0104` collision với `Autodesk.Revit.DB.ViewType` khi import cả hai namespace | Tên dài hơn spec gốc — đây là quyết định bắt buộc, không phải tuỳ chọn |

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
      "Id": "guid-string",
      "ViewName": "BudgetOverview",
      "AutoSync": true,
      "LastModified": "2024-04-22T14:53:00",
      "WorkSheet": "Budget Overview",
      "Region": "ChartTest",
      "RegionType": "NamedRange",
      "ViewType": "LegendView",
      "FilePath": "C:\\Project\\Chart-Sample.xlsx",
      "ImageInstanceId": 12345,
      "StoredWidth": 2.5,
      "StoredHeight": 1.8
    }
  ]
}
```

**Giải thích từng field:**

| Field | Type | Mô tả |
|---|---|---|
| `Id` | GUID string | Unique identifier của mapping |
| `ViewName` | string | Tên View Revit = SheetName, hoặc "SheetName_RegionName" nếu là Named Range |
| `AutoSync` | bool | true = tự động update khi dialog mở + file đã thay đổi |
| `LastModified` | DateTime | Timestamp lần update thành công cuối cùng (thủ công hoặc auto) |
| `WorkSheet` | string | Tên sheet trong file Excel |
| `Region` | string? | null = dùng Print Area / UsedRange; tên = Named Range cụ thể |
| `RegionType` | enum | `"NamedRange"` / `"PrintArea"` / `"UsedRange"` |
| `ViewType` | enum | `"LegendView"` / `"DraftingView"` |
| `FilePath` | string | Đường dẫn tuyệt đối tới file Excel |
| `ImageInstanceId` | long | `ElementId.Value` của ImageInstance trong Revit |
| `StoredWidth` | double | Width (feet) của ImageInstance — đọc từ Revit, không phải % |
| `StoredHeight` | double | Height (feet) của ImageInstance — đọc từ Revit |

**Logic Region:**
- `RegionType = "NamedRange"`: gọi `worksheet.Names` để lấy range
- `RegionType = "PrintArea"`: dùng `worksheet.PageSetup.PrintArea`
- `RegionType = "UsedRange"`: fallback — `worksheet.UsedRange`
- Ưu tiên khi export: NamedRange → PrintArea → UsedRange

**Logic ViewName:**
- Dùng Print Area / UsedRange → `ViewName = SheetName`
- Dùng Named Range "ChartTest" trên sheet "Budget Overview" → `ViewName = "Budget Overview_ChartTest"`

---

### 6.3 UI Spec — Bảng chính (WPF DataGrid)

**Columns theo thứ tự:**

| # | Column | Control | Ghi chú |
|---|---|---|---|
| 1 | Select | CheckBox | Chọn nhiều dòng để thực hiện batch action |
| 2 | Status Dot | Ellipse fill | Xanh = synced, Đỏ = file Excel mới hơn LastModified |
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
  ├─ 1. Load JSON (ArcTool_ExcelSync.json cạnh .rvt)
  │      → Deserialize thành List<ExcelMapping>
  │
  ├─ 2. Với mỗi mapping:
  │      ├─ Check File.Exists(FilePath)
  │      │    └─ Không tồn tại → Status = FileNotFound (icon cảnh báo khác)
  │      ├─ So sánh: File.GetLastWriteTime(FilePath) > LastModified
  │      │    └─ Nếu mới hơn → HasChanges = true → Status Dot đỏ
  │      └─ Nếu AutoSync = true && HasChanges = true
  │           → ExecuteUpdate(mapping, doc) tự động
  │
  └─ 3. Show dialog với bảng đã populate
```

---

### 6.5 Luồng Khi User Nhấn `+` (Thêm Dòng Mới)

```
User nhấn "+"
  │
  ├─ Tạo dòng mới với giá trị mặc định:
  │    AutoSync = false
  │    ViewType = "DraftingView" (mặc định)
  │    Region = null (UsedRange)
  │
  ├─ FilePath column: user click Browse button
  │    └─ OpenFileDialog → chọn .xlsx / .xls
  │         └─ Sau khi chọn:
  │              ├─ ExcelInteropService.OpenFile() → đọc sheet names
  │              ├─ WorkSheet dropdown: populate sheet names
  │              └─ ExcelInteropService.Dispose() ngay sau khi đọc xong
  │
  ├─ User chọn WorkSheet
  │    └─ Region dropdown: populate (Print Areas + Named Ranges của sheet đó)
  │
  ├─ ViewName: tự điền = SheetName (hoặc SheetName_RegionName nếu chọn Named Range)
  │    → ViewName tự cập nhật khi WorkSheet hoặc Region thay đổi
  │
  └─ User chọn ViewType → OK
```

---

### 6.6 Luồng Khi Nhấn Update (Per Row hoặc Update All)

```
ExecuteUpdate(ExcelMapping mapping, Document doc)
  │
  ├─ 1. Export Excel → Temp PNG
  │      ExcelInteropService.OpenFile(mapping.FilePath)
  │      ExcelInteropService.ExportRegion(mapping.WorkSheet, mapping.Region, tempPng)
  │      ExcelInteropService.Dispose()
  │
  ├─ 2. Đọc StoredWidth/StoredHeight TRƯỚC KHI xóa ảnh cũ
  │      var existingInst = doc.GetElement(mapping.ImageInstanceId) as ImageInstance
  │      if (existingInst != null && existingInst.IsValidObject)
  │      {
  │          mapping.StoredWidth  = existingInst.Width   ← đọc kích thước thực
  │          mapping.StoredHeight = existingInst.Height    (phản ánh resize của user)
  │      }
  │      // Nếu không tìm thấy instance → dùng giá trị StoredWidth/Height từ JSON
  │
  ├─ 3. Transaction("ArcTool: Refresh Excel Image")
  │      ├─ GetOrCreateView(mapping.ViewName, mapping.ViewType) → targetView
  │      │    ├─ Nếu view đã tồn tại → ghi đè (xóa ImageInstance cũ trong view đó)
  │      │    └─ Nếu chưa có → tạo mới (ViewDrafting.Create hoặc View.CreateLegend)
  │      ├─ Nếu existingInst valid → doc.Delete(existingInst.Id)
  │      ├─ ImageType.Create(doc, tempPng) → imageType
  │      ├─ ImageInstance.Create(doc, targetView, ...) → newInst
  │      ├─ newInst.Width  = mapping.StoredWidth   ← áp lại kích thước đã lưu
  │      ├─ newInst.Height = mapping.StoredHeight
  │      └─ Commit()
  │
  ├─ 4. Cập nhật mapping trong JSON:
  │      mapping.LastModified   = DateTime.Now
  │      mapping.ImageInstanceId = newInst.Id.Value
  │      mapping.StoredWidth    = newInst.Width
  │      mapping.StoredHeight   = newInst.Height
  │      SaveJson()
  │
  ├─ 5. Cập nhật UI: Status Dot → xanh
  │
  └─ 6. TryDeleteFile(tempPng)
```

**Lưu ý quan trọng — Smart Scale:**
- Lần import đầu tiên: `StoredWidth` và `StoredHeight` = giá trị mặc định của Revit sau khi create (100% scale)
- Giá trị này KHÔNG hiển thị cho user và KHÔNG có dialog nhập %
- Sau khi user kéo resize ảnh trực tiếp trong Revit → lần Update tiếp theo sẽ đọc Width/Height mới → lưu → áp lại

---

### 6.7 Xử Lý File Excel Không Tìm Thấy

```
Khi dialog mở + File.Exists(FilePath) == false:
  ├─ Status Dot = màu vàng (warning, khác với đỏ = has changes)
  ├─ Nút Update = disabled
  ├─ Tooltip trên dòng: "File not found. Click to relocate."
  └─ User click vào icon warning:
       └─ OpenFileDialog → chọn file mới
            ├─ mapping.FilePath = newPath
            ├─ SaveJson()
            └─ Re-check status (compare timestamp)
```

---

### 6.8 ExcelInteropService — Mở Rộng Cần Thêm

Service hiện tại (V5.2) chỉ có `ExportPrintAreaAsHighResImage()`. V3.0 cần thêm:

```csharp
// Lấy tất cả sheet names trong file
public List<string> GetSheetNames()

// Lấy tất cả Named Ranges trong một sheet cụ thể
public List<string> GetNamedRanges(string sheetName)

// Export theo sheet + region cụ thể (thay thế ExportPrintAreaAsHighResImage)
public bool ExportRegion(string sheetName, string regionName, string outputPath)
// regionName = null → Print Area → UsedRange (fallback tự động)
```

---

### 6.9 Tạo View Revit — API Notes (ĐÃ VERIFY)

#### Drafting View — ✅ API đầy đủ, stable

```csharp
// Tạo Drafting View mới hoàn toàn — không cần template
ViewFamilyType draftingType = new FilteredElementCollector(doc)
    .OfClass(typeof(ViewFamilyType))
    .Cast<ViewFamilyType>()
    .First(t => t.ViewFamily == ViewFamily.Drafting);

ViewDrafting view = ViewDrafting.Create(doc, draftingType.Id);
view.Name = viewName;
```

#### Legend View — ⚠️ KHÔNG CÓ API TẠO MỚI — Dùng Workaround Duplicate

**Thực trạng đã xác nhận:** Revit API đến phiên bản 2026 **không cung cấp method tạo Legend View mới từ đầu**. `ViewFamily.Legend` enum chỉ dùng để *đọc/lọc*, không có `Create()` tương ứng. Đây là giới hạn lâu năm của Autodesk, cộng đồng đã request nhiều lần nhưng chưa được giải quyết.

**Quyết định: Phương án B — Duplicate từ Legend Template**

Yêu cầu bắt buộc: Project Revit **phải có sẵn ít nhất 1 Legend View rỗng** đóng vai trò template. ArcTool sẽ duplicate cái đó và đặt tên theo sheet.

```csharp
/// <summary>
/// Tìm Legend View template (rỗng, dùng để duplicate).
/// Convention: tên phải là "ArcTool_LegendTemplate" hoặc bất kỳ Legend View nào trong project.
/// Phải gọi trong Transaction đang active.
/// </summary>
private View GetOrCreateLegendView(Document doc, string viewName)
{
    // Bước 1: Kiểm tra Legend View với tên đích đã tồn tại chưa
    var existing = new FilteredElementCollector(doc)
        .OfClass(typeof(View))
        .Cast<View>()
        .FirstOrDefault(v => v.ViewType == ViewType.Legend
                          && string.Equals(v.Name, viewName, StringComparison.OrdinalIgnoreCase));
    if (existing != null) return existing; // Ghi đè: dùng view cũ (xóa ImageInstance cũ bên trong)

    // Bước 2: Tìm Legend View template để duplicate
    View legendTemplate = new FilteredElementCollector(doc)
        .OfClass(typeof(View))
        .Cast<View>()
        .FirstOrDefault(v => v.ViewType == ViewType.Legend && !v.IsTemplate);

    if (legendTemplate == null)
    {
        // Không có Legend View nào trong project → báo lỗi rõ ràng cho user
        throw new InvalidOperationException(
            "Không tìm thấy Legend View nào trong project.\n\n" +
            "Để dùng tính năng này, hãy tạo thủ công 1 Legend View rỗng trong Revit " +
            "(View tab → Legend), sau đó chạy lại lệnh.");
    }

    // Bước 3: Duplicate với detailing (giữ lại các element annotation nếu có)
    ElementId newViewId = legendTemplate.Duplicate(ViewDuplicateOption.WithDetailing);
    View newLegendView = doc.GetElement(newViewId) as View;

    // Đổi tên theo sheet Excel
    try   { newLegendView.Name = viewName; }
    catch { newLegendView.Name = $"{viewName}_{DateTime.Now:HHmmss}"; } // fallback nếu tên trùng

    return newLegendView;
}
```

**UX Flow khi không có Legend Template:**

```
User chọn ViewType = "Legend View" cho một mapping
  └─ Khi nhấn Update:
       └─ GetOrCreateLegendView() throw InvalidOperationException
            └─ Tool hiển thị dialog:
                 "Không tìm thấy Legend View nào trong project.
                  Hãy tạo 1 Legend View rỗng trong Revit (View tab → Legend),
                  sau đó nhấn Update lại."
```

**Trade-off đã chấp nhận:**
- User phải tạo thủ công 1 Legend View rỗng lần đầu tiên
- Nếu project có nhiều Legend View, tool sẽ dùng cái tìm thấy đầu tiên làm template
- Nếu Legend template có annotation elements → chúng sẽ bị copy sang view mới (WithDetailing)
  → Khuyến nghị: đặt tên template rõ ràng, ví dụ "ArcTool_Template", để dễ quản lý

---

### 6.10 JSON Service — ArcToolSettingsService

File mới cần tạo: `Services/ArcToolSettingsService.cs`

```csharp
public class ArcToolSettingsService
{
    // Path: [rvt_folder]\ArcTool_ExcelSync.json
    public static string GetSettingsPath(Document doc)

    public static List<ExcelMapping> LoadMappings(Document doc)
    public static void SaveMappings(Document doc, List<ExcelMapping> mappings)
}
```

---

### 6.11 Checklist Implementation V3.0

```
Phase 1 — Models & Services (không phụ thuộc UI)
  [x] Tạo Models/ExcelMapping.cs ✅ Session 6.1 — build success
        NOTE: Enum đổi tên ExcelViewType / ExcelRegionType (tránh collision DB.ViewType)
        NOTE: Region = null (không phải "") cho "chưa chọn Named Range"
        NOTE: LastModified default = DateTime.MinValue (file luôn "changed" lần đầu — đúng ý)
  [ ] Tạo Services/ArcToolSettingsService.cs (Load/Save JSON)
  [ ] Mở rộng Services/ExcelInteropService.cs → V5.3:
        [ ] GetSheetNames()
        [ ] GetNamedRanges(sheetName)
        [ ] ExportRegion(sheetName, regionName, outputPath)
  [x] Verify Legend View creation API tại revitapidocs.com/2026 ✅ (không có Create())

Phase 2 — Logic Layer (không phụ thuộc UI)
  [ ] Viết ExcelSyncEngine.cs:
        [ ] CheckForChanges(List<ExcelMapping>, Document)
        [ ] ExecuteUpdate(ExcelMapping, Document)
        [ ] GetOrCreateView(viewName, viewType, Document) — trong Transaction

Phase 3 — UI
  [ ] Thiết kế ExcelToRevitWindow.xaml (WPF DataGrid theo UI Spec 6.3)
  [ ] ExcelToRevitWindow.xaml.cs:
        [ ] Load data vào DataGrid
        [ ] Handle "+" / "-" buttons
        [ ] Handle Update per-row button
        [ ] Handle "Update All" button
        [ ] Handle Browse file button
        [ ] Handle File Not Found warning

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

### Giai đoạn 3 — Excel to Revit V3.0 🔧 TIẾP THEO
- Xem chi tiết tại Section 6

### Giai đoạn 4 — Quick Dim (R&D) 📋 TƯƠNG LAI
- Nghiên cứu ReferenceArray extraction từ Wall, Column, Beam
- Revit Dim qua Reference, khác AutoCAD qua tọa độ điểm

---

## 8. QUY TẮC LẬP TRÌNH (BẮT BUỘC)

```csharp
// 1. Transaction attribute
[Transaction(TransactionMode.Manual)]

// 2. Namespace alias tránh conflict
using RevitTaskDialog = Autodesk.Revit.UI.TaskDialog;

// 3. ElementId: luôn dùng long
elem.Category.Id.Value == (long)BuiltInCategory.OST_Walls  // ĐÚNG
(int)elem.Category.Id.Value                                 // SAI

// 4. COM release: child → parent, null field gốc, KHÔNG release sau Delete()
if (_activeSheet != null) { ReleaseObject(_activeSheet); _activeSheet = null; }

// 5. Quick filter trước slow filter
new FilteredElementCollector(doc)
    .OfClass(typeof(Wall))          // quick — index-based
    .OfCategory(...)                // quick
    .Where(w => ...)                // slow — sau cùng

// 6. Smart Scale pattern — đọc kích thước TRƯỚC KHI xóa
if (existingInst?.IsValidObject == true)
{
    storedWidth  = existingInst.Width;
    storedHeight = existingInst.Height;
    doc.Delete(existingInst.Id);
    // KHÔNG ReleaseComObject — Revit managed object
}

// 7. JSON settings lưu cạnh .rvt
string dir = Path.GetDirectoryName(doc.PathName);
string jsonPath = Path.Combine(dir, "ArcTool_ExcelSync.json");

// 8. Verify Legend View API trước khi code
// → site:revitapidocs.com/2026 Legend View Create

// 9. Enum trong Models PHẢI có prefix để tránh collision với Revit API
// Sai:  public enum ViewType   → CS0104 khi file import Autodesk.Revit.DB
// Đúng: public enum ExcelViewType
// Sai:  public enum RegionType → tiềm năng collision tương lai
// Đúng: public enum ExcelRegionType
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
| `view.Duplicate(ViewDuplicateOption.WithDetailing)` | Tạo Legend View mới bằng cách duplicate template — workaround bắt buộc vì không có Create() |
| `ViewType.Legend` | Dùng để filter/đọc Legend View hiện có, không có Create() tương ứng |
| `ParameterFilterElement` | API cho Filter Manager |
| `View.AddFilter()` / `View.GetFilters()` | Copy/Paste filter |
| `Marshal.ReleaseComObject()` | COM release — child trước parent |
| `File.GetLastWriteTime(path)` | Check timestamp Excel file |

> **Tra cứu API bắt buộc:** https://www.revitapidocs.com/2026/

---

*ArcTool © 2026 — Internal development documentation*
*Session 6.1: Phase 1A — ExcelMapping.cs ✅ build success + enum naming decision locked*
