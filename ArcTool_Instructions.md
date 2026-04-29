## ĐỐI TÁC CỦA BẠN

Tôi là kiến trúc sư với **20 năm kinh nghiệm thực chiến** trong ngành thiết kế xây dựng. Tôi hiểu sâu về logic nghiệp vụ, quy trình thiết kế BIM và nhu cầu thực tế của công việc kiến trúc. Tôi **không phải coder chuyên nghiệp** — tôi tiếp cận lập trình như một *problem-solver*: tôi biết mình muốn gì, tôi hiểu tại sao cần làm, nhưng tôi cần bạn là người biến ý tưởng đó thành code cấp enterprise.

---

## DỰ ÁN: ARCTOOL PLUGIN

| Mục | Chi tiết |
|---|---|
| Tên dự án | ArcTool |
| Namespace | `ArcTool.Core` |
| Nền tảng | Autodesk Revit 2026 (API 2026) |
| Ngôn ngữ | C# (.NET 8.0) |
| IDE | Visual Studio Enterprise 2026 |
| UI | WPF (modeless) + WinForms (dialog) |

---

## TÍNH NĂNG HIỆN TẠI

### A. Ribbon UI — `App.cs` (V5.1 — STABLE)
- Tab `ArcTool` với 3 panel: Void Tools, Annotation Tools, Excel Tools
- SplitButton "Void Manager": Create Void (main) + Multi-Cut (dropdown)
- Helper `ConvertToImageSource(Bitmap)` chuyển Resource sang WPF ImageSource

### B. Create Void — `CreateVoidFromLinkCommand.cs` (V4.0 — STABLE)
- Tự động tạo Generic Model Void tại TẤT CẢ dầm trong file Link
- Face-Based placement, midpoint của LocationCurve
- ⚠️ Biết rủi ro: Unhosted nếu Link reload

### C. Multi-Cut — `MultiCutCommand.cs` (V2.0 — STABLE)
- Cắt Tường + Cột bằng Void đã tạo
- Broad Phase: BoundingBoxIntersectsFilter
- ⏳ Thiếu: Narrow Phase Solid Intersection

### D. Arrange Dimensions — `ArrangeDimensionCommand.cs` (V1.0 — STABLE)
- Tịnh tiến Dim cách đều theo Snap Distance × View Scale
- TransactionGroup → 1 lần Undo

### E. Excel Export Engine — `ExcelInteropService.cs` (V5.2 — STABLE)
- Hidden Excel, export Print Area / UsedRange → PNG 35x scale
- COM release đúng thứ tự child → parent
- Method hiện có: `OpenFile()`, `GetActiveSheetName()`, `ExportPrintAreaAsHighResImage()`
- ⏳ Cần thêm (V5.3): `GetSheetNames()`, `GetNamedRanges(sheetName)`, `ExportRegion(sheetName, regionName, outputPath)`

### F. Excel to Revit — `ExcelToRevitCommand.cs` (V1.0 — STABLE, chờ V3.0)
- Pipeline đơn giản: chọn Excel → export PNG → ImageType.Create() → ImageInstance
- ⏳ Cần nâng cấp lên V3.0: DataGrid UI, change detection, smart scale, auto-create view

### G. Filter Manager — `FilterManagerCommand.cs` + `FilterWindow.xaml` (SKELETON)
- UI WPF modeless đã xong, dùng Idling event
- ⏳ Cần: implement logic Copy/Paste Filter

### H. ExcelMapping Model — `Models/ExcelMapping.cs` (V1.0 ✅ STABLE — Session 6.1)
- POCO class, JsonSerializable — không phụ thuộc Revit API, compile độc lập
- Namespace: `ArcTool.Core.Models`
- Enum `ExcelRegionType`: NamedRange / PrintArea / UsedRange
- Enum `ExcelViewType`: DraftingView / LegendView
- ⚠️ Prefix `Excel` bắt buộc — tránh CS0104 collision với `Autodesk.Revit.DB.ViewType`
- `[JsonIgnore]` helpers: `IsFirstImport`, `HasStoredDimensions`, `BuildViewName()`
- `Region = null` (không phải `""`) = "chưa chọn Named Range"
- Sentinel: `ImageInstanceId = 0`, `StoredWidth/Height = 0.0`, `LastModified = DateTime.MinValue`

### I. Settings Service — `Services/ArcToolSettingsService.cs` (V1.0 ✅ STABLE — Session 6.2)
- Static class, chỉ phụ thuộc `Document` (đọc `doc.PathName`) và `ArcTool.Core.Models`
- **Atomic write**: ghi `.tmp` cùng thư mục → `File.Replace()` / `File.Move()` — không bao giờ corrupt JSON nếu crash
- **`JsonStringEnumConverter`**: enum serialize thành string (`"DraftingView"`) trong JSON
- **`JsonSerializerOptions`** cached `static readonly` — không allocate mới mỗi call
- **`LoadMappings()`**: trả về `List<>` rỗng nếu file không tồn tại hoặc corrupt — không crash
- **`SaveMappings()`**: throw `IOException` nếu ghi thất bại — caller phải hiện dialog
- **`GetSettingsPath()`**: throw `InvalidOperationException` khi `doc.PathName` rỗng
- **`FileExists(mapping)`**: guard cho Status Dot vàng
- **`HasFileChanged(mapping)`**: so sánh `File.GetLastWriteTime()` (local) > `mapping.LastModified` (local)
- ⚠️ Luôn dùng `DateTime.Now` khi gán `LastModified` — KHÔNG `DateTime.UtcNow`

---

## EXCEL TO REVIT V3.0 — SPEC TÓM TẮT

> Chi tiết đầy đủ ở CONTEXT.md Section 6. Phần này chỉ tóm tắt để dev nhanh.

### Tính năng 3 nhóm

| # | Nhóm | Mô tả |
|---|---|---|
| T1 | Auto-Create View | Drafting View (API đầy đủ) hoặc Legend View (Duplicate workaround) theo tên sheet |
| T2 | Change Detection | Check timestamp khi dialog mở. AutoSync = tự update. Thủ công = nhấn Update per-row |
| T3 | Smart Scale | Lưu Width/Height (feet) của ImageInstance vào JSON. Áp lại khi refresh |

### File mới cần tạo

| File | Mô tả | Trạng thái |
|---|---|---|
| `Models/ExcelMapping.cs` | POCO: Id, ViewName, AutoSync, LastModified, WorkSheet, Region, **ExcelRegionType**, **ExcelViewType**, FilePath, ImageInstanceId, StoredWidth, StoredHeight | ✅ DONE — Session 6.1 |
| `Services/ArcToolSettingsService.cs` | Load/Save JSON cạnh .rvt — atomic write pattern, FileExists(), HasFileChanged() | ✅ DONE — Session 6.2 |
| `Services/ExcelInteropService.cs` (mở rộng) | Thêm GetSheetNames(), GetNamedRanges(), ExportRegion() → V5.3 | ⏳ Phase 1C |
| `Services/ExcelSyncEngine.cs` | CheckForChanges(), ExecuteUpdate(), GetOrCreateView() | ⏳ Phase 2 |
| `UI/ExcelToRevitWindow.xaml` + `.cs` | WPF DataGrid 10 cột theo UI spec | ⏳ Phase 3 |

### Legend View — Giới hạn API đã verify
Revit API 2026 **không có method tạo Legend View từ đầu**. Workaround: `view.Duplicate(ViewDuplicateOption.WithDetailing)`. Yêu cầu: project phải có sẵn 1 Legend View rỗng tên `ArcTool_LegendTemplate`.

### Columns bảng UI (theo thứ tự)
Select | Status Dot | View Name | Auto Sync | Last Modified | WorkSheet | Region | View Type | File Path | Update button

---

## BUILD & DEPLOYMENT

### Build Commands
```powershell
cd "D:\OneDrive - MSFT\Plugin Revit\ArcTool"
dotnet build -c Debug
dotnet build -c Release
```

### Output Path
```
ArcTool.Core\Bin\x64\Debug\net8.0-windows\
```

### Deployment Checklist
- [ ] Build thành công (0 errors, 0 warnings)
- [ ] Tất cả commands đăng ký trong App.cs
- [ ] CONTEXT.md cập nhật
- [ ] SKILL.md cập nhật
- [ ] ArcTool_Instructions.md cập nhật
- [ ] Git commit & push
- [ ] .addin manifest đúng path

---

## TESTING — QUY TRÌNH THỦ CÔNG

### Test 1: Create Void

**Setup:** File RVT có Structural Link chứa dầm

**Steps:**
1. ArcTool Tab → Void Tools → Create Void
2. Chọn Family Void trong dialog
3. Pick file Link

**Assertions:**
- [ ] Void tạo tại midpoint dầm
- [ ] Height/Width/Length đọc từ dầm parameters
- [ ] Void là Generic Model, Face-Based

---

### Test 2: Multi-Cut

**Setup:** Có Void instances từ Test 1, có Tường + Cột

**Steps:**
1. ArcTool Tab → Void Tools → Multi-Cut
2. Quét chọn Voids
3. Quét chọn Tường + Cột

**Assertions:**
- [ ] Tường bị cắt
- [ ] Cột kết cấu bị cắt
- [ ] Cột kiến trúc bị cắt

---

### Test 3: Arrange Dimensions

**Setup:** View có ít nhất 3 Linear Dimension

**Steps:**
1. ArcTool Tab → Annotation Tools → Arrange Dimensions
2. Pick Dim đầu tiên (baseline)
3. Pick Dim thứ 2, 3 liên tiếp
4. Nhấn ESC để kết thúc

**Assertions:**
- [ ] Dim tịnh tiến đúng khoảng = Snap Distance × View Scale
- [ ] TransactionGroup: toàn bộ là 1 lần Undo

---

### Test 4: Excel to Revit V1.0 (Hiện tại)

**Setup:** File Excel có Print Area, Revit Sheet đang mở

**Steps:**
1. ArcTool Tab → Excel Tools → Excel to Revit
2. Chọn file Excel
3. Nhập Scale % (mặc định 100)

**Assertions:**
- [ ] Excel chạy ẩn, không hiện lên màn hình
- [ ] PNG export tại %TEMP%
- [ ] Ảnh đặt tại tâm View
- [ ] Scale áp đúng
- [ ] Temp PNG xóa sau commit

---

### Test 5: Excel to Revit V3.0 (Sau khi implement)

**Setup:** File Excel có nhiều sheet, một số sheet có Named Ranges. File Revit đã lưu (có PathName).

#### Test 5A — Import lần đầu (Drafting View)
1. Mở dialog Excel to Revit
2. Nhấn `+` → Browse chọn file Excel
3. WorkSheet dropdown: chọn sheet có Named Range
4. Region dropdown: chọn Named Range
5. View Type: Drafting View
6. Nhấn Update

**Assertions:**
- [ ] Drafting View mới tạo, tên = `[SheetName]_[RegionName]`
- [ ] Ảnh đặt tại tâm View
- [ ] Status Dot = xanh
- [ ] JSON lưu đúng: FilePath, WorkSheet, Region, ImageInstanceId, StoredWidth, StoredHeight
- [ ] LastModified = `DateTime.Now` (local time)

#### Test 5B — Import lần đầu (Legend View)
1. Tạo trước: 1 Legend View rỗng tên `ArcTool_LegendTemplate` trong Revit
2. Mở dialog, `+` → chọn Excel, sheet, region
3. View Type: Legend View
4. Nhấn Update

**Assertions:**
- [ ] Legend View mới tạo bằng cách Duplicate template
- [ ] Tên đúng = `[SheetName]` hoặc `[SheetName]_[RegionName]`
- [ ] Nếu không có Legend View nào trong project → hiện error dialog rõ ràng

#### Test 5C — Smart Scale
1. Sau Test 5A: kéo resize ảnh trực tiếp trong Revit View
2. Sửa nội dung file Excel và lưu
3. Mở lại dialog Excel to Revit
4. Status Dot = đỏ (file mới hơn LastModified)
5. Nhấn Update

**Assertions:**
- [ ] Ảnh mới có cùng Width/Height với ảnh đã resize (không reset về 100%)
- [ ] JSON cập nhật: StoredWidth/Height = kích thước mới
- [ ] Status Dot → xanh sau update

#### Test 5D — AutoSync
1. Check AutoSync = true cho 1 mapping
2. Sửa file Excel và lưu
3. Đóng và mở lại dialog

**Assertions:**
- [ ] Khi dialog mở: mapping có AutoSync=true tự động update (không cần nhấn nút)
- [ ] Nút Update per-row bị disabled khi AutoSync=true
- [ ] LastModified cập nhật = `DateTime.Now`

#### Test 5E — File Not Found
1. Di chuyển file Excel sang thư mục khác
2. Mở dialog

**Assertions:**
- [ ] Status Dot = màu vàng (`ArcToolSettingsService.FileExists()` trả về false)
- [ ] Nút Update disabled
- [ ] Click icon warning → OpenFileDialog cho phép chọn lại đường dẫn

#### Test 5F — ArcToolSettingsService (Unit-level)
1. Chạy lệnh khi file `.rvt` chưa được lưu (doc.PathName rỗng)

**Assertions:**
- [ ] `GetSettingsPath()` throw `InvalidOperationException` — dialog yêu cầu lưu file
- [ ] Không crash NullReferenceException

2. Corrupt thủ công file `ArcTool_ExcelSync.json` (xóa một nửa nội dung)
3. Mở lại dialog

**Assertions:**
- [ ] Tool vẫn mở được, DataGrid rỗng (không crash)
- [ ] File `.corrupt_[timestamp]` xuất hiện cạnh JSON

---

## RIBBON UI LAYOUT

```
ArcTool Tab (3 Panels)
├── Void Tools Panel
│   └── SplitButton "Void Manager"
│       ├── Create Void (main, synchronized)
│       ├── [Separator]
│       └── Multi-Cut (dropdown)
├── Annotation Tools Panel
│   └── Arrange Dimensions (push button)
└── Excel Tools Panel
    └── Excel to Revit (push button) → mở ExcelToRevitWindow (V3.0)
```

---

## TROUBLESHOOTING

| Error / Symptom | Nguyên nhân | Giải pháp |
|---|---|---|
| CS0104: TaskDialog ambiguous | Import WinForms + Revit.UI | Dùng alias: `using RevitTaskDialog = Autodesk.Revit.UI.TaskDialog;` |
| CS0104: ViewType ambiguous | Import `ArcTool.Core.Models` + `Autodesk.Revit.DB` trong cùng file | Enum trong Models phải có prefix: `ExcelViewType`, `ExcelRegionType` |
| Excel process vẫn còn trong Task Manager | COM object chưa release đúng | Kiểm tra Dispose() có null đủ `_activeSheet`, `_workbook`, `_excelApp` không |
| Ảnh quá nhỏ sau import | ScaleFactor sai | Đọc Width/Height từ `ImageInstance` sau Create, nhân với factor |
| Legend View không tạo được | Không có Legend View template | Tạo thủ công 1 Legend View rỗng tên `ArcTool_LegendTemplate` trong Revit |
| JSON không lưu được | File .rvt chưa được lưu lần nào | `doc.PathName` rỗng → `GetSettingsPath()` throw → hiện dialog yêu cầu lưu file |
| Status Dot không đổi sang đỏ | Timestamp compare sai hoặc mix UTC/local | `HasFileChanged()` dùng local time; `LastModified` phải gán `DateTime.Now` |
| Smart Scale reset về mặc định sau Update | Quên đọc Width/Height trước Delete | Đọc `existingInst.Width/Height` TRƯỚC `doc.Delete(existingInst.Id)` |
| JSON bị corrupt sau crash | `File.WriteAllText()` trực tiếp | Luôn dùng `ArcToolSettingsService.SaveMappings()` — atomic write |
| Enum deserialize lỗi từ JSON cũ | JSON cũ lưu enum dạng số (0, 1) | JSON mới dùng string ("DraftingView") — cần migration nếu upgrade |
| Named Ranges không hiện trong dropdown | Sheet chưa được chọn | `GetNamedRanges()` phụ thuộc tên sheet, cần chọn WorkSheet trước |
| Dimension skip một số Dim | `Dim.Curve == null` | `ArrangeDimensionCommand` có guard `if (baseLine == null) return false` |
| Integer Overflow trong filter | `(int)Category.Id.Value` | Luôn dùng `(long)BuiltInCategory.OST_Walls` |

---

## COMMON DEVELOPMENT TASKS

### Thêm Command mới vào Ribbon
1. Tạo file `.cs` trong `Commands/` folder
2. Implement `IExternalCommand`, đặt `[Transaction(TransactionMode.Manual)]`
3. Trong `App.cs`: tạo `PushButtonData` trỏ đến class mới
4. Thêm vào panel tương ứng

### Thêm method mới vào ExcelInteropService
1. Mở file Excel trước: `OpenFile(path)` phải được gọi trước
2. Method mới thao tác với `_workbook` hoặc `_activeSheet`
3. Release mọi COM object tạm thời ngay trong method (không để lại)
4. Đặt `using var svc = new ExcelInteropService()` ở caller — đảm bảo Dispose()

### Thêm field mới vào ExcelMapping (JSON)
1. Thêm property vào `Models/ExcelMapping.cs`
2. Nếu field mới không có trong JSON cũ → đặt default value hợp lý (nullable hoặc default)
3. `ArcToolSettingsService` không cần sửa — `JsonSerializer` tự xử lý field mới với default value
4. Cập nhật DataGrid binding trong `ExcelToRevitWindow.xaml`

### Gọi ArcToolSettingsService đúng cách
```csharp
// Load — luôn wrap try-catch vì có thể throw nếu doc.PathName rỗng
try { mappings = ArcToolSettingsService.LoadMappings(doc); }
catch (InvalidOperationException ex) { TaskDialog.Show("ArcTool", ex.Message); return Result.Failed; }

// Save — wrap thêm IOException
try { ArcToolSettingsService.SaveMappings(doc, mappings); }
catch (InvalidOperationException ex) { TaskDialog.Show("ArcTool", ex.Message); return Result.Failed; }
catch (IOException ex) { TaskDialog.Show("ArcTool Error", $"Không thể lưu: {ex.Message}"); return Result.Failed; }

// Check status — không cần try-catch, cả hai handle exception nội bộ
bool exists    = ArcToolSettingsService.FileExists(mapping);
bool hasChange = ArcToolSettingsService.HasFileChanged(mapping);
```

### Debug Excel COM issue
1. Build và chạy trong Revit
2. Mở Task Manager → Processes → tìm `EXCEL.EXE`
3. Nếu còn sau khi lệnh kết thúc → COM object chưa được release đúng
4. Đặt breakpoint tại `Dispose()` → kiểm tra từng field

### Debug JSON settings issue
1. Mở thư mục chứa file `.rvt`
2. Kiểm tra `ArcTool_ExcelSync.json` tồn tại và đúng format
3. Nếu thấy file `.tmp` → lần save trước bị interrupt giữa chừng (vô hại, sẽ bị overwrite lần sau)
4. Nếu thấy file `.corrupt_[timestamp]` → JSON đã bị corrupt, tool đã tự backup và reset về List rỗng

---

## RESOURCES & LINKS

- **Revit API Docs 2026:** https://www.revitapidocs.com/2026/
- **Autodesk Forum:** https://forums.autodesk.com/t5/revit-api/ct-p/area-p127
- **GitHub:** https://github.com/duyquang868/ArcTool
- **Excel COM Interop:** https://docs.microsoft.com/en-us/office/vba/api/overview/

---

*ArcTool Development & Usage Instructions © 2026*
*Last updated: Session 6.2 — ArcToolSettingsService.cs ✅ build success + atomic write pattern locked*
