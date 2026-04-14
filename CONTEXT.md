# ARCTOOL — AI SESSION CONTEXT
> Paste file này vào ĐẦU mỗi session chat mới với AI.
> Cập nhật sau mỗi session làm việc.
> Last updated: 2026-04-14

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
ArcTool/
├── ArcTool.slnx
├── CONTEXT.md                          ← file này
├── ArcTool.Core/
│   ├── App.cs                          ← Ribbon UI, entry point
│   ├── ArcTool.Core.csproj
│   ├── Commands/
│   │   ├── CreateVoidFromLinkCommand.cs
│   │   ├── MultiCutCommand.cs
│   │   ├── ArrangeDimensionCommand.cs
│   │   └── FilterManagerCommand.cs
│   ├── Services/
│   │   └── ExcelInteropService.cs
│   ├── UI/
│   │   ├── FilterWindow.xaml
│   │   └── FilterWindow.xaml.cs
│   ├── Utilities/
│   │   └── SelectionFilters.cs
│   └── Properties/
│       ├── Resources.resx
│       └── Resources.Designer.cs
└── ArcTool.TestConsole/
    └── TestConsole/
        └── TestConsole/
            ├── Program.cs              ← Test Excel export độc lập
            └── ArcTool.TestConsole.csproj
```

---

## 3. TÍNH NĂNG ĐÃ HOÀN THIỆN

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

### E. Excel Export Engine — `ExcelInteropService.cs` (V5)
- Đọc file Excel (hidden mode), export Print Area hoặc UsedRange thành PNG
- Scale factor: 35x (chú ý: comment trong code ghi sai là "50x" — đã biết)
- Dùng COM Interop, có `IDisposable` đúng chuẩn
- **TODO:** Viết logic Import ảnh vào Revit (ImageImportOptions, ImageInstance.Create)

### F. Filter Manager — `FilterManagerCommand.cs` + `FilterWindow.xaml`
- UI WPF modeless đã xong (FilterWindow với 2 DataGrid: Filters + Views)
- Command skeleton đã xong, dùng `Idling` event để real-time update
- **TODO:** Implement logic Copy/Paste Filter thực sự bằng `ParameterFilterElement` API

---

## 4. BUG ĐÃ PHÁT HIỆN — CHƯA FIX

> Cập nhật trạng thái: [ ] Chưa fix / [x] Đã fix

### 🔴 BUG NGHIÊM TRỌNG

- [ ] **MultiCutCommand** — `(int)elem.Category.Id.Value` gây Integer Overflow
  - Revit 2026: `ElementId.Value` trả về `long`, không phải `int`
  - Fix: `elem.Category.Id.Value == (long)BuiltInCategory.OST_Walls`

- [ ] **CreateVoidFromLinkCommand** — `GetParamValue` chỉ tìm trên `Symbol`, bỏ sót `Instance`
  - Hầu hết custom family dầm lưu Width/Height ở Instance → `beamWidth = 0` → Void không được tạo
  - Fix: tìm Instance trước, fallback về Symbol

- [ ] **CreateVoidFromLinkCommand** — `SetParam(voidInst, "Height", -beamHeight)` gán giá trị âm
  - Revit có thể từ chối commit. Cần thiết kế lại: dùng Mirror hoặc đổi hướng Family

- [ ] **FilterManagerCommand** — `_lastUpdate` là instance field, reset mỗi lần chạy lệnh
  - Fix: đổi thành `static DateTime _lastUpdate`

### 🟠 RỦI RO / CẦN CẢI THIỆN

- [ ] **ArrangeDimensionCommand** — baseline cập nhật dù Dim lỗi (Line == null)
  - Fix: `bool moved = MoveDimension(...); if (moved) baselineDim = nextDim;`

- [ ] **ArrangeDimensionCommand** — không kiểm tra `activeView.Scale == 0` (3D view)

- [ ] **CreateVoidFromLinkCommand** — `doc.Regenerate()` thừa trước `t.Commit()`
  - Revit tự Regenerate khi Commit. Xóa dòng này để tránh lag.

- [ ] **FilterManagerCommand** — `Idling` event là anti-pattern cho model lớn
  - Nâng cấp lên `IExternalEventHandler` + `ExternalEvent.Raise()`

- [ ] **ExcelInteropService** — `finally { obj = null; }` trong `ReleaseObject` vô nghĩa
  - Chỉ null local variable, không null biến caller. Xóa dòng đó.

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

---

## 6. QUY TẮC LẬP TRÌNH (BẮT BUỘC TUÂN THEO)

```csharp
// 1. Transaction attribute bắt buộc
[Transaction(TransactionMode.Manual)]

// 2. Mọi logic chính phải có try-catch
try { ... }
catch (Autodesk.Revit.Exceptions.OperationCanceledException) { return Result.Cancelled; }
catch (Exception ex) { message = ex.Message; return Result.Failed; }

// 3. Thông báo kết quả
TaskDialog.Show("ArcTool", "...");          // Kết quả cho user
uidoc.Application.Application.WriteJournalComment("...", true); // Tiến độ vòng lặp

// 4. Revit 2026: dùng long, không dùng int cho ElementId
elem.Category.Id.Value == (long)BuiltInCategory.OST_Walls  // ĐÚNG
(int)elem.Category.Id.Value == (int)BuiltInCategory.OST_Walls  // SAI

// 5. Đơn vị: Revit dùng feet nội bộ
// 1 foot = 304.8mm. Convert: UnitUtils.ConvertToInternalUnits(mm, UnitTypeId.Millimeters)
```

---

## 7. ROADMAP

### Giai đoạn 1 — Trả nợ kỹ thuật (ƯU TIÊN CAO)
- [ ] Fix bug `long` cast trong `MultiCutCommand`
- [ ] Fix `GetParamValue` tìm Instance trong `CreateVoidFromLinkCommand`
- [ ] Fix `_lastUpdate` thành `static` trong `FilterManagerCommand`
- [ ] Fix baseline update logic trong `ArrangeDimensionCommand`
- [ ] Xóa `doc.Regenerate()` thừa trong `CreateVoidFromLinkCommand`

### Giai đoạn 2 — Filter Manager
- [ ] Implement logic Copy Filter: đọc `ParameterFilterElement` từ View nguồn
- [ ] Implement logic Paste Filter: `view.AddFilter()`, set Visibility/Override
- [ ] Refactor `Idling` → `ExternalEvent` pattern
- [ ] Hoàn thiện MVVM binding cho `FilterWindow`

### Giai đoạn 3 — Excel to Revit Image
- [ ] Viết `ImageImportCommand.cs`
- [ ] Keyword API: `ImageImportOptions`, `ImageInstance.Create()`
- [ ] Xử lý đặt ảnh vào Sheet/View đúng tọa độ

### Giai đoạn 4 — Quick Dim (R&D)
- [ ] Nghiên cứu trích xuất `ReferenceArray` từ Face/Edge của Wall, Column, Beam
- [ ] Revit Dim qua Reference (khác AutoCAD qua điểm tọa độ)
- [ ] Xây dựng hàm `GetDimensionableReferences(Element)` trước khi gọi `NewDimension()`

---

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

> Tra cứu API bắt buộc tại: https://www.revitapidocs.com/2026/

---

## 9. CÁCH DÙNG FILE NÀY

**Đầu mỗi session mới:**
1. Mở file này, copy toàn bộ nội dung
2. Paste vào đầu chat với AI
3. Ghi rõ: "Session hôm nay tôi muốn làm: [nhiệm vụ cụ thể]"

**Cuối mỗi session:**
1. Cập nhật trạng thái bug ([ ] → [x])
2. Thêm quyết định kỹ thuật mới vào Mục 5
3. Commit file lên GitHub cùng với code

---
*ArcTool © 2026 — Internal development context file*
