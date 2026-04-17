## ĐỐI TÁC CỦA BẠN

Tôi là kiến trúc sư với **20 năm kinh nghiệm thực chiến** trong ngành thiết kế xây dựng. Tôi hiểu sâu về logic nghiệp vụ, quy trình thiết kế BIM và nhu cầu thực tế của công việc kiến trúc. Tôi **không phải coder chuyên nghiệp** — tôi tiếp cận lập trình như một *problem-solver*: tôi biết mình muốn gì, tôi hiểu tại sao cần làm, nhưng tôi cần bạn là người biến ý tưởng đó thành code cấp enterprise.

Nhiệm vụ của bạn: kết hợp hoàn hảo với kiến thức chuyên ngành kiến trúc của tôi để tạo ra phần mềm thực sự giải quyết vấn đề thực tế trong nghề. Bạn vừa là "thợ cả" vừa là "người thầy" — tôi học từ bạn, và bạn học bối cảnh nghiệp vụ từ tôi.

---

## DỰ ÁN: ARCTOOL PLUGIN

**Tên dự án:** ArcTool
**Namespace:** ArcTool.Core
**Nền tảng:** Autodesk Revit 2026 (API 2026)
**Ngôn ngữ:** C# (.NET 8.0)
**Công cụ:** Visual Studio Enterprise 2026
**Thư viện UI:** Windows Forms (Form chọn Family) & Revit Ribbon (App.cs)
**Quản lý Resource:** File `Resources.resx` (Access Modifier: Public) chứa Icon

---

## TÍNH NĂNG ĐÃ HOÀN THIỆN

### A. Giao diện Cốt lõi — App.cs (UI V5.1 - Final)
- Sử dụng **SplitButton** (Nút chia đôi) cho nhóm "Void Manager"
- Icon: Load từ `Properties.Resources` qua hàm helper `ConvertToImageSource` (chuyển Bitmap sang ImageSource)
- Đã tích hợp: **Create Void** (Nút chính - Synchronized) và **Multi-Cut** (ngăn cách bằng Separator)
- **Panel 3: Excel Tools** → Button "Excel to Revit" (NEW - SESSION 5) ✅
- Đã sẵn sàng mở rộng cho Filter Manager

### B. Lệnh tạo Void — CreateVoidFromLinkCommand.cs (V4.0 - Auto Generate)
- **Chức năng:** Tự động tạo Void (Generic Model) tại vị trí TẤT CẢ dầm (Structural Framing) trong file Link được chọn — không cần quét chọn PickBox
- **Logic vị trí:** Lấy trung điểm của LocationCurve
- **Logic kích thước:** Width/Length lấy từ tham số dầm Link; Height gán giá trị ÂM (-Height) để đảo chiều Void (Fix lỗi vị trí Bottom)
- ⚠️ **Đang trong diện refactor:** xử lý tọa độ Transform Matrix chưa áp dụng

### C. Lệnh cắt đa năng — MultiCutCommand.cs (V2.0 - Performance Optimized)
- **Chức năng:** Cắt Tường (Walls) và Cột (Columns) bằng các Void đã tạo
- **Quy trình:** Lọc chọn Void (Generic Model) → Lọc chọn vùng đối tượng mục tiêu (CutTargetSelectionFilter)
- **Thuật toán Broad Phase:** Dùng `BoundingBoxIntersectsFilter` kiểm tra va chạm sơ bộ, chống treo máy với dự án lớn
- ⚠️ **Chờ nâng cấp:** Narrow Phase Solid Intersection chưa được thêm vào

### D. Engine xuất ảnh từ Excel — ExcelInteropService.cs (V5.1 - SESSION 5 Fixed)
- **Trạng thái:** Core Engine hoàn thiện (đọc file Excel và xuất thành PNG 300 DPI) ✅
- **SESSION 5 Fixes:** 4 COM object management bugs fixed (BUG-E1 to E4)
  - ✅ BUG-E1: ReleaseObject + null field gốc
  - ✅ BUG-E2: Comment sai "50x" → "35x"
  - ✅ BUG-E3: COM release order (child → parent)
  - ✅ BUG-E4: Added _activeSheet to Dispose()

### D.1 Lệnh Excel to Revit — ExcelToRevitCommand.cs (V1.0 - SESSION 5 NEW) ✨
- **Trạng thái:** Complete unified pipeline (Excel → PNG → ImageType → ImageInstance) ✅
- **Deprecated:** ImageImportCommand.cs (removed SESSION 5)
  - Reason: API assumption sai ("Revit không expose ImageType.Create()")
  - Reality: Revit 2026 **CÓ HỖ TRỢ** ImageType.Create() công khai ✅
- **Pipeline:**
  1. User picks Excel file (OpenFileDialog)
  2. ExcelInteropService exports → Temp PNG at %TEMP% (GUID name)
  3. ImageType.Create() → Creates from PNG (300 DPI, embedded)
  4. Modern dialog → asks scale % (TableLayoutPanel UI)
  5. Calculate View center (BoundingBoxXYZ with 3-tier fallback)
  6. ImageInstance.Create() → places at center (BoxPlacement.Center)
  7. Apply scale factor (width × factor, height × factor)
  8. Cleanup → delete temp PNG after commit
  9. Show success message
- **SESSION 5 Fix:** Ambiguous reference (TaskDialog + TextBox)
  - Added alias: `using RevitTaskDialog = Autodesk.Revit.UI.TaskDialog;`
  - Explicit qualification: `new System.Windows.Forms.TextBox`
- **Supported Views:** Sheet, Drafting, Floor Plan, Ceiling Plan, Section, Elevation, Detail, Legend
- **Unsupported:** 3D view, Walkthrough, Rendering

### E. Nền móng Filter Manager
- **Trạng thái:** Đã khởi tạo `FilterManagerCommand.cs` và UI WPF (`FilterWindow.xaml` / `FilterWindow.xaml.cs`)
- ⏳ **Sẵn sàng cho:** Tích hợp MVVM và thư viện `ParameterFilterElement` của Revit 2026 API

---

## ROADMAP PHÁT TRIỂN

### Giai đoạn 1 — Cải tiến & Trả nợ kỹ thuật (Technical Debt)
> *Nâng cấp chất lượng các lệnh hiện có trước khi xây thêm*

- **CreateVoidFromLinkCommand:** Áp dụng đúng Transformation Matrix của file Link để fix lỗi lệch tọa độ
- **MultiCutCommand:** Thêm Narrow Phase bằng Solid Intersection — thay thế việc chỉ dựa vào BoundingBox

### Giai đoạn 2 — Filter Manager
- Xây dựng UI/UX đầy đủ bằng WPF (kiến trúc MVVM)
- Tích hợp các bộ lọc (Filter) chuẩn của Revit 2026 qua `ParameterFilterElement` API

### Giai đoạn 3 — Excel to Revit Image (COMPLETED - SESSION 5) ✅
- ✅ **ExcelInteropService:** Core engine hoàn thiện + COM fixes
- ✅ **ExcelToRevitCommand:** Unified pipeline (Excel → PNG → ImageType → ImageInstance)
- ✅ **Integration:** Thêm button vào Ribbon (Tab ArcTool → Panel Excel Tools)
- ✅ **Session 5 Fixes:** Ambiguous reference + COM management
- ✅ **Build:** Successful (0 errors, 0 warnings)

### Giai đoạn 4 — Quick Dim (R&D)
- Đang trong giai đoạn nghiên cứu thuật toán
- **Khác biệt cốt lõi so với AutoCAD:** Revit yêu cầu Dimension qua **Reference objects** (mặt phẳng, đường hình học), không phải qua tọa độ điểm
- **Việc cần làm trước:** Xây dựng hàm trích xuất `ReferenceArray` chuẩn xác trước khi gọi `NewDimension`

---

## TESTING & DEBUGGING

### Manual Testing Procedures

#### Test 1: Create Void
1. Open Revit → mở file RVT có Structural Link
2. ArcTool Tab → Void Tools → Create Void
3. Select Family: Dialog chọn Void family
4. Pick Link: Chọn file Link
5. **Expected:** Void instances tạo tại tất cả dầm locations

**Assertions:**
- [ ] Void được tạo tại midpoint của dầm
- [ ] Height/Width/Length lấy từ dầm parameters
- [ ] Void là Generic Model type
- [ ] Void là Face-Based (bám vào Link geometry)

#### Test 2: Multi-Cut
1. Have: Void instances từ Test 1
2. ArcTool Tab → Void Tools → Multi-Cut
3. Pick Void: Chọn 1 Void
4. **Expected:** Walls + Columns bị cắt bởi Void

**Assertions:**
- [ ] Tường (Walls) bị cắt
- [ ] Cột kết cấu (Structural Columns) bị cắt
- [ ] Cột kiến trúc (Architectural Columns) bị cắt

#### Test 3: Arrange Dimensions
1. Open View với linear dimensions
2. ArcTool Tab → Annotation Tools → Arrange Dimensions
3. Pick baseline: Chọn Dim đầu tiên
4. Pick subsequent: Chọn Dim thứ 2, 3, 4... liên tiếp
5. Press ESC: Kết thúc lệnh
6. **Expected:** Tất cả Dim tiếp theo được tịnh tiến cách đều

**Assertions:**
- [ ] Dim tịnh tiến theo snap distance × view scale
- [ ] Aspect ratio giữ nguyên
- [ ] TransactionGroup gộp thành 1 undo

#### Test 4: Excel to Revit (NEW - SESSION 5) ✨
1. **Prepare Excel:** file .xlsx với Print Area định nghĩa
2. **Open Revit Sheet**
3. **ArcTool Tab → Excel Tools → Excel to Revit**
4. **Choose Excel:** Dialog file browser
5. **ExcelInteropService exports:** PNG (hidden, 300 DPI)
6. **Enter Scale %:** Dialog (default 100%)
7. **Expected:** Ảnh import vào Sheet center, scale áp dụng

**Assertions:**
- [ ] Excel hidden mode, không bị lock
- [ ] PNG export thành công tại %TEMP%
- [ ] ImageType.Create() tạo từ PNG
- [ ] ImageInstance place tại View center
- [ ] Scale áp dụng đúng (width × factor)
- [ ] Temp PNG xóa sau commit
- [ ] Error messages chi tiết (Excel open → export → ImageType → ImageInstance)

---

## RIBBON UI LAYOUT

```
ArcTool Tab (3 Panels)
├── Void Tools Panel
│   └── SplitButton "Void Manager"
│       ├── Create Void (main)
│       ├── [Separator]
│       └── Multi-Cut (dropdown)
├── Annotation Tools Panel
│   └── Arrange Dimensions (push button)
└── Excel Tools Panel
    └── Excel to Revit (push button) ✅ NEW - SESSION 5
```

---

## BUILD & DEPLOYMENT

### Visual Studio Setup
1. Open ArcTool.sln
2. Project: ArcTool.Core, Target: .NET 8.0
3. Build → Build Solution (Ctrl+Shift+B)
4. Output: `ArcTool.Core\Bin\x64\Debug\net8.0-windows\`

### Build Commands
```powershell
cd D:\OneDrive - MSFT\Plugin Revit\ArcTool
dotnet build -c Debug
dotnet build -c Release
```

### Deployment Checklist
- [ ] Build successful (0 errors, 0 warnings)
- [ ] All commands registered in App.cs
- [ ] CONTEXT.md updated
- [ ] SKILL.md updated  
- [ ] ArcTool_Instructions.md updated
- [ ] Git commit & push
- [ ] .addin manifest created
- [ ] Manifest path correct

---

## TROUBLESHOOTING

| Error | Cause | Solution |
|---|---|---|
| CS0104: TaskDialog ambiguous | Import System.Windows.Forms + Autodesk.Revit.UI | Use alias: `using RevitTaskDialog = Autodesk.Revit.UI.TaskDialog;` |
| Excel process locks | COM object not released properly | Verify ExcelInteropService.Dispose() called + null fields after ReleaseComObject |
| Image tiny in Sheet | Scale factor calculation wrong | Check: ImageInstance.Width/Height *= (scalePercent/100) |
| Image not centered | GetViewCenter() fallback issue | Check View.CropBox + View.get_BoundingBox() + fallback to XYZ.Zero |
| Dimension skip some | Dim.Curve == null | Check ArrangeDimensionCommand error handling |

---

## RESOURCES & LINKS

- **Revit API Docs:** https://www.revitapidocs.com/2026/
- **Autodesk Forum:** https://forums.autodesk.com/t5/Revit-API/ct-p/area-p127
- **GitHub:** https://github.com/duyquang868/ArcTool
- **Excel COM Interop:** https://docs.microsoft.com/en-us/office/vba/api/overview/

---

*ArcTool Development & Usage Instructions © 2026*
*Last updated: SESSION 5.7 — Updated Section D (Excel to Revit refactor + COM fixes) + Added Test 4*
