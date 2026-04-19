## ĐỐI TÁC CỦA BẠN

Tôi là kiến trúc sư với **20 năm kinh nghiệm thực chiến** trong ngành thiết kế xây dựng. Nhiệm vụ của bạn: kết hợp với kiến thức chuyên ngành của tôi để tạo ra phần mềm thực sự giải quyết vấn đề thực tế trong nghề.

---

## DỰ ÁN: ARCTOOL PLUGIN

| Mục | Chi tiết |
|---|---|
| Namespace | `ArcTool.Core` |
| Nền tảng | Autodesk Revit 2026 (API 2026) |
| Ngôn ngữ | C# (.NET 8.0) |
| IDE | Visual Studio Enterprise 2026 |
| UI | WinForms (dialogs) + WPF (modeless windows) |

---

## TÍNH NĂNG ĐÃ HOÀN THIỆN

### A. Ribbon UI — `App.cs` (V5.2)
- Panel 1: **Void Tools** → SplitButton (Create Void + Multi-Cut)
- Panel 2: **Annotation Tools** → Arrange Dimensions
- Panel 3: **Excel Tools** → Excel to Revit
- ⚠️ GAP: `FilterManagerCommand` chưa có Ribbon Button

### B. Create Void — `CreateVoidFromLinkCommand.cs` (V4.0)
- Tự động tạo Void tại TẤT CẢ dầm trong file Link, Face-Based, Transform Matrix đúng

### C. Multi-Cut — `MultiCutCommand.cs` (V2.0)
- Cắt Walls + Columns bằng Void, BoundingBox broad-phase

### D. Arrange Dimensions — `ArrangeDimensionCommand.cs` (V1.0)
- Tịnh tiến Dim theo Snap Distance × View Scale, TransactionGroup → 1 Undo

### E. Excel Export Engine — `ExcelInteropService.cs` (V5.2)
- Export Print Area → PNG (35x scale), `GetActiveSheetName()`, COM bugs đã fix

### F. Settings — `ArcToolSettings.cs`
- `%AppData%\ArcTool\settings.json`: LastScale, LastUsed, LastExcelFile

### G. Filter Manager — `FilterManagerCommand.cs` + `FilterWindow.xaml`
- WPF modeless UI, Idling real-time. ⚠️ Không có Ribbon Button

### H. Sync Status Window — `SyncStatusWindow.xaml` + `SyncStatusWindow.xaml.cs` (NEW — Session 5.10)
- WPF floating window xuất hiện góc dưới màn hình sau khi import thành công
- **🟢 Xanh** = file Excel chưa thay đổi kể từ lần import cuối
- **🔴 Đỏ** = file Excel đã được lưu lại → có thay đổi cần cập nhật
- Nút "Cập nhật": chỉ enabled khi tick đỏ → user nhấn để trigger refresh thủ công
- Nút "✕": đóng window, dừng FileSystemWatcher
- `SetStatus(bool)` thread-safe (tự Dispatcher.Invoke)

### I. Excel to Revit — `ExcelToRevitCommand.cs` (V2.1 — Session 5.10)

**Luồng thực thi:**
```
[Execute()]
  ├─ StopWatcher() — dừng watcher cũ nếu đang chạy
  ├─ PromptForExcelFile()
  ├─ ExcelInteropService { OpenFile → GetActiveSheetName → ExportPrintAreaAsHighResImage }
  ├─ ArcToolSettings.Load()
  ├─ ImportOptionsDialog { Scale, CreateNewView }
  ├─ Transaction {
  │    GetOrCreateDraftingView() [nếu chọn view mới]
  │    ImageType.Create()
  │    ImageInstance.Create()
  │    apply scale
  │    → ghi nhớ inst.Width, inst.Height vào StoredWidth/StoredHeight
  │  }
  ├─ ArcToolSettings.Save()
  ├─ Setup ExcelRefreshHandler { StoredWidth, StoredHeight, TargetViewId, ... }
  ├─ ExternalEvent.Create(handler)
  ├─ SetupWatcher(excelPath) — FileSystemWatcher bắt đầu theo dõi
  ├─ ShowStatusWindow() — SyncStatusWindow hiện góc dưới phải
  └─ CleanupTempFile()
```

**Luồng khi Excel thay đổi (Manual Sync):**
```
FileSystemWatcher.Changed/Renamed
  └─ debounce 2.5s (lock + Timer.Restart)
       └─ _statusWindow.SetStatus(hasChanges: true)  ← CHỈ đổi màu tick → ĐỎ
                                                        KHÔNG tự refresh ảnh

User nhấn "Cập nhật" trong SyncStatusWindow
  └─ _updateEvent.Raise()  ← non-blocking, từ UI thread của window
       └─ Revit main thread gọi ExcelRefreshHandler.Execute(UIApplication)
            ├─ ExcelInteropService.Export → tempPng
            ├─ Đọc existingInst.Width/Height (SMART SCALE — trước khi xóa)
            ├─ Transaction {
            │    doc.Delete(existingInst)
            │    ImageType.Create(tempPng)
            │    ImageInstance.Create(tại vị trí cũ)
            │    newInst.Width  = StoredWidth   ← kích thước thực tế user đang dùng
            │    newInst.Height = StoredHeight
            │    cập nhật ImageInstanceId
            │  }
            ├─ File.Delete(tempPng)
            └─ OnRefreshComplete() → _statusWindow.SetStatus(false)  ← đổi tick → XANH
```

**Smart Scale — Logic chi tiết:**
- `StoredWidth/StoredHeight` được đọc từ `existingInst.Width/Height` TRƯỚC khi xóa
- Nếu user đã tự tay resize ảnh trong Revit → giá trị này phản ánh kích thước mới
- Giá trị này được dùng để set kích thước ảnh mới → layout không bị phá vỡ
- `ScaleFactor` (% từ dialog) chỉ dùng cho lần import đầu tiên, không dùng trong refresh

---

## CẤU TRÚC FILE MỚI (Session 5.10)

```
ArcTool.Core/
├── UI/
│   ├── FilterWindow.xaml + .cs
│   ├── SyncStatusWindow.xaml       ← NEW: WPF floating status indicator
│   └── SyncStatusWindow.xaml.cs   ← NEW: code-behind với SetStatus(bool)
```

---

## RIBBON UI LAYOUT

```
ArcTool Tab (3 Panels — Panel 4 chờ Giai đoạn 2)
├── Void Tools Panel → SplitButton (Create Void + Multi-Cut)
├── Annotation Tools Panel → Arrange Dimensions
└── Excel Tools Panel → Excel to Revit ✅ V2.1
                        [SyncStatusWindow: floating, góc dưới phải màn hình]
```

---

## TESTING PROCEDURES

### Test 1–3: Không thay đổi (Create Void, Multi-Cut, Arrange Dimensions)

### Test 4: Excel to Revit V2.1 — Manual Sync + Smart Scale

**Bước 1: Import lần đầu**
1. ArcTool → Excel Tools → Excel to Revit
2. Chọn file .xlsx có Print Area
3. Kiểm tra dialog: tên sheet, scale pre-fill từ lần trước
4. Nhấn OK

**Assertions sau import:**
- [ ] Ảnh xuất hiện trong View đúng scale
- [ ] `SyncStatusWindow` hiện góc dưới phải, tick 🟢
- [ ] `settings.json` được cập nhật

**Bước 2: Test phát hiện thay đổi**
1. Mở file Excel, sửa nội dung, Ctrl+S
2. Chờ ~2.5 giây
3. **Expected:** Tick chuyển 🔴 "File Excel đã thay đổi"
4. **Kiểm tra:** Revit KHÔNG tự thay ảnh → đây là hành vi mong muốn

**Assertions:**
- [ ] Tick đỏ sau ~2.5s kể từ khi lưu Excel
- [ ] Ảnh trong Revit KHÔNG thay đổi (chờ user nhấn)
- [ ] Nút "Cập nhật" enabled

**Bước 3: Test Manual Refresh**
1. Nhấn "Cập nhật" trong SyncStatusWindow
2. **Expected:** Ảnh được thay thế, tick chuyển 🟢

**Assertions:**
- [ ] Ảnh mới tại đúng vị trí ảnh cũ
- [ ] Kích thước ảnh mới = kích thước ảnh cũ (trước khi nhấn Cập nhật)
- [ ] Tick trở về 🟢

**Bước 4: Test Smart Scale**
1. Sau import lần đầu (scale 100%), kéo resize ảnh trong Revit lên ~150%
2. Mở Excel, sửa, Ctrl+S → chờ tick đỏ → nhấn "Cập nhật"
3. **Expected:** Ảnh mới có kích thước 150% (không phải 100% gốc)

**Assertions:**
- [ ] Kích thước ảnh mới = kích thước ảnh SAU KHI user resize (không phải scale từ dialog)
- [ ] Layout không bị phá vỡ

**Bước 5: Test đóng window**
1. Nhấn ✕ trên SyncStatusWindow
2. Sửa Excel, Ctrl+S
3. **Expected:** Tick không xuất hiện, watcher đã dừng (không còn theo dõi)

### Test 5: Chạy lại lệnh (watcher restart)
1. Chạy Excel to Revit lần 1 → SyncStatusWindow hiện
2. Chạy lại lệnh → chọn file Excel khác
3. **Expected:** Window cũ đóng, window mới hiện theo dõi file mới

---

## CÁC TASK PHÁT TRIỂN THƯỜNG GẶP

### Thay đổi màu sắc SyncStatusWindow
- File: `SyncStatusWindow.xaml.cs`
- Static brushes: `GreenBrush` và `RedBrush` ở đầu class
- XAML: `StatusDot` Ellipse, `TxtStatus` TextBlock

### Thêm Setting mới vào ArcToolSettings
1. Thêm property + `[JsonPropertyName("...")]` + default value
2. Dùng trong dialog sau `settings.Load()`
3. Save sau `t.Commit()`: `settings.XxxProp = value; settings.Save();`

### Thêm Smart Scale cho lệnh khác
```csharp
// Đọc kích thước TRƯỚC KHI xóa
double savedW = existingElement.Width;
double savedH = existingElement.Height;

// Xóa và tạo mới...

// Áp dụng kích thước đã lưu
newElement.Width  = savedW;
newElement.Height = savedH;
```

---

## BUILD & DEPLOYMENT

```powershell
dotnet build -c Debug
```

**Output:** `ArcTool.Core\Bin\x64\Debug\net8.0-windows\`

**Deployment Checklist:**
- [ ] Build successful (0 errors)
- [ ] `SyncStatusWindow.xaml` + `.xaml.cs` trong `UI/` folder
- [ ] `ArcToolSettings.cs` trong `Services/` folder
- [ ] `ExcelInteropService.cs` V5.2 (có `GetActiveSheetName()`)
- [ ] CONTEXT.md + SKILL.md + Instructions.md updated
- [ ] Git commit & push

---

## TROUBLESHOOTING

| Error | Cause | Solution |
|---|---|---|
| CS0104: TaskDialog ambiguous | WinForms + Revit.UI import | `using RevitTaskDialog = Autodesk.Revit.UI.TaskDialog;` |
| Tick không đổi đỏ sau khi lưu Excel | Watcher không bắt event | Check `NotifyFilter = LastWrite \| Size`, hook cả `Renamed` event |
| Tick không đổi xanh sau Cập nhật | `OnRefreshComplete` chưa được set | Gán `_handler.OnRefreshComplete = () => _statusWindow?.SetStatus(false)` |
| Ảnh refresh về kích thước sai | Dùng ScaleFactor thay vì đọc từ instance | Đọc `existingInst.Width/Height` TRƯỚC `doc.Delete()` |
| `IsValidObject` false crash | Element bị xóa bởi user | Guard: `if (existingInst != null && existingInst.IsValidObject)` |
| SyncStatusWindow không hiện | `Show()` từ non-UI thread | Đảm bảo `ShowStatusWindow()` được gọi trong `Execute()` của IExternalCommand |
| Debounce trigger nhiều lần | Timer không được lock đúng | Đảm bảo `lock(_debounceLock)` bao toàn bộ Stop + Dispose + new Timer |
| Drafting View tạo trùng | Không check tên trước | `GetOrCreateDraftingView()`: `FirstOrDefault` tên trước `ViewDrafting.Create()` |

---

## REFERENCES

- **Revit API Docs:** https://www.revitapidocs.com/2026/
- **GitHub:** https://github.com/duyquang868/ArcTool

---

*Last updated: Session 5.10 — ExcelToRevitCommand V2.1 (Manual Sync + Smart Scale) + SyncStatusWindow*
