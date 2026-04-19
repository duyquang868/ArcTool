# ARCTOOL — AI SESSION CONTEXT
> Last updated: 2026-04-19 — Session 5.12: ExcelToRevit V3.0 — Full Rebuild

---

## 1. THÔNG TIN DỰ ÁN

| Mục | Chi tiết |
|---|---|
| Tên dự án | ArcTool |
| Namespace chính | `ArcTool.Core` |
| Nền tảng | Autodesk Revit 2026 (API 2026) |
| Ngôn ngữ | C# / .NET 8.0 |
| IDE | Visual Studio Enterprise 2026 |
| UI Framework | WPF (modeless windows) + WinForms (dialogs) |
| Revit Unit System | `UnitTypeId` (ForgeTypeId) |

---

## 2. CẤU TRÚC THƯ MỤC

```
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
```

---

## 3. TRẠNG THÁI HIỆN TẠI — BUG REGISTRY (ExcelToRevit ecosystem)

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

```csharp
using System;
using System.IO;
using System.Windows;
using System.Windows.Media;

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
```

---

## 6. CHECKLIST SỬA TỪNG FILE

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

### Giai đoạn 2 — Filter Manager
- [ ] **[ƯU TIÊN #1]** Thêm Ribbon Button cho FilterManagerCommand vào App.cs
- [ ] Implement Copy/Paste Filter, MVVM binding, Idling → ExternalEvent

### Giai đoạn 4 — Quick Dim (R&D)
- [ ] `GetDimensionableReferences(Element)` trước `NewDimension()`

---

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
