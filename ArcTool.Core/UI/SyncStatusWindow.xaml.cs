using System;
using System.IO;
using System.Windows;

// BUG-1 FIX: namespace phải khớp TUYỆT ĐỐI với x:Class trong XAML
// XAML khai báo: x:Class="ArcTool.Core.UI.SyncStatusWindow"
// → namespace phải là ArcTool.Core.UI, không phải ArcTool.UI
namespace ArcTool.Core.UI
{
    /// <summary>
    /// Toast popup — CHỈ hiện khi FileSystemWatcher phát hiện file Excel thay đổi.
    /// Không phải persistent window.
    ///
    /// Lifecycle:
    ///   Show()  → khi watcher báo change (sau debounce 2.5s)
    ///   Close() → user nhấn "✕" HOẶC user nhấn "Cập nhật" (sau khi raise ExternalEvent)
    ///
    /// Watcher vẫn tiếp tục chạy sau khi toast đóng — toast chỉ là thông báo,
    /// không phải vòng đời của watcher.
    /// </summary>
    public partial class SyncStatusWindow : Window
    {
        // Callback được gọi khi user nhấn "Cập nhật"
        // → ExcelToRevitCommand.ShowToast() truyền vào: () => _reopenEvent.Raise()
        private readonly Action _onUpdateClicked;

        /// <param name="changedFilePath">Đường dẫn đầy đủ của file Excel đã thay đổi</param>
        /// <param name="onUpdateClicked">Action gọi khi user nhấn "Cập nhật"</param>
        public SyncStatusWindow(string changedFilePath, Action onUpdateClicked)
        {
            InitializeComponent();
            _onUpdateClicked = onUpdateClicked ?? throw new ArgumentNullException(nameof(onUpdateClicked));

            // Hiện tên file ngắn gọn để user nhận ra đây là file nào
            TxtFileName.Text = Path.GetFileName(changedFilePath);
        }

        // BUG-3 FIX: handler này phải được wire trong XAML bằng Loaded="Window_Loaded"
        // (đã thêm vào SyncStatusWindow.xaml)
        private void Window_Loaded(object sender, RoutedEventArgs e)
        {
            // Đặt toast ở góc dưới bên phải màn hình (cách viền 20px)
            // WorkArea loại trừ Taskbar — đảm bảo toast không bị taskbar che
            var workArea = SystemParameters.WorkArea;
            Left = workArea.Right  - Width  - 20;
            Top  = workArea.Bottom - Height - 20;
        }

        // BUG-2 FIX: XAML khai báo Click="BtnApply_Click"
        // → tên method phải là BtnApply_Click, không phải BtnUpdate_Click
        private void BtnApply_Click(object sender, RoutedEventArgs e)
        {
            // 1. Đóng toast trước để UI sạch
            // 2. Raise ExternalEvent → Revit main thread sẽ mở lại dialog ExcelToRevitCommand
            _onUpdateClicked.Invoke();
            Close();
        }

        private void BtnClose_Click(object sender, RoutedEventArgs e)
        {
            // Chỉ đóng toast — watcher vẫn tiếp tục theo dõi file.
            // Nếu Excel thay đổi lần nữa, toast sẽ hiện lại.
            Close();
        }
    }
}
