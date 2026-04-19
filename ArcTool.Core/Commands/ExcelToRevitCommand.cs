using System;
using System.IO;
using System.Windows;
using Autodesk.Revit.Attributes;
using Autodesk.Revit.DB;
using Autodesk.Revit.DB.Events;
using Autodesk.Revit.UI;
using Autodesk.Revit.UI.Events;
using ArcTool.Core.UI;

// BUG-1 FIX: Namespace alias để tránh ambiguity giữa Revit TaskDialog và Windows TextBox
using RevitTaskDialog = Autodesk.Revit.UI.TaskDialog;
using Timer = System.Timers.Timer;

namespace ArcTool.Core.Commands
{
    [Transaction(TransactionMode.Manual)]
    public class ExcelToRevitCommand : IExternalCommand
    {
        private static FileSystemWatcher _watcher;
        private static Timer _debounceTimer;
        private static SyncStatusWindow _currentToast;
        private static ExternalEvent _reopenEvent;

        public Result Execute(ExternalCommandData commandData, ref string message, ElementSet elements)
        {
            UIApplication uiApp = commandData.Application;
            UIDocument uidoc = uiApp.ActiveUIDocument;
            Document doc = uidoc.Document;

            try
            {
                // Initialize external event handler for Revit API operations
                if (_reopenEvent == null)
                {
                    _reopenEvent = ExternalEvent.Create(new ReopenHandler(uidoc));
                }

                // Prompt user to select Excel file
                RevitTaskDialog taskDialog = new RevitTaskDialog("Excel to Revit");
                taskDialog.MainInstruction = "Select an Excel file to import";
                TaskDialogResult result = taskDialog.Show();

                // For demo purposes, use a placeholder path
                string excelPath = @"C:\Temp\sample.xlsx";

                if (!string.IsNullOrEmpty(excelPath) && File.Exists(excelPath))
                {
                    SetupWatcher(excelPath);
                    RevitTaskDialog.Show("Success", $"Now watching {excelPath} for changes");
                }

                return Result.Succeeded;
            }
            catch (Exception ex)
            {
                message = ex.Message;
                return Result.Failed;
            }
        }

        /// <summary>
        /// Setup file watcher to monitor Excel file changes.
        /// Handles both Changed and Renamed events to catch Office save operations.
        /// </summary>
        private static void SetupWatcher(string excelPath)
        {
            // Dispose watcher cũ (nếu có) trước khi tạo mới
            StopWatcher();

            string dir = Path.GetDirectoryName(excelPath);
            string file = Path.GetFileName(excelPath);
            if (string.IsNullOrEmpty(dir)) return;

            _watcher = new FileSystemWatcher(dir, file)
            {
                // BUG-4 FIX: PHẢI thêm NotifyFilters.FileName để Renamed event hoạt động.
                //
                // Excel/Office lưu file theo cơ chế:
                //   1. Ghi nội dung vào file temp (e.g. ~$Report.xlsx)    → LastWrite event
                //   2. Rename file temp thành tên gốc (Report.xlsx)       → Renamed event ← BỊ MISS
                //   3. Xóa file gốc cũ
                //
                // Nếu không có FileName trong NotifyFilter, Renamed event KHÔNG BAO GIỜ fire
                // → hook _watcher.Renamed vô nghĩa → nhiều trường hợp save không phát hiện được.
                NotifyFilter = NotifyFilters.LastWrite
                             | NotifyFilters.Size
                             | NotifyFilters.FileName, // ← thêm dòng này
                EnableRaisingEvents = true
            };

            // Hook cả Changed và Renamed.
            // Excel/Office save = "write temp file → rename → delete old" → cần bắt Renamed.
            _watcher.Changed += (s, e) => ScheduleToast(e.FullPath);
            _watcher.Renamed += (s, e) => ScheduleToast(e.FullPath);
        }

        /// <summary>
        /// Schedule toast display with debouncing to prevent multiple notifications for single save.
        /// </summary>
        private static void ScheduleToast(string changedFilePath)
        {
            _debounceTimer?.Stop();
            _debounceTimer?.Dispose();
            _debounceTimer = new Timer(2500) { AutoReset = false };
            _debounceTimer.Elapsed += (s, args) =>
            {
                ShowToast(changedFilePath);
            };
            _debounceTimer.Start();
        }

        /// <summary>
        /// FileSystemWatcher gọi trên background thread.
        /// Phải marshal sang WPF UI thread trước khi tạo/show Window.
        /// </summary>
        private static void ShowToast(string changedFilePath)
        {
            // BUG-5 FIX: Application.Current có thể null trong một số Revit plugin contexts.
            // Guard null trước khi access Dispatcher để tránh NullReferenceException
            // trên background thread — exception này bị swallow âm thầm, toast không hiện.
            var dispatcher = System.Windows.Application.Current?.Dispatcher;
            if (dispatcher == null) return;

            dispatcher.BeginInvoke(new Action(() =>
            {
                // Đóng toast cũ để tránh stack nhiều popup
                _currentToast?.Close();
                _currentToast = null;

                _currentToast = new SyncStatusWindow(
                    changedFilePath: changedFilePath,
                    onUpdateClicked: () =>
                    {
                        // User nhấn "Cập nhật" → Raise ExternalEvent
                        // Revit sẽ gọi ReopenHandler.Execute() trên main thread
                        // → mở lại toàn bộ dialog ExcelToRevitCommand
                        _reopenEvent?.Raise();
                    }
                );

                // Null reference khi toast tự đóng (user bấm ✕ hoặc sau khi Cập nhật)
                _currentToast.Closed += (s, e) =>
                {
                    if (ReferenceEquals(s, _currentToast))
                        _currentToast = null;
                };

                _currentToast.Show();
            }));
        }

        /// <summary>
        /// Stop the file watcher and clean up resources.
        /// </summary>
        private static void StopWatcher()
        {
            _watcher?.Dispose();
            _watcher = null;

            _debounceTimer?.Stop();
            _debounceTimer?.Dispose();
            _debounceTimer = null;

            // Close toast if it's open
            _currentToast?.Close();
            _currentToast = null;
        }

        /// <summary>
        /// External event handler to reopen the ExcelToRevitCommand dialog.
        /// </summary>
        private class ReopenHandler : IExternalEventHandler
        {
            private readonly UIDocument _uidoc;

            public ReopenHandler(UIDocument uidoc)
            {
                _uidoc = uidoc;
            }

            public void Execute(UIApplication app)
            {
                // Placeholder: Reopen the dialog or trigger update
                RevitTaskDialog.Show("Update", "File was updated. Ready to reimport.");
            }

            public string GetName()
            {
                return "ExcelToRevitCommand_Reopen";
            }
        }
    }
}
