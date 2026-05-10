using Autodesk.Revit.Attributes;
using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using ArcTool.Core.Services;
using ArcTool.UI;
using System;

namespace ArcTool.Core.Commands
{
    /// <summary>
    /// ExcelToRevitCommand — V3.0
    ///
    /// Entry point duy nhất cho tính năng Excel to Revit.
    /// V3.0 thay thế toàn bộ pipeline V1.0 (chọn file → scale dialog → import trực tiếp)
    /// bằng ExcelToRevitWindow — WPF dialog đầy đủ tính năng:
    ///   - DataGrid quản lý nhiều mapping Excel ↔ Revit View
    ///   - Change Detection (Status Dot xanh/đỏ/vàng)
    ///   - Smart Scale Persistence (lưu kích thước user đã resize)
    ///   - AutoSync (tự động update khi dialog mở)
    ///
    /// Transaction model:
    ///   - Command tự nó KHÔNG mở Transaction.
    ///   - Mọi Transaction do ExcelSyncEngine.ExecuteUpdate() mở bên trong window.
    ///   - Window được show modal (ShowDialog) → vẫn nằm trong API context của Execute().
    ///   - Do đó [Transaction(TransactionMode.Manual)] là đúng — không conflict.
    ///
    /// V3.0 — Phase 4 Integration: thay thế ExcelToRevitCommand.cs V1.0
    /// </summary>
    [Transaction(TransactionMode.Manual)]
    public class ExcelToRevitCommand : IExternalCommand
    {
        public Result Execute(ExternalCommandData commandData, ref string message, ElementSet elements)
        {
            UIApplication uiapp = commandData.Application;
            Document doc = uiapp.ActiveUIDocument?.Document;

            if (doc == null)
            {
                Autodesk.Revit.UI.TaskDialog.Show("ArcTool Error",
                    "Không có Document nào đang mở.");
                return Result.Failed;
            }

            // Guard sớm: nếu file .rvt chưa lưu thì ArcToolSettingsService.LoadMappings()
            // sẽ throw InvalidOperationException bên trong window.
            // Hiện dialog tại đây để UX rõ ràng hơn là để window mở rồi crash.
            if (string.IsNullOrWhiteSpace(doc.PathName))
            {
                Autodesk.Revit.UI.TaskDialog.Show("ArcTool — Excel to Revit",
                    "File Revit chưa được lưu.\n\n" +
                    "Vui lòng lưu file (.rvt) trước khi sử dụng tính năng Excel to Revit.\n" +
                    "File cài đặt (ArcTool_ExcelSync.json) sẽ được tạo tại cùng thư mục với file Revit.");
                return Result.Cancelled;
            }

            try
            {
                // Tạo window với Document context
                var window = new ExcelToRevitWindow(doc);

                // Đặt owner = cửa sổ chính Revit để dialog hiển thị đúng vị trí
                // và luôn nằm trên cửa sổ Revit (không bị che khuất)
                var helper = new System.Windows.Interop.WindowInteropHelper(window);
                helper.Owner = Autodesk.Windows.ComponentManager.ApplicationWindow;

                // ShowDialog() — Modal:
                //   - Block Execute() cho đến khi window đóng
                //   - Duy trì Revit API context → ExcelSyncEngine.ExecuteUpdate() hoạt động
                //   - Tương tự pattern WinForms dialog của CreateVoidFromLinkCommand
                window.ShowDialog();

                return Result.Succeeded;
            }
            catch (Autodesk.Revit.Exceptions.OperationCanceledException)
            {
                return Result.Cancelled;
            }
            catch (Exception ex)
            {
                message = ex.Message;
                Autodesk.Revit.UI.TaskDialog.Show("ArcTool Error",
                    $"Không thể mở cửa sổ Excel to Revit:\n{ex.Message}");
                return Result.Failed;
            }
        }
    }
}
