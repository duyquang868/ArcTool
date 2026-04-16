using System;
using System.Collections.Generic;
using System.Linq;
using Autodesk.Revit.Attributes;
using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using ArcTool.UI;
using Autodesk.Revit.UI.Events;

namespace ArcTool.Core.Commands
{
    [Transaction(TransactionMode.Manual)]
    public class FilterManagerCommand : IExternalCommand
    {
        // Giữ instance tĩnh để cửa sổ không bị giải phóng bộ nhớ khi lệnh kết thúc
        private static FilterWindow _ui;
        private static DateTime _lastUpdate = DateTime.MinValue;

        public Result Execute(ExternalCommandData commandData, ref string message, ElementSet elements)
        {
            UIApplication uiapp = commandData.Application;
            Document doc = uiapp.ActiveUIDocument.Document;

            // Nếu cửa sổ đang mở thì chỉ cần đưa lên trên cùng
            if (_ui != null && _ui.IsVisible)
            {
                _ui.Focus();
                return Result.Succeeded;
            }

            _ui = new FilterWindow();

            // Load dữ liệu lần đầu
            RefreshAllData(doc);

            // Đăng ký sự kiện Idling để cập nhật Real-time
            uiapp.Idling += OnIdling;

            // Thiết lập cửa sổ Modeless (.Show)
            var helper = new System.Windows.Interop.WindowInteropHelper(_ui);
            helper.Owner = Autodesk.Windows.ComponentManager.ApplicationWindow;

            _ui.Closed += (s, e) => {
                uiapp.Idling -= OnIdling; // Hủy đăng ký khi đóng
                _ui = null;
            };

            _ui.Show();
            return Result.Succeeded;
        }

        // Sự kiện quét thay đổi mỗi khi Revit rảnh
        private void OnIdling(object sender, IdlingEventArgs e)
        {
            UIApplication uiapp = sender as UIApplication;
            if (uiapp.ActiveUIDocument == null) return;

            Document doc = uiapp.ActiveUIDocument.Document;
            Autodesk.Revit.DB.View activeView = doc.ActiveView;

            // Giới hạn tần suất cập nhật để tránh lag (ví dụ: 1 giây/lần)
            if ((DateTime.Now - _lastUpdate).TotalMilliseconds < 1000) return;
            _lastUpdate = DateTime.Now;

            // 1. Cập nhật tên View Real-time nếu có thay đổi
            if (_ui != null && _ui.ActiveViewName != activeView.Name)
            {
                _ui.UpdateActiveViewInfo(activeView.Name);
            }

            // 2. Tự động cập nhật danh sách Filter nếu số lượng thay đổi (Ví dụ bạn vừa thêm Filter mới)
            int currentFilterCount = new FilteredElementCollector(doc).OfClass(typeof(ParameterFilterElement)).Count();
            if (_ui != null && _ui.FiltersSource.Count != currentFilterCount)
            {
                RefreshAllData(doc);
            }
        }

        private void RefreshAllData(Document doc)
        {
            // Lọc trùng Filter toàn dự án
            var filters = new FilteredElementCollector(doc)
                .OfClass(typeof(ParameterFilterElement))
                .Cast<ParameterFilterElement>()
                .GroupBy(f => f.Name)
                .Select(g => g.First())
                .OrderBy(x => x.Name)
                .ToList();

            _ui.Dispatcher.Invoke(() => {
                _ui.FiltersSource.Clear();
                foreach (var f in filters)
                {
                    _ui.FiltersSource.Add(new FilterItem { Name = f.Name, IsEnabled = true, IsVisible = true, Data = f });
                }

                // Cập nhật danh sách View/Sheet
                var views = new FilteredElementCollector(doc)
                    .OfClass(typeof(Autodesk.Revit.DB.View))
                    .Cast<Autodesk.Revit.DB.View>()
                    .Where(v => !v.IsTemplate && (v.ViewType == ViewType.DrawingSheet || v.CanUseTemporaryVisibilityModes()))
                    .OrderBy(v => v.Name)
                    .ToList();

                _ui.ViewsSource.Clear();
                foreach (var v in views)
                {
                    _ui.ViewsSource.Add(new ViewItem { ViewName = v.Name, FilterCount = v.GetFilters().Count, IsSelected = false, Data = v });
                }
            });
        }
    }
}