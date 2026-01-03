using System;
using System.Collections.Generic;
using System.Reflection;
using Autodesk.Revit.UI;
// using System.Windows.Media.Imaging; // Requires PresentationCore reference

namespace Arctool.Core
{
    public class App : IExternalApplication
    {
        public Result OnStartup(UIControlledApplication application)
        {
            string tabName = "ArcTool";
            string panelName = "Graphics Tools";

            try { application.CreateRibbonTab(tabName); } catch { }

            RibbonPanel panel = GetOrCreatePanel(application, tabName, panelName);
            string assemblyPath = Assembly.GetExecutingAssembly().Location;

            // ĐỔI TÊN NÚT: Từ Wall Filter thành Filter Manager
            PushButtonData btnFilterData = new PushButtonData(
                "btnFilterManager",
                "Filter\nManager", 
                assemblyPath,
                "Arctool.Core.Commands.FilterManagerCommand"); // Chú ý tên Class lệnh ở đây

            btnFilterData.ToolTip = "Tạo, quản lý và Copy Filters giữa các View/Templates.";

            // Thêm icon 32x32 (Đảm bảo file ảnh có Build Action là 'Resource')
            // btnFilterData.LargeImage = new BitmapImage(new Uri("pack://application:,,,/Arctool.Core;component/Resources/FilterIcon_32.png"));
            
            panel.AddItem(btnFilterData);

            return Result.Succeeded;
        }

        private RibbonPanel GetOrCreatePanel(UIControlledApplication app, string tabName, string panelName)
        {
            List<RibbonPanel> panels = app.GetRibbonPanels(tabName);
            foreach (RibbonPanel p in panels)
            {
                if (p.Name == panelName) return p;
            }
            return app.CreateRibbonPanel(tabName, panelName);
        }

        public Result OnShutdown(UIControlledApplication application) => Result.Succeeded;
    }
}