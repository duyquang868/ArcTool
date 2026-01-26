using System;
using System.Collections.Generic; // Cần thiết cho List<>
using System.Reflection;
using Autodesk.Revit.UI;
// using System.Windows.Media.Imaging; // Uncomment nếu bạn có icon

namespace Arctool.Core
{
    public class App : IExternalApplication
    {
        public Result OnStartup(UIControlledApplication application)
        {
            string tabName = "ArcTool";

            // Tạo Tab nếu chưa có
            try { application.CreateRibbonTab(tabName); } catch { }

            string assemblyPath = Assembly.GetExecutingAssembly().Location;

            // --- PANEL 1: GRAPHICS TOOLS (Cũ) ---
            RibbonPanel panelGraphics = GetOrCreatePanel(application, tabName, "Graphics Tools");

            PushButtonData btnFilterData = new PushButtonData(
                "btnFilterManager",
                "Filter\nManager",
                assemblyPath,
                "Arctool.Core.Commands.FilterManagerCommand");

            btnFilterData.ToolTip = "Tạo, quản lý và Copy Filters giữa các View/Templates.";
            // btnFilterData.LargeImage = ...
            panelGraphics.AddItem(btnFilterData);


            // --- PANEL 2: MODELING TOOLS (Mới) ---
            RibbonPanel panelModeling = GetOrCreatePanel(application, tabName, "Modeling Tools");

            PushButtonData btnVoidData = new PushButtonData(
                "btnCreateVoidLink",
                "Create Void\nFrom Link",
                assemblyPath,
                "Arctool.Core.Commands.CreateVoidFromLinkCommand"); // Namespace và tên class phải chính xác

            btnVoidData.ToolTip = "Tạo khối Void cắt tường dựa trên dầm trong file Link.";
            // Nếu có icon, thêm ở đây:
            // btnVoidData.LargeImage = new BitmapImage(new Uri("pack://application:,,,/Arctool.Core;component/Resources/VoidIcon_32.png"));

            panelModeling.AddItem(btnVoidData);

            return Result.Succeeded;
        }

        // Hàm helper giữ nguyên như cũ
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
