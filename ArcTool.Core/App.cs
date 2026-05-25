using Autodesk.Revit.UI;
using Autodesk.Revit.DB;
using System;
using System.Reflection;
using System.Windows.Media.Imaging;
using System.IO;
using System.Drawing; // Cần Add Reference: System.Drawing
using System.Windows.Media; // Cần Add Reference: PresentationCore
using ArcTool.Core.Services;

namespace ArcTool.Core
{
    public class App : IExternalApplication
    {
        private AddInId _addInId;

        /// <summary>
        /// The AddInId captured during OnStartup(). Used by commands that need to
        /// call CoordinateUpdaterService from outside the document-event context.
        /// Phase D proved that deriving AddInId from document event sender is unreliable.
        /// </summary>
        internal static AddInId AddInId { get; private set; }

        public Result OnStartup(UIControlledApplication application)
        {
            _addInId = application.ActiveAddInId;
            App.AddInId = application.ActiveAddInId;

            string tabName = "ArcTool";
            string voidPanelName = "Void Tools";
            string annotationPanelName = "Annotation Tools"; // Tên Panel mới cho các lệnh Dim/Text

            // 1. TẠO TAB ARCTOOL
            try
            {
                application.CreateRibbonTab(tabName);
            }
            catch { /* Bỏ qua lỗi nếu Tab đã tồn tại */ }

            // 2. LẤY HOẶC TẠO PANEL "VOID TOOLS"
            RibbonPanel voidPanel = null;
            foreach (RibbonPanel p in application.GetRibbonPanels(tabName))
            {
                if (p.Name == voidPanelName)
                {
                    voidPanel = p;
                    break;
                }
            }
            if (voidPanel == null) voidPanel = application.CreateRibbonPanel(tabName, voidPanelName);

            string assemblyPath = Assembly.GetExecutingAssembly().Location;

            // --- A. NHÓM LỆNH VOID MANAGER ---
            // Tạo SplitButtonData cho Void Manager
            SplitButtonData splitData = new SplitButtonData("splitBtnVoid", "Void\nManager");
            SplitButton splitBtn = voidPanel.AddItem(splitData) as SplitButton;

            if (splitBtn != null)
            {
                // Nút 1: Create Void
                PushButtonData btnCreate = new PushButtonData("btnCreateVoid", "Create\nVoid", assemblyPath, "ArcTool.Core.Commands.CreateVoidFromLinkCommand");
                // Icon assignment is optional - commented out until resources are added
                // btnCreate.LargeImage = ConvertToImageSource(Properties.Resources.icon_create_void_32);
                btnCreate.ToolTip = "Tự động tạo Void (Generic Model) tại vị trí tất cả Dầm trong file Link được chọn.";

                // Nút 2: Multi-Cut
                PushButtonData btnCut = new PushButtonData("btnMultiCut", "Multi-Cut", assemblyPath, "ArcTool.Core.Commands.MultiCutCommand");
                // Icon assignment is optional - commented out until resources are added
                // btnCut.LargeImage = ConvertToImageSource(Properties.Resources.icon_multi_cut_32);
                btnCut.ToolTip = "Cắt Tường (Walls) và Cột (Columns) bằng các Void đã tạo. Sử dụng thuật toán BoundingBox tối ưu.";

                // Thêm các nút vào SplitButton
                splitBtn.AddPushButton(btnCreate);
                splitBtn.AddSeparator();
                splitBtn.AddPushButton(btnCut);

                // Set nút mặc định là nút Create
                splitBtn.IsSynchronizedWithCurrentItem = true;
            }

            // --- B. NHÓM LỆNH ANNOTATION TOOLS (MỚI THÊM) ---
            // Lấy hoặc tạo Panel "Annotation Tools"
            RibbonPanel annotationPanel = null;
            foreach (RibbonPanel p in application.GetRibbonPanels(tabName))
            {
                if (p.Name == annotationPanelName)
                {
                    annotationPanel = p;
                    break;
                }
            }
            if (annotationPanel == null) annotationPanel = application.CreateRibbonPanel(tabName, annotationPanelName);

            // Khởi tạo nút Arrange Dimension
            PushButtonData arrangeDimBtnData = new PushButtonData(
                "btnArrangeDimension",
                "Arrange\nDimensions",
                assemblyPath,
                "ArcTool.Core.Commands.ArrangeDimensionCommand" // Trỏ đúng vào thư mục Commands như kiến trúc mới
            );

            // LƯU Ý: Mở file Properties/Resources.resx và thêm icon tên là "icon_arrange_dim_32" trước khi bỏ comment dòng dưới nhé!
            // arrangeDimBtnData.LargeImage = ConvertToImageSource(Properties.Resources.icon_arrange_dim_32);

            // Thiết lập Tooltip (UX quan trọng)
            arrangeDimBtnData.ToolTip = "Tự động sắp xếp khoảng cách các đường Dimension liên tục dựa trên hệ số Snap Distance và View Scale.";
            arrangeDimBtnData.LongDescription = "Click chọn đường Dim gốc, sau đó liên tục click chọn các đường Dim tiếp theo để hệ thống tự động tịnh tiến chúng cách đều nhau. Nhấn ESC để kết thúc lệnh.";

            // Thêm nút vào Panel
            annotationPanel.AddItem(arrangeDimBtnData);

            // --- C. NHÓM LỆNH IMPORT EXCEL (MỚI THÊM PHASE 3) ---
            // Lấy hoặc tạo Panel "Excel Tools"
            string excelToolsName = "Excel Tools";
            RibbonPanel excelPanel = null;
            foreach (RibbonPanel p in application.GetRibbonPanels(tabName))
            {
                if (p.Name == excelToolsName)
                {
                    excelPanel = p;
                    break;
                }
            }
            if (excelPanel == null) excelPanel = application.CreateRibbonPanel(tabName, excelToolsName);

            // Nút: Import Image từ Excel Export
            PushButtonData importImageBtnData = new PushButtonData(
                "btnExcelToRevit", 
                "Excel to\nRevit", 
                assemblyPath,
                "ArcTool.Core.Commands.ExcelToRevitCommand"
            );

            importImageBtnData.ToolTip = "Import ảnh PNG từ Excel Export vào Revit Sheet. Hỗ trợ tùy chỉnh vị trí và scale.";
            importImageBtnData.LongDescription = "1. Chọn file ảnh PNG (từ lệnh Excel Export)\n2. Nhập vị trí đặt ảnh (X, Y, Z) tính bằng Feet\n3. Tùy chỉnh Scale %\n4. Ảnh sẽ được insert vào Sheet hiện tại";

            excelPanel.AddItem(importImageBtnData);

            // --- D. NHÓM LỆNH COORDINATE TOOLS (PHASE A) ---
            string coordinateToolsName = "Coordinate Tools";
            RibbonPanel coordinatePanel = null;
            foreach (RibbonPanel p in application.GetRibbonPanels(tabName))
            {
                if (p.Name == coordinateToolsName)
                {
                    coordinatePanel = p;
                    break;
                }
            }
            if (coordinatePanel == null) coordinatePanel = application.CreateRibbonPanel(tabName, coordinateToolsName);

            PushButtonData registerCoordParamsBtnData = new PushButtonData(
                "btnRegisterCoordParams",
                "Register\nElement Type",
                assemblyPath,
                "ArcTool.Core.Commands.RegisterCoordParamsCommand"
            );

            registerCoordParamsBtnData.ToolTip = "Open coordinate settings and register AT_CoordX / AT_CoordY / AT_CoordZ for supported 3D elements: Structural Columns and Structural Foundations. Safe to run multiple times.";

            coordinatePanel.AddItem(registerCoordParamsBtnData);

            PushButtonData registerDetailItemBtnData = new PushButtonData(
                "btnRegisterDetailItemCoordType",
                "Register\nDetail Type",
                assemblyPath,
                "ArcTool.Core.Commands.RegisterDetailItemCoordTypeCommand");
            registerDetailItemBtnData.ToolTip =
                "Select one Detail Item instance and register its type name for coordinate processing. " +
                "The Detail Item registry is stored as JSON next to the RVT file and must be copied with the model.";
            coordinatePanel.AddItem(registerDetailItemBtnData);

            PushButtonData runBatchBtnData = new PushButtonData(
                "btnRunCoordBatch",
                "Write\nCoordinates",
                assemblyPath,
                "ArcTool.Core.Commands.RunCoordBatchCommand");
            runBatchBtnData.ToolTip =
                "Reads coordinates for all registered coordinate elements and writes them into " +
                "AT_CoordX / AT_CoordY / AT_CoordZ shared parameters. " +
                "Skips elements whose values have not changed. " +
                "Run 'Register Element Type' for 3D elements, or 'Register Detail Type' for Detail Items, if this is a new project.";
            coordinatePanel.AddItem(runBatchBtnData);

            PushButtonData toggleBtnData = new PushButtonData(
                "btnToggleCoordUpdater",
                "Auto\nUpdate",
                assemblyPath,
                "ArcTool.Core.Commands.ToggleCoordUpdaterCommand");
            toggleBtnData.ToolTip =
                "Enable or disable real-time coordinate auto-update for the current document. " +
                "When enabled, AT_CoordX / AT_CoordY / AT_CoordZ are updated automatically " +
                "whenever a registered coordinate element is moved or modified. " +
                "Current state is shown when you click the button.";
            coordinatePanel.AddItem(toggleBtnData);

            application.ControlledApplication.DocumentOpened += OnDocumentOpened;
            application.ControlledApplication.DocumentCreated += OnDocumentCreated;
            application.ControlledApplication.DocumentClosing += OnDocumentClosing;

            return Result.Succeeded;
        }

        public Result OnShutdown(UIControlledApplication application)
        {
            return Result.Succeeded;
        }

        private void OnDocumentOpened(object sender, Autodesk.Revit.DB.Events.DocumentOpenedEventArgs e)
        {
            CoordinateUpdaterService.RegisterForDocument(e.Document, _addInId);
        }

        private void OnDocumentCreated(object sender, Autodesk.Revit.DB.Events.DocumentCreatedEventArgs e)
        {
            if (e.Document == null)
            {
                return;
            }

            CoordinateUpdaterService.RegisterForDocument(e.Document, _addInId);
        }

        private void OnDocumentClosing(object sender, Autodesk.Revit.DB.Events.DocumentClosingEventArgs e)
        {
            CoordinateUpdaterService.UnregisterForDocument(e.Document, _addInId);
        }

        // --- HÀM HỖ TRỢ CHUYỂN ĐỔI ẢNH TỪ RESOURCE SANG REVIT ---
        public static ImageSource ConvertToImageSource(Bitmap bitmap)
        {
            if (bitmap == null) return null;

            using (MemoryStream memory = new MemoryStream())
            {
                bitmap.Save(memory, System.Drawing.Imaging.ImageFormat.Png);
                memory.Position = 0;

                BitmapImage bitmapImage = new BitmapImage();
                bitmapImage.BeginInit();
                bitmapImage.StreamSource = memory;
                bitmapImage.CacheOption = BitmapCacheOption.OnLoad;
                bitmapImage.EndInit();

                return bitmapImage;
            }
        }
    }
}