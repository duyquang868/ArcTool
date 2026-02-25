using Autodesk.Revit.UI;
using System;
using System.Reflection;
using System.Windows.Media.Imaging;
using System.IO;
using System.Drawing; // Cần Add Reference: System.Drawing
using System.Windows.Media; // Cần Add Reference: PresentationCore

namespace ArcTool.Core
{
    public class App : IExternalApplication
    {
        public Result OnStartup(UIControlledApplication application)
        {
            string tabName = "ArcTool";
            string panelName = "Void Tools";

            // 1. Tạo Tab và Panel
            try { application.CreateRibbonTab(tabName); } catch { }
            RibbonPanel panel = null;
            foreach (RibbonPanel p in application.GetRibbonPanels(tabName))
            {
                if (p.Name == panelName)
                {
                    panel = p;
                    break;
                }
            }
            if (panel == null) panel = application.CreateRibbonPanel(tabName, panelName);

            string assemblyPath = Assembly.GetExecutingAssembly().Location;

            // 2. TẠO SPLIT BUTTON (NÚT CHIA ĐÔI - GIAO DIỆN PRO)
            // SplitButton có Icon to ở trên, Tên và Mũi tên to ở dưới.
            SplitButtonData splitData = new SplitButtonData("splitBtnVoid", "Void\nManager");
            SplitButton splitBtn = panel.AddItem(splitData) as SplitButton;

            // 3. ĐỊNH NGHĨA LỆNH 1: TẠO VOID (CREATE)
            PushButtonData btnCreate = new PushButtonData(
                "btnCreateVoid",
                "Create Void\n(Auto Link)",
                assemblyPath,
                "ArcTool.Core.Commands.CreateVoidFromLinkCommand");

            btnCreate.ToolTip = "Tự động tạo Void từ tất cả dầm trong file Link.";

            // --- GẮN RESOURCE TỪ FILE CỦA BẠN ---
            // Lưu ý: Thay 'icon_create_32' bằng tên file thật bạn đã import trong Resources
            btnCreate.LargeImage = ConvertToImageSource(Properties.Resources.icon_create_32);
            btnCreate.Image = ConvertToImageSource(Properties.Resources.icon_create_16); // Icon nhỏ (nếu có)

            // 4. ĐỊNH NGHĨA LỆNH 2: CẮT (MULTI-CUT)
            PushButtonData btnCut = new PushButtonData(
                "btnMultiCut",
                "Multi-Cut\n(Wall & Col)",
                assemblyPath,
                "ArcTool.Core.Commands.MultiCutCommand");

            btnCut.ToolTip = "Quét chọn để cắt Tường và Cột.";

            // --- GẮN RESOURCE TỪ FILE CỦA BẠN ---
            // Lưu ý: Thay 'icon_cut_32' bằng tên file thật bạn đã import trong Resources
            btnCut.LargeImage = ConvertToImageSource(Properties.Resources.icon_cut_32);
            btnCut.Image = ConvertToImageSource(Properties.Resources.icon_cut_16); // Icon nhỏ (nếu có)

            // 5. THÊM VÀO NÚT TỔNG + KẺ NGANG
            if (splitBtn != null)
            {
                // Thêm nút Create
                splitBtn.AddPushButton(btnCreate);

                // Thêm dòng kẻ ngang phân cách (Separator)
                splitBtn.AddSeparator();

                // Thêm nút Cut
                splitBtn.AddPushButton(btnCut);

                // Set nút mặc định là nút Create (Hiện icon này lên mặt tiền)
                splitBtn.IsSynchronizedWithCurrentItem = true;
            }

            return Result.Succeeded;
        }

        public Result OnShutdown(UIControlledApplication application)
        {
            return Result.Succeeded;
        }

        // --- HÀM HỖ TRỢ CHUYỂN ĐỔI ẢNH TỪ RESOURCE SANG REVIT ---
        // Hàm này giúp lấy ảnh từ Properties.Resources (System.Drawing.Bitmap)
        // và chuyển thành ImageSource mà Revit hiểu được.
        public static ImageSource ConvertToImageSource(Bitmap bitmap)
        {
            if (bitmap == null) return null;

            using (MemoryStream memory = new MemoryStream())
            {
                // Lưu bitmap vào memory stream dưới dạng PNG
                bitmap.Save(memory, System.Drawing.Imaging.ImageFormat.Png);
                memory.Position = 0;

                // Tạo BitmapImage từ stream
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