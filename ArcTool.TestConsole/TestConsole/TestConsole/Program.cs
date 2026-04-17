using System;
using System.IO;
using System.Diagnostics;
using ArcTool.Core.Services;

namespace ArcTool.TestConsole
{
    class Program
    {
        // Bắt buộc có STAThread cho COM Interop hoạt động ổn định
        [STAThread]
        static void Main(string[] args)
        {
            Console.OutputEncoding = System.Text.Encoding.UTF8;
            Console.WriteLine("--- ARC TOOL: EXCEL EXPORT DIAGNOSTIC ---");

            // 1. CẤU HÌNH ĐƯỜNG DẪN (INPUT & OUTPUT)
            // LƯU Ý QUAN TRỌNG: Bạn hãy sửa đường dẫn này trỏ đến file Excel thật trên máy bạn để test
            string excelFilePath = @"D:\Quang mini\OneDrive - MSFT\Plugin Revit\ArcTool\ArcTool.TestConsole\Testconsole.xlsx";

            string desktop = Environment.GetFolderPath(Environment.SpecialFolder.Desktop);
            string pngPath = Path.Combine(desktop, "Revit_Export_Result.png");

            // Xóa file kết quả cũ nếu có
            if (File.Exists(pngPath)) File.Delete(pngPath);

            // Kiểm tra file Input có tồn tại không
            if (!File.Exists(excelFilePath))
            {
                Console.ForegroundColor = ConsoleColor.Yellow;
                Console.WriteLine($"[CẢNH BÁO] Không tìm thấy file Excel tại: {excelFilePath}");
                Console.WriteLine("Vui lòng mở code 'Program.cs' và sửa lại biến 'excelFilePath' cho đúng.");
                Console.ResetColor();
                Console.ReadKey();
                return;
            }

            // 2. THỰC THI QUY TRÌNH (PIPELINE)
            Console.WriteLine($"\n[*] Đang mở Excel (Hidden Mode): {Path.GetFileName(excelFilePath)}...");

            bool isSuccess = false;

            // Sử dụng 'using' để đảm bảo Dispose() được gọi, giải phóng RAM ngay lập tức
            using (var excelService = new ExcelInteropService())
            {
                if (excelService.OpenFile(excelFilePath))
                {
                    Console.WriteLine("[*] Đã mở file thành công. Đang xử lý vùng in (Print Area)...");

                    // Gọi hàm xuất theo Print Area (Logic mới nhất)
                    // Nếu file Excel chưa set Print Area, nó sẽ tự fallback về UsedRange
                    isSuccess = excelService.ExportPrintAreaAsHighResImage(pngPath);

                    if (isSuccess)
                    {
                        Console.WriteLine("✅ Export lệnh thành công!");
                    }
                    else
                    {
                        Console.ForegroundColor = ConsoleColor.Red;
                        Console.WriteLine("❌ Export thất bại. (Kiểm tra lại Clipboard hoặc trạng thái Excel)");
                        Console.ResetColor();
                    }
                }
                else
                {
                    Console.ForegroundColor = ConsoleColor.Red;
                    Console.WriteLine("❌ Không thể khởi tạo Excel Application.");
                    Console.ResetColor();
                }
            } // Kết thúc block 'using', Excel Process sẽ được kill tại đây.

            // 3. HIỂN THỊ KẾT QUẢ
            if (isSuccess && File.Exists(pngPath))
            {
                Console.ForegroundColor = ConsoleColor.Green;
                Console.WriteLine($"\n✅ QUY TRÌNH HOÀN TẤT TUYỆT ĐỐI!");
                Console.WriteLine($"📂 File ảnh: {pngPath}");
                Console.ResetColor();

                // Tự động mở ảnh lên xem
                try
                {
                    Process.Start(new ProcessStartInfo(pngPath) { UseShellExecute = true });
                }
                catch { }
            }

            Console.WriteLine("\nNhấn phím bất kỳ để thoát...");
            Console.ReadKey();
        }
    }
}