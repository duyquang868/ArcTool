using System;
using System.IO;
// Cần cài NuGet: SkiaSharp và Svg.Skia
using SkiaSharp;
using Svg.Skia;

namespace ArcTool.Core.Services
{
    public class ImageProcessingService
    {
        /// <summary>
        /// Chuyển đổi file SVG sang PNG với độ phân giải tùy chỉnh (DPI).
        /// </summary>
        /// <param name="svgPath">Đường dẫn file SVG đầu vào.</param>
        /// <param name="pngPath">Đường dẫn file PNG đầu ra.</param>
        /// <param name="dpi">Mật độ điểm ảnh (Revit thường cần 72, 150 hoặc 300). Mặc định 300 cho nét.</param>
        /// <returns>True nếu thành công.</returns>
        public bool ConvertSvgToPng(string svgPath, string pngPath, float dpi = 300.0f)
        {
            if (!File.Exists(svgPath)) return false;

            try
            {
                using (var svg = new SKSvg())
                {
                    // 1. Load SVG
                    svg.Load(svgPath);

                    if (svg.Picture == null) return false;

                    // 2. Tính toán kích thước ảnh dựa trên DPI
                    // SVG thường có kích thước gốc (Point/Pixel). Muốn nét phải nhân tỉ lệ lên.
                    // 72 DPI là chuẩn màn hình cũ.
                    float scaleFactor = dpi / 72.0f;

                    var svgSize = svg.Picture.CullRect;
                    int width = (int)(svgSize.Width * scaleFactor);
                    int height = (int)(svgSize.Height * scaleFactor);

                    // 3. Tạo Bitmap (Canvas vẽ)
                    var info = new SKImageInfo(width, height);
                    using (var surface = SKSurface.Create(info))
                    {
                        var canvas = surface.Canvas;
                        canvas.Clear(SKColors.Transparent); // Nền trong suốt

                        // 4. Scale Matrix để vẽ hình to ra theo đúng DPI
                        var matrix = SKMatrix.CreateScale(scaleFactor, scaleFactor);

                        // 5. Vẽ SVG lên Canvas
                        canvas.DrawPicture(svg.Picture, ref matrix);

                        // 6. Lưu ra file PNG
                        using (var image = surface.Snapshot())
                        using (var data = image.Encode(SKEncodedImageFormat.Png, 100))
                        using (var stream = File.OpenWrite(pngPath))
                        {
                            data.SaveTo(stream);
                        }
                    }
                }
                return true;
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"Lỗi convert ảnh: {ex.Message}");
                return false;
            }
        }
    }
}