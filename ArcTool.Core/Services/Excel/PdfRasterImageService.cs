using System;
using System.IO;
using System.Runtime.InteropServices;
using PDFtoImage;
using SkiaSharp;

namespace ArcTool.Core.Services.Excel
{
    /// <summary>
    /// Engine-agnostic raster stage of the Excel/WPS-to-Revit image pipeline: renders a PDF's first
    /// page to a 300 DPI PNG (PDFtoImage/PDFium) and crops the surrounding white margins (SkiaSharp,
    /// threshold 240). This class contains zero COM, zero spreadsheet knowledge, and zero Interop.
    ///
    /// Engine-agnostic port of the legacy single-provider render/crop stage plus native-loader helpers,
    /// extracted during the MS-Excel/WPS provider split.
    ///
    /// Temp-PDF lifetime note: this service only reads <c>pdfPath</c> and writes
    /// <c>outputPngPath</c>. It never deletes the input PDF. In the original code the temp PDF was
    /// deleted in a <c>finally</c> block that lived in the same routine as the render/crop. After
    /// this split, deleting the temp PDF is the caller's responsibility (the coordinator,
    /// <c>SpreadsheetImageExportService</c>) — do not add a delete here.
    /// </summary>
    public static class PdfRasterImageService
    {
        private static readonly object PdfiumLoadLock = new object();
        private static readonly object SkiaLoadLock = new object();

        private static bool _pdfiumLoadAttempted;
        private static IntPtr _pdfiumHandle;

        private static bool _skiaLoadAttempted;
        private static IntPtr _skiaHandle;

        /// <summary>
        /// Renders <paramref name="pdfPath"/> (page 0) to a 300 DPI PNG at
        /// <paramref name="outputPngPath"/>, then crops the surrounding white margins (threshold 240)
        /// in place. Engine-agnostic port of the legacy render/crop routine: same DPI, same threshold,
        /// same four-direction scan order (top, bottom, left,
        /// right), same edge-case behavior — an all-white page, a zero-size crop rectangle, or a
        /// failed <c>ExtractSubset</c>/encode are not specially guarded here, exactly as they were not
        /// guarded in the original method. Does not delete <paramref name="pdfPath"/>; the caller owns
        /// that file's lifetime.
        /// </summary>
        /// <returns>
        /// <c>false</c> only if pdfium failed to load (render never attempted); <c>true</c> in every
        /// other case, including when SkiaSharp fails to load (PNG was written, crop skipped) and
        /// every early-return crop branch below — exactly as <c>ExportRangeInternal</c> returned
        /// today. Exceptions are logged then rethrown, matching the original's
        /// "propagate to caller, do not swallow" policy.
        /// </returns>
        public static bool RenderPdfToCroppedPng(string pdfPath, string outputPngPath)
        {
            try
            {
                if (!EnsurePdfiumLoaded())
                {
                    System.Diagnostics.Debug.WriteLine(
                        "PdfRasterImageService Export Error: pdfium.dll không load được từ add-in folder hoặc runtimes/win-*/native.");
                    return false;
                }

                using var pdfStream = File.OpenRead(pdfPath);
                Conversion.SavePng(
                    outputPngPath,
                    pdfStream,
                    0,
                    false,
                    null,
                    new RenderOptions { Dpi = 300, WithAnnotations = false });

                if (!EnsureSkiaSharpLoaded())
                {
                    System.Diagnostics.Debug.WriteLine(
                        "PdfRasterImageService Export Error: libSkiaSharp.dll không load được từ add-in folder, native, hoặc runtimes/win-*/native.");
                    return true;
                }

                const byte WhiteThreshold = 240;
                using var bitmap = SKBitmap.Decode(outputPngPath);
                if (bitmap == null || bitmap.Width <= 0 || bitmap.Height <= 0)
                    return true;

                int top = 0;
                while (top < bitmap.Height)
                {
                    bool found = false;
                    for (int x = 0; x < bitmap.Width; x++)
                    {
                        if (IsNonWhitePixel(bitmap.GetPixel(x, top), WhiteThreshold))
                        {
                            found = true;
                            break;
                        }
                    }

                    if (found) break;
                    top++;
                }

                if (top >= bitmap.Height)
                    return true;

                int bottom = bitmap.Height - 1;
                while (bottom >= top)
                {
                    bool found = false;
                    for (int x = 0; x < bitmap.Width; x++)
                    {
                        if (IsNonWhitePixel(bitmap.GetPixel(x, bottom), WhiteThreshold))
                        {
                            found = true;
                            break;
                        }
                    }

                    if (found) break;
                    bottom--;
                }

                int left = 0;
                while (left < bitmap.Width)
                {
                    bool found = false;
                    for (int y = top; y <= bottom; y++)
                    {
                        if (IsNonWhitePixel(bitmap.GetPixel(left, y), WhiteThreshold))
                        {
                            found = true;
                            break;
                        }
                    }

                    if (found) break;
                    left++;
                }

                int right = bitmap.Width - 1;
                while (right >= left)
                {
                    bool found = false;
                    for (int y = top; y <= bottom; y++)
                    {
                        if (IsNonWhitePixel(bitmap.GetPixel(right, y), WhiteThreshold))
                        {
                            found = true;
                            break;
                        }
                    }

                    if (found) break;
                    right--;
                }

                int cropWidth = right - left + 1;
                int cropHeight = bottom - top + 1;
                if (cropWidth <= 0 || cropHeight <= 0)
                    return true;

                var cropRect = new SKRectI(left, top, right + 1, bottom + 1);
                var cropInfo = new SKImageInfo(cropWidth, cropHeight, bitmap.ColorType, bitmap.AlphaType);
                using var cropped = new SKBitmap(cropInfo);
                if (!bitmap.ExtractSubset(cropped, cropRect))
                    return true;

                using var image = SKImage.FromBitmap(cropped);
                using var data = image.Encode(SKEncodedImageFormat.Png, 100);
                if (data == null)
                    return true;

                using var outputStream = File.Open(outputPngPath, FileMode.Create, FileAccess.Write, FileShare.None);
                data.SaveTo(outputStream);
                return true;
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"PdfRasterImageService Export Error: {ex.Message}");
                throw; // Propagate lên caller — không nuốt exception
            }
        }

        private static bool IsNonWhitePixel(SKColor color, byte whiteThreshold)
        {
            return color.Alpha < whiteThreshold
                || color.Red < whiteThreshold
                || color.Green < whiteThreshold
                || color.Blue < whiteThreshold;
        }

        private static string GetRuntimeFolder()
        {
            return RuntimeInformation.ProcessArchitecture switch
            {
                Architecture.X64 => "win-x64",
                Architecture.X86 => "win-x86",
                Architecture.Arm64 => "win-arm64",
                _ => null
            };
        }

        private static string[] GetNativeLibraryCandidates(string libraryFileName)
        {
            string assemblyDir = Path.GetDirectoryName(typeof(PdfRasterImageService).Assembly.Location);
            if (string.IsNullOrWhiteSpace(assemblyDir))
                return Array.Empty<string>();

            string runtimeFolder = GetRuntimeFolder();
            if (string.IsNullOrWhiteSpace(runtimeFolder))
                return Array.Empty<string>();

            return new[]
            {
                Path.Combine(assemblyDir, libraryFileName),
                Path.Combine(assemblyDir, "native", libraryFileName),
                Path.Combine(assemblyDir, "runtimes", runtimeFolder, "native", libraryFileName)
            };
        }

        private static bool EnsurePdfiumLoaded()
        {
            lock (PdfiumLoadLock)
            {
                if (_pdfiumLoadAttempted)
                    return _pdfiumHandle != IntPtr.Zero;

                _pdfiumLoadAttempted = true;

                if (NativeLibrary.TryLoad("pdfium", out _pdfiumHandle))
                    return true;

                foreach (string candidate in GetNativeLibraryCandidates("pdfium.dll"))
                {
                    if (!File.Exists(candidate))
                        continue;

                    if (NativeLibrary.TryLoad(candidate, out _pdfiumHandle))
                        return true;
                }

                return false;
            }
        }

        private static bool EnsureSkiaSharpLoaded()
        {
            lock (SkiaLoadLock)
            {
                if (_skiaLoadAttempted)
                    return _skiaHandle != IntPtr.Zero;

                _skiaLoadAttempted = true;

                if (NativeLibrary.TryLoad("libSkiaSharp", out _skiaHandle))
                    return true;

                foreach (string candidate in GetNativeLibraryCandidates("libSkiaSharp.dll"))
                {
                    if (!File.Exists(candidate))
                        continue;

                    if (NativeLibrary.TryLoad(candidate, out _skiaHandle))
                        return true;
                }

                return false;
            }
        }
    }
}
