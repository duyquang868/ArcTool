# RESEARCH — ExcelInteropService V6.0 Export Fix
> Session: 2026-05-13 | Phạm vi: `Services/ExcelInteropService.cs` → V6.0
> Status: Research hoàn tất, sẵn sàng implement

---

## 1. Vấn đề gốc (Root Cause)

`ExportRangeInternal()` dùng `CopyPicture(xlScreen)` → bị giới hạn ~20 dòng.

**Cơ chế lỗi:**
- Excel `Visible = false` → không có real screen DC
- Excel tạo virtual DC với chiều cao cố định theo màn hình vật lý (~700-800 points)
- Row height mặc định ~15 points → tối đa ~46 row trên lý thuyết
- Hidden mode overhead → thực tế chỉ ~20 row
- Nội dung vượt quá chiều cao DC bị **cắt silently** — không có exception, không cảnh báo
- `chart.Paste()` → paste nội dung bị cắt → PNG bị trắng hoặc thiếu phần dưới

**Lý do đổi sang `xlScreen` (BUG-E7):** Để fix méo hình DPI mismatch khi dùng `xlPrinter`. Fix đúng vấn đề cũ nhưng sinh ra vấn đề mới nghiêm trọng hơn.

---

## 2. Quyết định: Hướng C — Excel → PDF → PNG

### Pipeline mới (thay thế hoàn toàn `ExportRangeInternal()`)

```
Worksheet.ExportAsFixedFormat(xlTypePDF) → temp.pdf
    ↓  (không dùng clipboard, không phụ thuộc screen DC)
PDFtoImage render 300 DPI → temp.png
    ↓  (PDFium engine — cùng Chrome)
SkiaSharp crop white margins → output.png final
```

**Ưu điểm so với pipeline cũ:**

| Tiêu chí | Pipeline cũ | Pipeline mới |
|---|---|---|
| Giới hạn số dòng | ~20 dòng | Không giới hạn |
| Clipboard | Phụ thuộc | Không dùng |
| Hidden mode | Lỗi với xlScreen | Hoàn toàn ổn |
| Chất lượng chữ | Bitmap 35x | Vector PDF → PDFium 300 DPI |
| BUG-E7 (méo hình) | Có | Không có |

---

## 3. Thư viện

### Thêm mới: `PDFtoImage` (sungaila)

```xml
<!-- Thêm vào ArcTool.Core.csproj -->
<PackageReference Include="PDFtoImage" Version="5.2.1" />
```

| Tiêu chí | Chi tiết |
|---|---|
| License | **MIT — 100% free, kể cả commercial** |
| Rendering engine | **PDFium** (Chrome) |
| Output | SkiaSharp `SKBitmap` → PNG |
| Downloads | 4.7 triệu |
| Cập nhật | April 2026 (active) |
| .NET 8 Windows | ✅ |
| SkiaSharp compatibility | ✅ Compatible với SkiaSharp 3.x đã có trong project |

```csharp
// API usage
using PDFtoImage;
using var pdfStream = File.OpenRead(tempPdfPath);
Conversion.SavePng(pdfStream, outputPngPath,
    new RenderOptions { Dpi = 300, WithAnnotations = false });
```

### Tái sử dụng: `SkiaSharp 3.119.2` (đã có trong project)

Không chỉ compatible — được dùng **tích cực** để crop white margins sau render:

```csharp
using var bmp = SKBitmap.Decode(pngPath);
// Scan 4 cạnh → tìm pixel không trắng → crop box → encode lại
```

---

## 4. PageSetup — Không cần Save & Restore

**Lý do:** File Excel đưa vào ArcTool là file **dữ liệu thô riêng biệt** (workflow VN).
Không ai dùng file này để in hồ sơ hay ký hợp đồng → tool tự do modify PageSetup.

**PageSetup set trước khi export (không restore):**

```csharp
ws.PageSetup.PrintArea      = targetRange.Address[false, false];
ws.PageSetup.Zoom           = false;          // BẮT BUỘC khi dùng FitToPages
ws.PageSetup.FitToPagesWide = 1;
ws.PageSetup.FitToPagesTall = 1;
ws.PageSetup.TopMargin      = 0;
ws.PageSetup.BottomMargin   = 0;
ws.PageSetup.LeftMargin     = 0;
ws.PageSetup.RightMargin    = 0;
ws.PageSetup.PaperSize      = XlPaperSize.xlPaperEsheet; // 864×1118mm ≈ A0
```

**Lý do chọn `xlPaperEsheet`:**
- ANSI E = 864×1118mm — khổ lớn nhất có sẵn trong `XlPaperSize` enum
- KHÔNG cần `xlPaperUser` + Windows API phức tạp
- FitToPages co/giãn nội dung vừa 1 trang → margin trắng thừa do SkiaSharp crop
- Bảng bao nhiêu dòng cũng chứa được

**Toàn bộ `XlPaperSize` enum theo kích thước (từ lớn đến nhỏ, phần liên quan):**

| Enum | Kích thước |
|---|---|
| `xlPaperEsheet` | 864 × 1118 mm ← **CHỌN** |
| `xlPaperDsheet` | 559 × 864 mm |
| `xlPaperCsheet` | 431 × 559 mm |
| `xlPaperA3` | 297 × 420 mm |
| `xlPaperB4` | 250 × 354 mm |

---

## 5. Code cũ — Giữ / Xóa / Sửa

### Giữ nguyên (không sửa):
- `OpenFile()`, `GetActiveSheetName()`, `GetSheetNames()`, `GetNamedRanges()`
- `Dispose()`, `ReleaseObject()`
- Range resolution logic trong `ExportRegion()`: NamedRange → PrintArea → UsedRange

### Xóa hoàn toàn:
- `ExportRangeInternal(Range range, string outputPath)` — toàn bộ chart/clipboard pipeline
- `const FIXED_SCALE_FACTOR = 35.0`
- `const MAX_EXCEL_DIMENSION = 32000`

### Sửa nhẹ:
- `ExportRegion()` — bỏ `_activeSheet` swap pattern (không cần nữa); range resolution giữ nguyên
- `ExportPrintAreaAsHighResImage()` — redirect sang `ExportRangeInternal()` mới

---

## 6. Temp file lifecycle

```
tempPdfPath = Path.Combine(Path.GetTempPath(), $"ArcTool_ExcelSync_{Guid.NewGuid():N}.pdf")
tempPngPath = Path.Combine(Path.GetTempPath(), $"ArcTool_ExcelSync_{Guid.NewGuid():N}.png")

// Cả hai đều delete trong finally block
// (tương tự pattern hiện tại của tempPng trong ExcelSyncEngine)
```

---

## 7. Tóm tắt thay đổi `.csproj`

```xml
<!-- THÊM -->
<PackageReference Include="PDFtoImage" Version="5.2.1" />

<!-- GIỮ NGUYÊN — đã có, không đổi -->
<PackageReference Include="SkiaSharp" Version="3.119.2" />
<PackageReference Include="Svg.Skia" Version="3.4.1" />
```

---

## 8. Version target

`ExcelInteropService.cs` → **V6.0**
Phạm vi thay đổi: chỉ file này, không đụng file nào khác.
