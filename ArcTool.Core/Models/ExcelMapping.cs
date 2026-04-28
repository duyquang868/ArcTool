using System;
using System.Text.Json.Serialization;

namespace ArcTool.Core.Models
{
    // ──────────────────────────────────────────────────────────────────────────
    //  ENUMS
    //  Đặt trong cùng file với ExcelMapping để tránh phải import thêm namespace.
    //
    //  CẢNH BÁO ĐẶT TÊN:
    //  - "ExcelViewType" thay vì "ViewType" để tránh xung đột với
    //    Autodesk.Revit.DB.ViewType ở mọi file import cả hai namespace.
    //  - "ExcelRegionType" đồng nhất prefix để dễ tìm kiếm trong codebase.
    // ──────────────────────────────────────────────────────────────────────────

    /// <summary>
    /// Loại vùng dữ liệu cần export từ sheet Excel.
    /// Thứ tự ưu tiên khi export: NamedRange → PrintArea → UsedRange.
    /// </summary>
    public enum ExcelRegionType
    {
        /// <summary>
        /// Dùng Named Range cụ thể (trường Region phải có giá trị).
        /// Ưu tiên cao nhất — cho phép export từng vùng riêng biệt trong cùng 1 sheet.
        /// </summary>
        NamedRange = 0,

        /// <summary>
        /// Dùng Print Area đã thiết lập trong Page Setup của sheet.
        /// Fallback khi không chọn Named Range.
        /// </summary>
        PrintArea = 1,

        /// <summary>
        /// Dùng toàn bộ vùng có dữ liệu (UsedRange).
        /// Fallback cuối cùng — kết quả có thể rộng hơn mong muốn.
        /// </summary>
        UsedRange = 2
    }

    /// <summary>
    /// Loại Revit View sẽ được tạo/ghi đè khi import ảnh.
    /// Đặt tên "ExcelViewType" để tránh xung đột với Autodesk.Revit.DB.ViewType.
    /// </summary>
    public enum ExcelViewType
    {
        /// <summary>
        /// Drafting View — Revit API có ViewDrafting.Create(), stable.
        /// Khuyến nghị cho bảng biểu, sơ đồ không cần link với model.
        /// </summary>
        DraftingView = 0,

        /// <summary>
        /// Legend View — Revit API 2026 KHÔNG có Create().
        /// Workaround: Duplicate từ legend template rỗng tên "ArcTool_LegendTemplate".
        /// Yêu cầu: project phải có sẵn ít nhất 1 Legend View.
        /// </summary>
        LegendView = 1
    }

    // ──────────────────────────────────────────────────────────────────────────
    //  MAIN MODEL
    // ──────────────────────────────────────────────────────────────────────────

    /// <summary>
    /// Ánh xạ giữa một vùng dữ liệu Excel và một View trong Revit.
    /// Được serialize/deserialize thành JSON, lưu tại cùng thư mục với file .rvt.
    /// File JSON: ArcTool_ExcelSync.json
    ///
    /// LIFECYCLE:
    ///   Tạo mới → User chọn file/sheet/region → Nhấn Update → LastModified cập nhật.
    ///   Mở dialog lần sau → HasChanges = (file.LastWriteTime > LastModified).
    ///   AutoSync = true → tự động update khi dialog mở nếu HasChanges.
    /// </summary>
    public class ExcelMapping
    {
        // ── IDENTITY ──────────────────────────────────────────────────────────

        /// <summary>
        /// Unique identifier của mapping. Sinh bằng Guid.NewGuid().ToString()
        /// khi tạo dòng mới. Không bao giờ thay đổi sau khi tạo.
        /// Dùng để track mapping kể cả khi ViewName hoặc FilePath bị đổi.
        /// </summary>
        [JsonPropertyName("id")]
        public string Id { get; set; } = Guid.NewGuid().ToString();

        // ── REVIT TARGET ──────────────────────────────────────────────────────

        /// <summary>
        /// Tên View sẽ được tạo/ghi đè trong Revit.
        /// Convention:
        ///   - RegionType = PrintArea/UsedRange → ViewName = SheetName
        ///   - RegionType = NamedRange          → ViewName = "SheetName_RegionName"
        /// Tự động sinh khi user chọn WorkSheet và Region trong dialog.
        /// </summary>
        [JsonPropertyName("viewName")]
        public string ViewName { get; set; } = string.Empty;

        /// <summary>
        /// Loại View Revit sẽ tạo. Mặc định DraftingView vì không cần template.
        /// Nếu chọn LegendView, project phải có sẵn 1 Legend View rỗng làm template.
        /// </summary>
        [JsonPropertyName("viewType")]
        public ExcelViewType ViewType { get; set; } = ExcelViewType.DraftingView;

        // ── EXCEL SOURCE ──────────────────────────────────────────────────────

        /// <summary>
        /// Đường dẫn tuyệt đối đến file Excel (.xlsx hoặc .xls).
        /// Nếu file bị di chuyển, Status Dot chuyển vàng và user phải chọn lại.
        /// </summary>
        [JsonPropertyName("filePath")]
        public string FilePath { get; set; } = string.Empty;

        /// <summary>
        /// Tên sheet trong file Excel (phân biệt hoa/thường theo Excel).
        /// </summary>
        [JsonPropertyName("workSheet")]
        public string WorkSheet { get; set; } = string.Empty;

        /// <summary>
        /// Tên Named Range cần export. null = dùng PrintArea hoặc UsedRange.
        /// Chỉ có giá trị khi RegionType = NamedRange.
        /// KHÔNG dùng string.Empty thay cho null — cần phân biệt "chưa chọn" vs "chọn rồi".
        /// </summary>
        [JsonPropertyName("region")]
        public string Region { get; set; } = null;

        /// <summary>
        /// Chiến lược resolve vùng export. Xác định thứ tự ưu tiên.
        /// </summary>
        [JsonPropertyName("regionType")]
        public ExcelRegionType RegionType { get; set; } = ExcelRegionType.PrintArea;

        // ── SYNC CONTROL ──────────────────────────────────────────────────────

        /// <summary>
        /// true = tự động update khi dialog mở và phát hiện file Excel thay đổi.
        /// false = chỉ update khi user nhấn nút Update thủ công.
        /// Khi AutoSync = true, nút Update per-row bị disabled trong UI.
        /// </summary>
        [JsonPropertyName("autoSync")]
        public bool AutoSync { get; set; } = false;

        /// <summary>
        /// Thời điểm update thành công lần cuối (local time của máy tính).
        /// Mặc định DateTime.MinValue → khi dialog mở lần đầu, mọi file đều được
        /// coi là "có thay đổi chưa sync" (HasChanges = true) — behavior đúng và mong muốn.
        /// Cập nhật sau mỗi lần ExecuteUpdate() thành công.
        /// </summary>
        [JsonPropertyName("lastModified")]
        public DateTime LastModified { get; set; } = DateTime.MinValue;

        // ── REVIT IMAGE STATE ─────────────────────────────────────────────────

        /// <summary>
        /// ElementId.Value của ImageInstance hiện tại trong Revit.
        /// Sentinel value = 0 nghĩa là "chưa import lần nào" (tương đương ElementId.InvalidElementId).
        /// Dùng long thay vì int vì ElementId.Value là long kể từ Revit 2024.
        /// </summary>
        [JsonPropertyName("imageInstanceId")]
        public long ImageInstanceId { get; set; } = 0;

        /// <summary>
        /// Chiều rộng (Width) của ImageInstance tính bằng feet (đơn vị nội bộ Revit).
        /// Sentinel value = 0 → chưa có kích thước; ExcelSyncEngine phải guard:
        ///   if (StoredWidth > 0) newInst.Width = StoredWidth;
        /// Được cập nhật sau mỗi lần import/update thành công.
        /// Phản ánh kích thước user đã resize trực tiếp trong Revit (Smart Scale).
        /// </summary>
        [JsonPropertyName("storedWidth")]
        public double StoredWidth { get; set; } = 0.0;

        /// <summary>
        /// Chiều cao (Height) của ImageInstance tính bằng feet (đơn vị nội bộ Revit).
        /// Sentinel value = 0 → cùng logic với StoredWidth.
        /// </summary>
        [JsonPropertyName("storedHeight")]
        public double StoredHeight { get; set; } = 0.0;

        // ── COMPUTED HELPERS (không serialize) ────────────────────────────────

        /// <summary>
        /// true nếu chưa import lần nào (ImageInstanceId chưa được gán).
        /// Dùng trong UI để quyết định label nút Update: "Import" vs "Update".
        /// </summary>
        [JsonIgnore]
        public bool IsFirstImport => ImageInstanceId == 0;

        /// <summary>
        /// true nếu cả StoredWidth lẫn StoredHeight đều có giá trị hợp lệ.
        /// Guard cho ExcelSyncEngine trước khi áp lại kích thước.
        /// </summary>
        [JsonIgnore]
        public bool HasStoredDimensions => StoredWidth > 0.0 && StoredHeight > 0.0;

        /// <summary>
        /// Sinh ViewName tự động từ WorkSheet và Region theo convention đã chốt.
        /// Gọi method này mỗi khi user thay đổi WorkSheet hoặc Region trong dialog.
        /// Không tự gọi trong constructor để tránh side effect khi deserialize JSON.
        /// </summary>
        public string BuildViewName()
        {
            if (string.IsNullOrWhiteSpace(WorkSheet))
                return string.Empty;

            // Named Range → "SheetName_RegionName"
            if (RegionType == ExcelRegionType.NamedRange && !string.IsNullOrWhiteSpace(Region))
                return $"{WorkSheet}_{Region}";

            // PrintArea hoặc UsedRange → "SheetName"
            return WorkSheet;
        }
    }
}
