using System;
using System.Collections.Generic;
using System.IO;
using System.Text.Json;
using System.Text.Json.Serialization;
using ArcTool.Core.Models;
using Autodesk.Revit.DB;

namespace ArcTool.Core.Services
{
    /// <summary>
    /// Quản lý load/save danh sách ExcelMapping dưới dạng JSON.
    ///
    /// File JSON lưu tại: [thư mục chứa .rvt]\ArcTool_ExcelSync.json
    /// Chiến lược ghi: ATOMIC WRITE — ghi vào .tmp trước, sau đó Replace/Move.
    /// Đảm bảo JSON không bao giờ ở trạng thái corrupt nếu Revit crash giữa chừng.
    ///
    /// V1.0 — Phase 1B, phục vụ ExcelToRevit V3.0
    /// Phụ thuộc: ArcTool.Core.Models.ExcelMapping (Phase 1A), Autodesk.Revit.DB.Document
    /// </summary>
    public static class ArcToolSettingsService
    {
        // ── CONSTANTS ─────────────────────────────────────────────────────────

        private const string JsonFileName = "ArcTool_ExcelSync.json";

        // Cache JsonSerializerOptions — không allocate mới mỗi lần call.
        // WriteIndented = true để JSON dễ đọc/debug trực tiếp bằng text editor.
        // JsonStringEnumConverter: serialize enum thành string ("DraftingView")
        //   thay vì số (0) — forward-compatible khi thêm enum value mới.
        // NOTE: Nếu thêm enum value mới vào ExcelViewType/ExcelRegionType và file JSON
        //   cũ chứa string đó → DeserializeException. Cần migration strategy khi đó.
        private static readonly JsonSerializerOptions SerializerOptions = new JsonSerializerOptions
        {
            WriteIndented = true,
            PropertyNameCaseInsensitive = true, // Tolerate "ViewName" vs "viewName" khi đọc JSON cũ
            Converters = { new JsonStringEnumConverter() }
        };

        // ── PUBLIC API ────────────────────────────────────────────────────────

        /// <summary>
        /// Trả về đường dẫn tuyệt đối của file JSON settings.
        /// </summary>
        /// <exception cref="InvalidOperationException">
        /// Khi doc.PathName rỗng — file Revit chưa được lưu lần nào.
        /// Caller phải bắt exception này và hiện dialog yêu cầu user lưu file trước.
        /// </exception>
        public static string GetSettingsPath(Document doc)
        {
            if (string.IsNullOrWhiteSpace(doc?.PathName))
                throw new InvalidOperationException(
                    "File Revit chưa được lưu.\n\n" +
                    "Vui lòng lưu file (.rvt) trước khi sử dụng tính năng Excel to Revit.\n" +
                    "File JSON settings sẽ được tạo tại cùng thư mục với file Revit.");

            string rvtDir = Path.GetDirectoryName(doc.PathName)
                            ?? throw new InvalidOperationException(
                                $"Không thể xác định thư mục chứa file Revit: '{doc.PathName}'");

            return Path.Combine(rvtDir, JsonFileName);
        }

        /// <summary>
        /// Đọc danh sách ExcelMapping từ JSON settings.
        /// Trả về List rỗng nếu file chưa tồn tại hoặc bị corrupt — không crash.
        /// </summary>
        /// <exception cref="InvalidOperationException">
        /// Khi doc.PathName rỗng (propagate từ GetSettingsPath).
        /// </exception>
        public static List<ExcelMapping> LoadMappings(Document doc)
        {
            string path = GetSettingsPath(doc); // Throw nếu doc.PathName rỗng

            // File chưa tồn tại = lần đầu dùng tính năng → trả về List rỗng bình thường
            if (!File.Exists(path))
                return new List<ExcelMapping>();

            try
            {
                string json = File.ReadAllText(path, System.Text.Encoding.UTF8);

                // Deserialize → List<ExcelMapping>
                // Trả về null nếu JSON là "null" literal → dùng ?? để fallback
                List<ExcelMapping> result = JsonSerializer.Deserialize<List<ExcelMapping>>(json, SerializerOptions)
                                           ?? new List<ExcelMapping>();

                return result;
            }
            catch (JsonException jsonEx)
            {
                // File JSON bị corrupt (schema lỗi, truncated, parse error)
                // Log để developer debug, nhưng không crash — user tiếp tục được dùng tool
                System.Diagnostics.Debug.WriteLine(
                    $"[ArcToolSettingsService] JSON corrupt tại '{path}': {jsonEx.Message}");

                // Backup file lỗi để debug sau nếu cần
                TryBackupCorruptFile(path);

                return new List<ExcelMapping>();
            }
            catch (Exception ex)
            {
                // IOException (file đang bị lock), UnauthorizedAccessException, v.v.
                System.Diagnostics.Debug.WriteLine(
                    $"[ArcToolSettingsService] Lỗi đọc JSON '{path}': {ex.Message}");

                return new List<ExcelMapping>();
            }
        }

        /// <summary>
        /// Ghi danh sách ExcelMapping ra JSON với chiến lược ATOMIC WRITE.
        ///
        /// Quy trình:
        ///   1. Serialize → string
        ///   2. Ghi vào [JsonFileName].tmp (cùng thư mục)
        ///   3. File.Replace() nếu file đích đã tồn tại (atomic trên cùng volume)
        ///      hoặc File.Move() nếu file đích chưa tồn tại (lần đầu)
        ///
        /// Kết quả: nếu crash ở bước 2 → .tmp bị corrupt, JSON gốc vẫn nguyên vẹn.
        /// </summary>
        /// <exception cref="InvalidOperationException">
        /// Khi doc.PathName rỗng (propagate từ GetSettingsPath).
        /// </exception>
        /// <exception cref="IOException">
        /// Khi không thể ghi file do quyền truy cập hoặc disk đầy.
        /// Caller nên bắt và hiện dialog lỗi cho user.
        /// </exception>
        public static void SaveMappings(Document doc, List<ExcelMapping> mappings)
        {
            if (mappings == null) throw new ArgumentNullException(nameof(mappings));

            string finalPath = GetSettingsPath(doc); // Throw nếu doc.PathName rỗng
            string tempPath = finalPath + ".tmp";

            // Bước 1: Serialize
            string json = JsonSerializer.Serialize(mappings, SerializerOptions);

            // Bước 2: Ghi ra .tmp (nếu crash ở đây, JSON gốc vẫn nguyên)
            File.WriteAllText(tempPath, json, System.Text.Encoding.UTF8);

            // Bước 3: Atomic replace
            if (File.Exists(finalPath))
            {
                // File.Replace(source, destination, backup):
                //   - source      = .tmp (file mới vừa ghi)
                //   - destination = .json (file cũ cần replace)
                //   - backup      = null  (không giữ backup — JSON cũ bị ghi đè)
                // Trên cùng volume: Replace là atomic ở tầng filesystem NTFS.
                File.Replace(tempPath, finalPath, destinationBackupFileName: null);
            }
            else
            {
                // Lần đầu chưa có file JSON → Move thay vì Replace
                // File.Move là atomic rename trên cùng volume (NTFS).
                File.Move(tempPath, finalPath);
            }
        }

        // ── PER-ROW STATUS HELPERS ────────────────────────────────────────────

        /// <summary>
        /// Kiểm tra file Excel của mapping có tồn tại không.
        /// Dùng cho: Status Dot màu vàng khi file bị move/xóa.
        /// </summary>
        public static bool FileExists(ExcelMapping mapping)
        {
            if (mapping == null || string.IsNullOrWhiteSpace(mapping.FilePath))
                return false;

            return File.Exists(mapping.FilePath);
        }

        /// <summary>
        /// Kiểm tra file Excel có được sửa sau lần sync cuối không.
        /// Trả về false (thay vì throw) nếu file không tồn tại — caller dùng FileExists() riêng.
        ///
        /// Logic: File.GetLastWriteTime(path) > mapping.LastModified
        /// NOTE: LastModified mặc định = DateTime.MinValue → mọi file đều "changed" lần đầu.
        ///       Đây là behavior đúng: buộc user phải nhấn Update ít nhất 1 lần.
        /// </summary>
        public static bool HasFileChanged(ExcelMapping mapping)
        {
            if (mapping == null || string.IsNullOrWhiteSpace(mapping.FilePath))
                return false;

            if (!File.Exists(mapping.FilePath))
                return false; // File mất → xử lý qua FileExists(), không phải HasFileChanged()

            try
            {
                DateTime fileLastWrite = File.GetLastWriteTime(mapping.FilePath);

                // So sánh theo local time (GetLastWriteTime trả về local).
                // mapping.LastModified cũng phải là local time khi được gán.
                // ExcelSyncEngine sẽ dùng DateTime.Now (không phải UtcNow) khi lưu LastModified.
                return fileLastWrite > mapping.LastModified;
            }
            catch (Exception ex)
            {
                // IOException nếu file đang bị lock hoặc network path mất
                System.Diagnostics.Debug.WriteLine(
                    $"[ArcToolSettingsService.HasFileChanged] Lỗi đọc timestamp '{mapping.FilePath}': {ex.Message}");

                return false; // Không chắc → không trigger update để tránh false positive
            }
        }

        // ── PRIVATE HELPERS ───────────────────────────────────────────────────

        /// <summary>
        /// Đổi tên file JSON corrupt thành .corrupt để debug sau.
        /// Không throw — đây là best-effort, không nghiêm trọng nếu thất bại.
        /// </summary>
        private static void TryBackupCorruptFile(string corruptPath)
        {
            try
            {
                string backupPath = corruptPath + $".corrupt_{DateTime.Now:yyyyMMdd_HHmmss}";

                // Chỉ backup nếu chưa có quá nhiều file corrupt (tránh disk đầy)
                string dir = Path.GetDirectoryName(corruptPath) ?? ".";
                string prefix = Path.GetFileName(corruptPath) + ".corrupt_";
                int count = Directory.GetFiles(dir, prefix + "*").Length;

                if (count < 5) // Giữ tối đa 5 bản backup corrupt
                    File.Copy(corruptPath, backupPath, overwrite: false);
            }
            catch
            {
                // Backup thất bại → bỏ qua, không nghiêm trọng
            }
        }
    }
}