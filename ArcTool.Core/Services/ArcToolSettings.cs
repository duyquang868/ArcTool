using System;
using System.IO;
using System.Text.Json;
using System.Text.Json.Serialization;

namespace ArcTool.Core.Services
{
    /// <summary>
    /// Lưu trữ cài đặt người dùng giữa các lần chạy lệnh.
    /// File: %AppData%\ArcTool\settings.json
    /// </summary>
    public class ArcToolSettings
    {
        private static readonly string _settingsPath = Path.Combine(
            Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData),
            "ArcTool", "settings.json");

        [JsonPropertyName("lastScale")]
        public double LastScale { get; set; } = 100.0;

        /// <summary>Null nếu chưa từng lưu</summary>
        [JsonPropertyName("lastUsed")]
        public DateTime? LastUsed { get; set; }

        [JsonPropertyName("lastExcelFile")]
        public string LastExcelFile { get; set; } = string.Empty;

        /// <summary>
        /// Load từ disk. Trả về defaults nếu file không tồn tại hoặc bị corrupt.
        /// </summary>
        public static ArcToolSettings Load()
        {
            try
            {
                if (File.Exists(_settingsPath))
                {
                    string json = File.ReadAllText(_settingsPath);
                    return JsonSerializer.Deserialize<ArcToolSettings>(json) ?? new ArcToolSettings();
                }
            }
            catch { /* JSON corrupt hoặc quyền đọc bị từ chối → dùng defaults */ }

            return new ArcToolSettings();
        }

        /// <summary>
        /// Save xuống disk. Tự động cập nhật LastUsed = DateTime.Now.
        /// Không throw — save failure là non-critical.
        /// </summary>
        public void Save()
        {
            try
            {
                string dir = Path.GetDirectoryName(_settingsPath);
                if (dir != null) Directory.CreateDirectory(dir);

                LastUsed = DateTime.Now;

                string json = JsonSerializer.Serialize(this, new JsonSerializerOptions
                {
                    WriteIndented = true
                });
                File.WriteAllText(_settingsPath, json);
            }
            catch { /* Non-critical — tiếp tục bình thường dù save thất bại */ }
        }
    }
}
