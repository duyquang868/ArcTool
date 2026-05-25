using System;
using System.IO;
using System.Text;
using ArcTool.Core.Models;
using Autodesk.Revit.DB;

namespace ArcTool.Core.Services
{
    /// <summary>
    /// Writes best-effort append-only coordinate diagnostics next to the active RVT file.
    /// Logging failures are intentionally swallowed so support infrastructure never interrupts Revit workflows.
    /// </summary>
    public static class CoordinateLogService
    {
        private const string Settings = "SETTINGS";
        private const string Registration = "REGISTRATION";
        private const string Batch = "BATCH";
        private const string Toggle = "TOGGLE";
        private const string Error = "ERROR";
        private const string LogFileName = "ArcTool_Coord.log";

        private static readonly object _sync = new object();

        /// <summary>
        /// Append one log line. Category is one of the defined CATEGORIES constants.
        /// Never throws. File I/O errors are silently swallowed.
        /// Thread-safety: uses lock(_sync) to prevent interleaved writes on projects where future multi-thread scenarios might apply.
        /// </summary>
        /// <param name="doc">Revit document whose path determines the primary log location.</param>
        /// <param name="category">Log category label written into the line prefix.</param>
        /// <param name="message">Diagnostic message written after the category prefix.</param>
        public static void Log(Document doc, string category, string message)
        {
            try
            {
                string safeCategory = string.IsNullOrWhiteSpace(category) ? Error : category;
                string safeMessage = message ?? string.Empty;
                string line = $"[{DateTime.Now:yyyy-MM-dd HH:mm:ss}] [{safeCategory}] {safeMessage}";

                string path = GetLogPath(doc);
                string fallbackPath = GetFallbackLogPath();
                bool usingPrimaryPath = !string.Equals(path, fallbackPath, StringComparison.OrdinalIgnoreCase);

                if (TryAppendLine(path, line))
                {
                    return;
                }

                if (usingPrimaryPath)
                {
                    TryAppendLine(fallbackPath, line);
                }
            }
            catch
            {
            }
        }

        /// <summary>
        /// Convenience: log a batch run summary using counts from CoordBatchSummary.
        /// Logs one line: BATCH category with total/written/skipped/unsupported/failed counts.
        /// </summary>
        /// <param name="doc">Revit document whose path determines the primary log location.</param>
        /// <param name="summary">Batch summary returned by the coordinate batch service.</param>
        public static void LogBatch(Document doc, CoordBatchSummary summary)
        {
            try
            {
                if (summary == null)
                {
                    Log(doc, Batch, "Summary=null");
                    return;
                }

                string msg = $"Total={summary.TotalCollected} " +
                             $"Written={summary.WrittenCount} " +
                             $"Skipped={summary.SkippedCount} " +
                             $"Unsupported={summary.UnsupportedCount} " +
                             $"Failed={summary.FailedCount}";
                Log(doc, Batch, msg);
            }
            catch
            {
            }
        }

        /// <summary>
        /// Convenience: log a settings-change event with before/after values.
        /// </summary>
        /// <param name="doc">Revit document whose path determines the primary log location.</param>
        /// <param name="axisMappingBefore">Axis mapping key before the dialog was accepted.</param>
        /// <param name="axisMappingAfter">Axis mapping key after the dialog was accepted.</param>
        /// <param name="outputUnitBefore">Output unit key before the dialog was accepted.</param>
        /// <param name="outputUnitAfter">Output unit key after the dialog was accepted.</param>
        public static void LogSettingsChange(
            Document doc,
            string axisMappingBefore, string axisMappingAfter,
            string outputUnitBefore, string outputUnitAfter)
        {
            try
            {
                string msg = $"AxisMapping: {axisMappingBefore}→{axisMappingAfter}  " +
                             $"Unit: {outputUnitBefore}→{outputUnitAfter}";
                Log(doc, Settings, msg);
            }
            catch
            {
            }
        }

        /// <summary>
        /// Convenience: log an updater toggle event.
        /// </summary>
        /// <param name="doc">Revit document whose path determines the primary log location.</param>
        /// <param name="isNowEnabled">True when auto-update is now enabled; false when it is now disabled.</param>
        public static void LogToggle(Document doc, bool isNowEnabled)
        {
            try
            {
                Log(doc, Toggle, isNowEnabled ? "Auto-Update ENABLED" : "Auto-Update DISABLED");
            }
            catch
            {
            }
        }

        private static bool TryAppendLine(string path, string line)
        {
            try
            {
                lock (_sync)
                {
                    File.AppendAllText(path, line + Environment.NewLine, Encoding.UTF8);
                }

                return true;
            }
            catch
            {
                return false;
            }
        }

        private static string GetLogPath(Document doc)
        {
            try
            {
                if (!string.IsNullOrWhiteSpace(doc?.PathName))
                {
                    string directory = Path.GetDirectoryName(doc.PathName);
                    if (!string.IsNullOrWhiteSpace(directory))
                    {
                        return Path.Combine(directory, LogFileName);
                    }
                }
            }
            catch
            {
            }

            return GetFallbackLogPath();
        }

        private static string GetFallbackLogPath()
        {
            try
            {
                return Path.Combine(Path.GetTempPath(), LogFileName);
            }
            catch
            {
                return LogFileName;
            }
        }
    }
}
