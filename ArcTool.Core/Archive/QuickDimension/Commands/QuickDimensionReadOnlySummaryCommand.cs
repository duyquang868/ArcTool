#nullable enable
using System;
using System.Linq;
using System.Text;
using ArcTool.Core.Archive.QuickDimension.Models;
using ArcTool.Core.Archive.QuickDimension.Services;
using Autodesk.Revit.Attributes;
using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using Autodesk.Revit.UI.Selection;
using RevitTaskDialog = Autodesk.Revit.UI.TaskDialog;
using RevitView = Autodesk.Revit.DB.View;

namespace ArcTool.Core.Archive.QuickDimension.Commands
{
    /// <summary>
    /// Session 2.7 read-only smoke command for the wall-axis projection model: selects one host Wall,
    /// picks the placement side, and reports ordered wall-side anchors plus Door/Window jamb candidates
    /// without creating dimensions.
    /// </summary>
    [Transaction(TransactionMode.Manual)]
    [Regeneration(RegenerationOption.Manual)]
    public class QuickDimensionReadOnlySummaryCommand : IExternalCommand
    {
        private const int MaxCandidateLines = 60;
        private const int MaxDiagnosticLines = 24;

        public Result Execute(ExternalCommandData commandData, ref string message, ElementSet elements)
        {
            UIDocument? uidoc = commandData.Application.ActiveUIDocument;
            Document? doc = uidoc?.Document;

            if (uidoc == null || doc == null)
            {
                message = "No active document is available.";
                RevitTaskDialog.Show("ArcTool Error", message);
                return Result.Failed;
            }

            RevitView activeView = doc.ActiveView;
            if (!IsSupportedPlanView(activeView))
            {
                RevitTaskDialog.Show(
                    "ArcTool - Quick Dimension Read-Only Summary",
                    "Quick Dimension Session 2.7 supports active Plan Views only.\n\n" +
                    $"Current view type: {activeView.ViewType}");
                return Result.Cancelled;
            }

            try
            {
                Reference wallReference = uidoc.Selection.PickObject(
                    ObjectType.Element,
                    new WallAxisSelectionFilter(),
                    "Quick Dimension (wall-axis): select ONE host Wall (straight, non-curtain) as the dimension axis.");

                if (wallReference == null || doc.GetElement(wallReference.ElementId) is not Wall selectedWall || !selectedWall.IsValidObject)
                {
                    RevitTaskDialog.Show(
                        "ArcTool - Quick Dimension Read-Only Summary",
                        "The selected element is not a valid host Wall.");
                    return Result.Cancelled;
                }

                if (selectedWall.WallType?.Kind == WallKind.Curtain)
                {
                    RevitTaskDialog.Show(
                        "ArcTool - Quick Dimension Read-Only Summary",
                        "Curtain walls are outside the Quick Dimension MVP scope.");
                    return Result.Cancelled;
                }

                if (selectedWall.Location is not LocationCurve locationCurve || locationCurve.Curve is not Line wallLine)
                {
                    RevitTaskDialog.Show(
                        "ArcTool - Quick Dimension Read-Only Summary",
                        "The selected wall does not expose a straight (Line) LocationCurve. Arc/non-line host walls are excluded from MVP.");
                    return Result.Cancelled;
                }

                XYZ wallStart = wallLine.GetEndPoint(0);
                XYZ wallEnd = wallLine.GetEndPoint(1);

                XYZ sidePickPoint = uidoc.Selection.PickPoint(
                    ObjectSnapTypes.None,
                    "Quick Dimension (wall-axis): pick a point on the LEFT or RIGHT side of the wall to set the dimension placement side.");

                QuickDimensionOptions options = QuickDimensionOptions.Default;
                QuickDimensionLineContext lineContext = QuickDimensionLineContext.CreateFromWallAxis(
                    selectedWall.Id,
                    wallStart,
                    wallEnd,
                    sidePickPoint,
                    options.MinimumDimensionLineLength);

                if (lineContext.SideSign == 0)
                {
                    RevitTaskDialog.Show(
                        "ArcTool - Quick Dimension Read-Only Summary",
                        "The side pick point lies on the wall axis. Please pick a point clearly on the LEFT or RIGHT side of the wall.");
                    return Result.Cancelled;
                }

                QuickDimensionReadOnlyResult result = QuickDimensionReadOnlyEngine.CollectCandidates(
                    doc,
                    activeView,
                    lineContext,
                    options);

                // XML trace log is mandatory for audit, but must never mask the summary if writing fails
                // (for example, an unsaved document). Write it defensively and report the outcome inline.
                string xmlLogStatus = TryWriteXmlLog(doc, activeView, selectedWall, result);

                RevitTaskDialog.Show(
                    "ArcTool - Quick Dimension Read-Only Summary",
                    BuildSummaryMessage(result, activeView, selectedWall, xmlLogStatus));

                return Result.Succeeded;
            }
            catch (Autodesk.Revit.Exceptions.OperationCanceledException)
            {
                return Result.Cancelled;
            }
            catch (InvalidOperationException ex)
            {
                RevitTaskDialog.Show(
                    "ArcTool - Quick Dimension Read-Only Summary",
                    $"Quick Dimension could not build the wall axis:\n\n{ex.Message}");
                return Result.Cancelled;
            }
            catch (Exception ex)
            {
                message = ex.Message;
                RevitTaskDialog.Show(
                    "ArcTool Error",
                    $"Quick Dimension read-only summary failed.\n\n{ex.Message}");
                return Result.Failed;
            }
        }

        private static string TryWriteXmlLog(
            Document doc,
            RevitView activeView,
            Wall selectedWall,
            QuickDimensionReadOnlyResult result)
        {
            try
            {
                string logPath = QuickDimensionReadOnlyXmlLogService.WriteReadOnlySummaryLog(doc, activeView, selectedWall, result);
                return logPath;
            }
            catch (Exception ex)
            {
                return $"XML log failed: {ex.Message}";
            }
        }

        private sealed class WallAxisSelectionFilter : ISelectionFilter
        {
            public bool AllowElement(Element elem)
            {
                if (elem is not Wall wall || !wall.IsValidObject)
                {
                    return false;
                }

                if (wall.WallType?.Kind == WallKind.Curtain)
                {
                    return false;
                }

                return wall.Location is LocationCurve locationCurve && locationCurve.Curve is Line;
            }

            public bool AllowReference(Reference reference, XYZ position)
            {
                return false;
            }
        }

        private static string BuildSummaryMessage(QuickDimensionReadOnlyResult result, RevitView activeView, Wall selectedWall, string xmlLogPath)
        {
            StringBuilder sb = new StringBuilder();

            sb.AppendLine("=== QUICK DIMENSION READ-ONLY SUMMARY (WALL-AXIS PROJECTION MODEL) ===");
            sb.AppendLine();
            sb.AppendLine($"XML log: {xmlLogPath}");
            sb.AppendLine();
            sb.AppendLine($"View: {activeView.Name} ({activeView.ViewType})");
            sb.AppendLine($"Selected Wall: {BuildWallLabel(selectedWall)}");

            string sideLabel = result.LineContext.SideSign switch
            {
                1 => "Left (+CCW normal of axis direction)",
                -1 => "Right (-CCW normal of axis direction)",
                _ => "Unspecified"
            };
            sb.AppendLine($"Placement side: {sideLabel}");

            double wallAxisLengthMm = UnitUtils.ConvertFromInternalUnits(result.LineContext.Length, UnitTypeId.Millimeters);
            sb.AppendLine($"Wall-axis length: {wallAxisLengthMm:0.##} mm");
            sb.AppendLine($"Final candidate records: {result.CandidateCount}");
            sb.AppendLine($"Can create chain dimension later: {FormatBoolean(result.CanCreateChainDimension)}");
            sb.AppendLine();

            AppendSourceSummary(sb, result);
            AppendOrderedCandidates(sb, result);
            AppendDiagnostics(sb, result);

            return sb.ToString();
        }

        private static string BuildWallLabel(Wall wall)
        {
            string typeName = wall.WallType?.Name ?? string.Empty;
            return string.IsNullOrWhiteSpace(typeName)
                ? $"Wall {wall.Id.Value}"
                : $"Wall {wall.Id.Value} ({typeName})";
        }

        private static void AppendSourceSummary(StringBuilder sb, QuickDimensionReadOnlyResult result)
        {
            sb.AppendLine("--- FINAL CANDIDATE RECORDS ---");
            sb.AppendLine($"Grid: {result.GridCount}");
            sb.AppendLine($"Wall: {result.WallCount}");
            sb.AppendLine($"Door edge records: {result.DoorCount}");
            sb.AppendLine($"Window edge records: {result.WindowCount}");
            sb.AppendLine();

            sb.AppendLine("--- SOURCE ELEMENT SUMMARY ---");
            foreach (QuickDimensionSourceSummary summary in result.SourceSummaries)
            {
                sb.AppendLine(
                    $"{summary.SourceType}: collected {summary.CollectedCount}, " +
                    $"accepted elements {summary.AcceptedCount}, rejected {summary.RejectedCount}");
            }
            sb.AppendLine();
        }

        private static void AppendOrderedCandidates(StringBuilder sb, QuickDimensionReadOnlyResult result)
        {
            sb.AppendLine("--- ORDERED FINAL CANDIDATES ---");

            if (result.Candidates.Count == 0)
            {
                sb.AppendLine("No final candidates were accepted.");
                sb.AppendLine();
                return;
            }

            int index = 1;
            foreach (QuickDimensionCandidate candidate in result.Candidates.Take(MaxCandidateLines))
            {
                string hostLabel = candidate.HostElementValue.HasValue
                    ? $", Host {candidate.HostElementValue.Value}"
                    : string.Empty;

                double stationMm = UnitUtils.ConvertFromInternalUnits(candidate.ParameterOnDimensionLine, UnitTypeId.Millimeters);
                sb.AppendLine(
                    $"{index:00}. t={stationMm:0.##} mm | " +
                    $"{candidate.SourceType} | {candidate.DisplayName} | " +
                    $"{candidate.ReferenceStrategy} | Id {candidate.ElementValue}{hostLabel}");
                index++;
            }

            if (result.Candidates.Count > MaxCandidateLines)
            {
                sb.AppendLine($"... {result.Candidates.Count - MaxCandidateLines} more candidate records not shown.");
            }

            sb.AppendLine();
        }

        private static void AppendDiagnostics(StringBuilder sb, QuickDimensionReadOnlyResult result)
        {
            int infoCount = result.Diagnostics.Count(diagnostic => diagnostic.Severity == QuickDimensionDiagnosticSeverity.Info);
            int warningCount = result.Diagnostics.Count(diagnostic => diagnostic.Severity == QuickDimensionDiagnosticSeverity.Warning);
            int errorCount = result.Diagnostics.Count(diagnostic => diagnostic.Severity == QuickDimensionDiagnosticSeverity.Error);

            sb.AppendLine("--- DIAGNOSTICS ---");
            sb.AppendLine($"Total: {result.DiagnosticCount} | Info: {infoCount} | Warning: {warningCount} | Error: {errorCount}");

            var rejectionGroups = result.Diagnostics
                .Where(diagnostic => diagnostic.IsRejected)
                .GroupBy(diagnostic => diagnostic.Reason)
                .OrderByDescending(group => group.Count())
                .ThenBy(group => group.Key)
                .ToList();

            if (rejectionGroups.Count > 0)
            {
                sb.AppendLine("Rejected reason counts:");
                foreach (var group in rejectionGroups)
                {
                    sb.AppendLine($"  {group.Key}: {group.Count()}");
                }
            }

            var importantDiagnostics = result.Diagnostics
                .Where(diagnostic => diagnostic.Severity != QuickDimensionDiagnosticSeverity.Info || diagnostic.IsRejected)
                .Take(MaxDiagnosticLines)
                .ToList();

            if (importantDiagnostics.Count > 0)
            {
                sb.AppendLine();
                sb.AppendLine("Important diagnostics:");
                foreach (QuickDimensionDiagnostic diagnostic in importantDiagnostics)
                {
                    string elementLabel = diagnostic.ElementValue.HasValue
                        ? $" ElementId {diagnostic.ElementValue.Value}."
                        : string.Empty;

                    sb.AppendLine($"[{diagnostic.Severity}] {diagnostic.Reason}.{elementLabel} {diagnostic.Message}");
                }

                int remainingCount = result.Diagnostics
                    .Where(diagnostic => diagnostic.Severity != QuickDimensionDiagnosticSeverity.Info || diagnostic.IsRejected)
                    .Skip(MaxDiagnosticLines)
                    .Count();

                if (remainingCount > 0)
                {
                    sb.AppendLine($"... {remainingCount} more warning/error diagnostics not shown.");
                }
            }
            else
            {
                sb.AppendLine("No warning or error diagnostics.");
            }
        }

        private static string FormatBoolean(bool value)
        {
            return value ? "Yes" : "No";
        }

        private static bool IsSupportedPlanView(RevitView view)
        {
            if (view.IsTemplate)
            {
                return false;
            }

            return view.ViewType == ViewType.FloorPlan
                || view.ViewType == ViewType.CeilingPlan
                || view.ViewType == ViewType.EngineeringPlan
                || view.ViewType == ViewType.AreaPlan;
        }
    }
}
