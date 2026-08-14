using System;
using System.Text;
using ArcTool.Core.Archive.QuickDimension.Models;
using ArcTool.Core.Archive.QuickDimension.Services;
using Autodesk.Revit.Attributes;
using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using RevitTaskDialog = Autodesk.Revit.UI.TaskDialog;
using RevitView = Autodesk.Revit.DB.View;

namespace ArcTool.Core.Archive.QuickDimension.Commands
{
    /// <summary>
    /// Session 1.4 spike command: tests Door/Window opening reference strategies for NewDimension.
    /// Tests three strategies:
    /// 1. FamilyInstance.GetReferences(FamilyInstanceReferenceType.Left/Right)
    /// 2. Options.ComputeReferences = true + Face.Reference
    /// 3. Host Wall FindInserts + Opening Geometry
    /// </summary>
    [Transaction(TransactionMode.Manual)]
    public class QuickDimensionDoorWindowReferenceSpikeCommand : IExternalCommand
    {
        public Result Execute(ExternalCommandData commandData, ref string message, ElementSet elements)
        {
            UIDocument uidoc = commandData.Application.ActiveUIDocument;
            Document doc = uidoc?.Document;

            if (doc == null)
            {
                message = "No active document is available.";
                RevitTaskDialog.Show("ArcTool Error", message);
                return Result.Failed;
            }

            RevitView activeView = doc.ActiveView;
            if (activeView is not ViewPlan)
            {
                RevitTaskDialog.Show(
                    "ArcTool — Quick Dimension Door/Window Spike",
                    "Quick Dimension Phase 1.4 only supports active Plan Views.");
                return Result.Cancelled;
            }

            try
            {
                XYZ firstPoint = uidoc.Selection.PickPoint(
                    Autodesk.Revit.UI.Selection.ObjectSnapTypes.None,
                    "Quick Dimension Door/Window spike: pick the first point of the dimension line.");

                XYZ secondPoint = uidoc.Selection.PickPoint(
                    Autodesk.Revit.UI.Selection.ObjectSnapTypes.None,
                    "Quick Dimension Door/Window spike: pick the second point of the dimension line.");

                QuickDimensionDoorWindowProbeSummary summary = QuickDimensionDoorWindowReferenceProbeService.RunDoorWindowReferenceProbe(
                    doc,
                    activeView,
                    firstPoint,
                    secondPoint);

                RevitTaskDialog.Show(
                    "ArcTool — Quick Dimension Door/Window Spike",
                    BuildSummaryMessage(summary));

                return summary.AnyStrategySucceeded ? Result.Succeeded : Result.Failed;
            }
            catch (Autodesk.Revit.Exceptions.OperationCanceledException)
            {
                return Result.Cancelled;
            }
            catch (Exception ex)
            {
                message = ex.Message;
                RevitTaskDialog.Show(
                    "ArcTool Error",
                    $"Quick Dimension Door/Window reference spike failed.\n\n{ex.Message}");
                return Result.Failed;
            }
        }

        private static string BuildSummaryMessage(QuickDimensionDoorWindowProbeSummary summary)
        {
            StringBuilder sb = new StringBuilder();

            sb.AppendLine("Session 1.4 Door/Window reference probe complete.");
            sb.AppendLine();

            sb.AppendLine("=== COLLECTION STATS ===");
            sb.AppendLine($"Doors collected: {summary.CollectedDoorCount}");
            sb.AppendLine($"Doors accepted: {summary.AcceptedDoorCount}");
            sb.AppendLine($"Windows collected: {summary.CollectedWindowCount}");
            sb.AppendLine($"Windows accepted: {summary.AcceptedWindowCount}");
            sb.AppendLine();
            sb.AppendLine($"Skipped (non-hosted): {summary.SkippedNonHostedCount}");
            sb.AppendLine($"Skipped (parallel to dim line): {summary.SkippedParallelCount}");
            sb.AppendLine($"Skipped (outside span): {summary.SkippedOutsideSpanCount}");
            sb.AppendLine();
            sb.AppendLine($"Total accepted candidates: {summary.TotalAccepted}");
            sb.AppendLine();

            sb.AppendLine("=== STRATEGY RESULTS ===");
            sb.AppendLine();
            sb.AppendLine(FormatStrategyResult(summary.FamilyInstanceResult));
            sb.AppendLine();
            sb.AppendLine(FormatStrategyResult(summary.GeometryResult));
            sb.AppendLine();
            sb.AppendLine(FormatStrategyResult(summary.OpeningGeometryResult));

            sb.AppendLine();
            sb.AppendLine("=== CONCLUSION ===");
            if (summary.AnyStrategySucceeded)
            {
                sb.AppendLine($"PASS: At least one strategy works for Door/Window dimensions.");
                sb.AppendLine($"Best strategy: {summary.BestStrategy}");
            }
            else
            {
                sb.AppendLine("FAIL: No strategy succeeded for Door/Window dimensions.");
                sb.AppendLine("Check if families have proper reference planes or if geometry extraction is working.");
            }

            return sb.ToString();
        }

        private static string FormatStrategyResult(QuickDimensionDoorWindowStrategyResult result)
        {
            if (result == null)
            {
                return "Strategy: N/A";
            }

            string status = result.Succeeded ? "PASS" : "FAIL";
            return $"[{status}] {result.Strategy}\n" +
                   $"  Total candidates: {result.TotalCandidates}\n" +
                   $"  Candidates with refs: {result.CandidatesWithReferences}\n" +
                   $"  References used: {result.ReferencesUsed}\n" +
                   $"  Result: {result.Message}";
        }
    }
}
