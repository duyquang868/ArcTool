using System;
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
    /// Session 1.3 spike command: tests mixed Grid + Wall reference arrays for NewDimension.
    /// Validates that references from different source types can coexist in the same ReferenceArray.
    /// </summary>
    [Transaction(TransactionMode.Manual)]
    public class QuickDimensionMixedReferenceSpikeCommand : IExternalCommand
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
                    "ArcTool — Quick Dimension Mixed Spike",
                    "Quick Dimension Phase 1.3 only supports active Plan Views.");
                return Result.Cancelled;
            }

            try
            {
                XYZ firstPoint = uidoc.Selection.PickPoint(
                    ObjectSnapTypes.None,
                    "Quick Dimension mixed spike: pick the first point of the dimension line.");

                XYZ secondPoint = uidoc.Selection.PickPoint(
                    ObjectSnapTypes.None,
                    "Quick Dimension mixed spike: pick the second point of the dimension line.");

                QuickDimensionMixedProbeSummary summary = QuickDimensionMixedReferenceProbeService.RunMixedReferenceProbe(
                    doc,
                    activeView,
                    firstPoint,
                    secondPoint);

                RevitTaskDialog.Show(
                    "ArcTool — Quick Dimension Mixed Spike",
                    BuildSummaryMessage(summary));

                return summary.MixedReferencesWork ? Result.Succeeded : Result.Failed;
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
                    $"Quick Dimension mixed reference spike failed.\n\n{ex.Message}");
                return Result.Failed;
            }
        }

        private static string BuildSummaryMessage(QuickDimensionMixedProbeSummary summary)
        {
            StringBuilder sb = new StringBuilder();

            sb.AppendLine("Session 1.3 mixed reference probe complete.");
            sb.AppendLine();

            sb.AppendLine("=== COLLECTION STATS ===");
            sb.AppendLine($"Grids collected: {summary.CollectedGridCount}");
            sb.AppendLine($"Grids accepted: {summary.AcceptedGridCount}");
            sb.AppendLine($"Grids skipped (arc): {summary.SkippedArcGridCount}");
            sb.AppendLine($"Grids skipped (parallel): {summary.SkippedParallelGridCount}");
            sb.AppendLine();
            sb.AppendLine($"Walls collected: {summary.CollectedWallCount}");
            sb.AppendLine($"Walls accepted: {summary.AcceptedWallCount}");
            sb.AppendLine($"Walls skipped (curtain): {summary.SkippedCurtainWallCount}");
            sb.AppendLine($"Walls skipped (parallel): {summary.SkippedParallelWallCount}");
            sb.AppendLine($"Walls skipped (no face ref): {summary.SkippedNoFaceReferenceCount}");
            sb.AppendLine();
            sb.AppendLine($"Total accepted candidates: {summary.TotalAcceptedCount}");
            sb.AppendLine();

            sb.AppendLine("=== SCENARIO RESULTS ===");
            sb.AppendLine();
            sb.AppendLine(FormatScenario(summary.SortedResult));
            sb.AppendLine();
            sb.AppendLine(FormatScenario(summary.ReversedResult));
            sb.AppendLine();
            sb.AppendLine(FormatScenario(summary.GridsOnlyResult));
            sb.AppendLine();
            sb.AppendLine(FormatScenario(summary.WallsOnlyResult));

            sb.AppendLine();
            sb.AppendLine("=== CONCLUSION ===");
            if (summary.MixedReferencesWork)
            {
                sb.AppendLine("PASS: Mixed Grid + Wall references work in the same ReferenceArray.");
            }
            else
            {
                sb.AppendLine("FAIL: Mixed Grid + Wall references do NOT work together.");
            }

            return sb.ToString();
        }

        private static string FormatScenario(QuickDimensionMixedScenarioResult result)
        {
            if (result == null)
            {
                return "Scenario: N/A";
            }

            string status = result.Succeeded ? "PASS" : "FAIL";
            return $"{result.Scenario}: {status}\n" +
                   $"  References: {result.TotalReferenceCount} ({result.GridReferenceCount} grids, {result.WallReferenceCount} walls)\n" +
                   $"  Result: {result.Message}";
        }
    }
}
