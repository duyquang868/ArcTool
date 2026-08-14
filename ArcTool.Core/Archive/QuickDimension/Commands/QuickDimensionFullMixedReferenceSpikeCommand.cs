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
    /// Session 1.5 spike command: tests full mixed Grid + Wall + Door + Window reference arrays.
    /// This is the final Phase 1 reference feasibility spike.
    /// </summary>
    [Transaction(TransactionMode.Manual)]
    public class QuickDimensionFullMixedReferenceSpikeCommand : IExternalCommand
    {
        public Result Execute(ExternalCommandData commandData, ref string message, ElementSet elements)
        {
            UIDocument uidoc = commandData.Application.ActiveUIDocument;
            Document doc = uidoc.Document;
            RevitView activeView = doc.ActiveView;

            if (activeView.ViewType != ViewType.FloorPlan &&
                activeView.ViewType != ViewType.CeilingPlan &&
                activeView.ViewType != ViewType.EngineeringPlan &&
                activeView.ViewType != ViewType.AreaPlan)
            {
                RevitTaskDialog.Show("ArcTool",
                    "Quick Dimension Full Mixed Spike requires an active Plan View.\n\n" +
                    $"Current view type: {activeView.ViewType}");
                return Result.Cancelled;
            }

            try
            {
                XYZ firstPoint = uidoc.Selection.PickPoint(
                    ObjectSnapTypes.Endpoints | ObjectSnapTypes.Midpoints | ObjectSnapTypes.Intersections,
                    "Pick the FIRST point of the dimension line");

                XYZ secondPoint = uidoc.Selection.PickPoint(
                    ObjectSnapTypes.Endpoints | ObjectSnapTypes.Midpoints | ObjectSnapTypes.Intersections,
                    "Pick the SECOND point of the dimension line");

                var summary = QuickDimensionFullMixedReferenceProbeService.RunFullMixedReferenceProbe(
                    doc, activeView, firstPoint, secondPoint);

                StringBuilder sb = new StringBuilder();
                sb.AppendLine("=== SESSION 1.5: FULL MIXED REFERENCE SPIKE ===");
                sb.AppendLine();

                sb.AppendLine("--- COLLECTION SUMMARY ---");
                sb.AppendLine($"Grids: {summary.CollectedGridCount} collected → {summary.AcceptedGridCount} accepted");
                if (summary.SkippedArcGridCount > 0)
                    sb.AppendLine($"  - Skipped arc grids: {summary.SkippedArcGridCount}");
                if (summary.SkippedParallelGridCount > 0)
                    sb.AppendLine($"  - Skipped parallel grids: {summary.SkippedParallelGridCount}");

                sb.AppendLine($"Walls: {summary.CollectedWallCount} collected → {summary.AcceptedWallCount} accepted");
                if (summary.SkippedCurtainWallCount > 0)
                    sb.AppendLine($"  - Skipped curtain walls: {summary.SkippedCurtainWallCount}");
                if (summary.SkippedParallelWallCount > 0)
                    sb.AppendLine($"  - Skipped parallel walls: {summary.SkippedParallelWallCount}");
                if (summary.SkippedNoFaceReferenceCount > 0)
                    sb.AppendLine($"  - Skipped no face reference: {summary.SkippedNoFaceReferenceCount}");

                sb.AppendLine($"Doors: {summary.CollectedDoorCount} collected → {summary.AcceptedDoorCount} accepted");
                sb.AppendLine($"Windows: {summary.CollectedWindowCount} collected → {summary.AcceptedWindowCount} accepted");
                if (summary.SkippedNonHostedCount > 0)
                    sb.AppendLine($"  - Skipped non-hosted: {summary.SkippedNonHostedCount}");
                if (summary.SkippedParallelOpeningCount > 0)
                    sb.AppendLine($"  - Skipped parallel openings: {summary.SkippedParallelOpeningCount}");
                if (summary.SkippedOutsideSpanCount > 0)
                    sb.AppendLine($"  - Skipped outside span: {summary.SkippedOutsideSpanCount}");
                if (summary.SkippedNoOpeningReferenceCount > 0)
                    sb.AppendLine($"  - Skipped no opening reference: {summary.SkippedNoOpeningReferenceCount}");

                sb.AppendLine();
                sb.AppendLine($"TOTAL: {summary.TotalCollected} collected → {summary.TotalAccepted} accepted");
                sb.AppendLine();

                sb.AppendLine("--- TEST RESULTS ---");
                sb.AppendLine();

                sb.AppendLine("★ FULL MIXED (Grid + Wall + Door + Window):");
                AppendTestResult(sb, summary.FullMixedResult);
                sb.AppendLine();

                sb.AppendLine("Grid + Wall:");
                AppendTestResult(sb, summary.GridWallResult);
                sb.AppendLine();

                sb.AppendLine("Wall + Opening (Door/Window):");
                AppendTestResult(sb, summary.WallOpeningResult);
                sb.AppendLine();

                sb.AppendLine("Grids Only:");
                AppendTestResult(sb, summary.GridsOnlyResult);
                sb.AppendLine();

                sb.AppendLine("Walls Only:");
                AppendTestResult(sb, summary.WallsOnlyResult);
                sb.AppendLine();

                sb.AppendLine("Openings Only (Door + Window):");
                AppendTestResult(sb, summary.OpeningsOnlyResult);
                sb.AppendLine();

                sb.AppendLine("=== SESSION 1.5 VERDICT ===");
                if (summary.FullMixedReferencesWork)
                {
                    if (summary.AllSourceTypesPresent)
                    {
                        sb.AppendLine("✓ PASS: Full mixed references (Grid + Wall + Door + Window) work!");
                        sb.AppendLine("  All four source types were present and accepted by NewDimension.");
                        sb.AppendLine("  Phase 1 reference feasibility spikes are COMPLETE.");
                    }
                    else
                    {
                        sb.AppendLine("⚠ PARTIAL PASS: Mixed references work, but not all source types present.");
                        sb.AppendLine("  Test with a model containing Grid, Wall, Door, AND Window.");
                    }
                }
                else
                {
                    sb.AppendLine("✗ FAIL: Full mixed references did not work.");
                    sb.AppendLine("  Review the test results above for details.");
                }

                RevitTaskDialog.Show("ArcTool - QD Full Mixed Spike", sb.ToString());
                return Result.Succeeded;
            }
            catch (Autodesk.Revit.Exceptions.OperationCanceledException)
            {
                return Result.Cancelled;
            }
            catch (Exception ex)
            {
                message = ex.Message;
                RevitTaskDialog.Show("ArcTool Error", $"Full Mixed Reference Spike failed:\n\n{ex.Message}");
                return Result.Failed;
            }
        }

        private static void AppendTestResult(StringBuilder sb, QuickDimensionFullMixedTestResult result)
        {
            if (result == null)
            {
                sb.AppendLine("  [No result]");
                return;
            }

            string status = result.Succeeded ? "PASS" : "FAIL";
            sb.AppendLine($"  [{status}] {result.TotalReferences} refs " +
                          $"(G:{result.GridCount} W:{result.WallCount} D:{result.DoorCount} Win:{result.WindowCount})");
            sb.AppendLine($"  {result.Message}");
        }
    }
}
