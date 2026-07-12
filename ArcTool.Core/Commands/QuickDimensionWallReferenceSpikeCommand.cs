using System;
using ArcTool.Core.Models;
using ArcTool.Core.Services;
using Autodesk.Revit.Attributes;
using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using Autodesk.Revit.UI.Selection;
using RevitTaskDialog = Autodesk.Revit.UI.TaskDialog;
using RevitView = Autodesk.Revit.DB.View;

namespace ArcTool.Core.Commands
{
    /// <summary>
    /// Session 1.2 spike command: tests wall face reference strategies for NewDimension.
    /// </summary>
    [Transaction(TransactionMode.Manual)]
    public class QuickDimensionWallReferenceSpikeCommand : IExternalCommand
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
                    "ArcTool — Quick Dimension Wall Spike",
                    "Quick Dimension Phase 1.2 only supports active Plan Views.");
                return Result.Cancelled;
            }

            try
            {
                XYZ firstPoint = uidoc.Selection.PickPoint(
                    ObjectSnapTypes.None,
                    "Quick Dimension wall spike: pick the first point of the dimension line.");

                XYZ secondPoint = uidoc.Selection.PickPoint(
                    ObjectSnapTypes.None,
                    "Quick Dimension wall spike: pick the second point of the dimension line.");

                QuickDimensionWallProbeSummary summary = QuickDimensionWallReferenceProbeService.RunWallReferenceProbe(
                    doc,
                    activeView,
                    firstPoint,
                    secondPoint);

                RevitTaskDialog.Show(
                    "ArcTool — Quick Dimension Wall Spike",
                    BuildSummaryMessage(summary));

                return summary.SideFacesResult.Succeeded || summary.GeometryResult.Succeeded
                    ? Result.Succeeded
                    : Result.Failed;
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
                    $"Quick Dimension wall reference spike failed.\n\n{ex.Message}");
                return Result.Failed;
            }
        }

        private static string BuildSummaryMessage(QuickDimensionWallProbeSummary summary)
        {
            return "Session 1.2 wall face reference probe complete.\n\n" +
                   $"Visible walls collected: {summary.CollectedWallCount}\n" +
                   $"Accepted walls on picked span: {summary.AcceptedWallCount}\n" +
                   $"Skipped curtain walls: {summary.SkippedCurtainWallCount}\n" +
                   $"Skipped walls parallel to dimension line: {summary.SkippedParallelWallCount}\n" +
                   $"Skipped walls with no face reference: {summary.SkippedNoFaceReferenceCount}\n\n" +
                   FormatStrategy(summary.SideFacesResult) + "\n\n" +
                   FormatStrategy(summary.GeometryResult);
        }

        private static string FormatStrategy(QuickDimensionWallStrategyProbeResult result)
        {
            string status = result.Succeeded ? "PASS" : "FAIL";
            return $"{result.Strategy}: {status}\n" +
                   $"References tested: {result.ReferenceCount}\n" +
                   $"Result: {result.Message}";
        }
    }
}
