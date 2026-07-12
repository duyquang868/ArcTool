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
    [Transaction(TransactionMode.Manual)]
    public class QuickDimensionGridReferenceSpikeCommand : IExternalCommand
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
                    "ArcTool — Quick Dimension Grid Spike",
                    "Quick Dimension Phase 1.1 only supports active Plan Views.");
                return Result.Cancelled;
            }

            try
            {
                XYZ firstPoint = uidoc.Selection.PickPoint(
                    ObjectSnapTypes.None,
                    "Quick Dimension grid spike: pick the first point of the dimension line.");

                XYZ secondPoint = uidoc.Selection.PickPoint(
                    ObjectSnapTypes.None,
                    "Quick Dimension grid spike: pick the second point of the dimension line.");

                QuickDimensionGridProbeSummary summary = QuickDimensionGridReferenceProbeService.RunGridReferenceProbe(
                    doc,
                    activeView,
                    firstPoint,
                    secondPoint);

                RevitTaskDialog.Show(
                    "ArcTool — Quick Dimension Grid Spike",
                    BuildSummaryMessage(summary));

                return summary.ElementReferenceResult.Succeeded || summary.CurveReferenceResult.Succeeded
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
                    $"Quick Dimension grid reference spike failed.\n\n{ex.Message}");
                return Result.Failed;
            }
        }

        private static string BuildSummaryMessage(QuickDimensionGridProbeSummary summary)
        {
            return "Session 1.1 grid reference probe complete.\n\n" +
                   $"Visible grids collected: {summary.CollectedGridCount}\n" +
                   $"Accepted straight grids on picked span: {summary.AcceptedGridCount}\n" +
                   $"Skipped arc grids: {summary.SkippedArcGridCount}\n" +
                   $"Skipped grids parallel to dimension line: {summary.SkippedParallelGridCount}\n\n" +
                   FormatStrategy(summary.ElementReferenceResult) + "\n\n" +
                   FormatStrategy(summary.CurveReferenceResult);
        }

        private static string FormatStrategy(QuickDimensionGridStrategyProbeResult result)
        {
            string status = result.Succeeded ? "PASS" : "FAIL";
            return $"{result.Strategy}: {status}\n" +
                   $"References tested: {result.ReferenceCount}\n" +
                   $"Result: {result.Message}";
        }
    }
}
