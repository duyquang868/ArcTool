#nullable enable
using System;
using System.Collections.Generic;
using System.Linq;
using ArcTool.Core.Models;
using ArcTool.Core.Services;
using Autodesk.Revit.Attributes;
using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using RevitDocument = Autodesk.Revit.DB.Document;
using RevitTaskDialog = Autodesk.Revit.UI.TaskDialog;

namespace ArcTool.Core.Commands
{
    /// <summary>
    /// Revit external command that runs the coordinate batch write workflow for supported coordinate elements.
    /// The command owns user-facing guards, transaction scope, axis-mapping selection, and the post-run summary dialog.
    /// </summary>
    [Transaction(TransactionMode.Manual)]
    public class RunCoordBatchCommand : IExternalCommand
    {
        /// <summary>
        /// Executes the batch coordinate write command in the active Revit document.
        /// </summary>
        /// <param name="commandData">Revit command context.</param>
        /// <param name="message">Failure message returned to Revit when the command fails.</param>
        /// <param name="elements">Element set returned to Revit when the command fails.</param>
        /// <returns>Revit command result.</returns>
        public Result Execute(ExternalCommandData commandData, ref string message, ElementSet elements)
        {
            RevitDocument doc;

            try
            {
                RevitDocument? activeDoc = commandData.Application.ActiveUIDocument?.Document;
                if (activeDoc == null)
                {
                    RevitTaskDialog.Show("ArcTool Error", "No document is open.");
                    return Result.Failed;
                }

                doc = activeDoc;
            }
            catch (Exception ex)
            {
                message = ex.Message;
                RevitTaskDialog.Show("ArcTool Error", $"Could not access the active document.\n\n{ex.Message}");
                return Result.Failed;
            }

            CoordinateProjectSettings settings = CoordinateProjectSettingsService.LoadOrDefault(doc);
            IReadOnlyList<CoordTriggerFilter> registeredFilters = CoordinateExtractionService.GetRegisteredTriggerFilters(doc);
            if (registeredFilters.Count == 0)
            {
                RevitTaskDialog.Show(
                    "ArcTool — Coordinate Batch",
                    "Shared parameters are not registered for any coordinate scope.\n\n" +
                    "Run 'Register Element Type' for 3D elements or 'Register Detail Type' for Detail Items first, then run this command again.");
                return Result.Failed;
            }

            CoordBatchSummary summary;

            using var tx = new Transaction(doc, "ArcTool: Write Coordinates");
            tx.Start();

            try
            {
                summary = CoordinateBatchService.RunBatch(doc, settings.AxisMapping, settings.ParameterUnit);
                tx.Commit();
            }
            catch (Exception ex)
            {
                tx.RollBack();
                message = ex.Message;
                RevitTaskDialog.Show("ArcTool Error",
                    $"Batch run failed and was rolled back.\n\n{ex.Message}");
                return Result.Failed;
            }

            CoordinateLogService.LogBatch(doc, summary);

            string body = $"Coordinate write complete.\n\n" +
                $"  Total elements found: {summary.TotalCollected}\n" +
                $"  Written (new/changed): {summary.WrittenCount}\n" +
                $"  Skipped (no change):   {summary.SkippedCount}\n" +
                $"  Unsupported geometry:  {summary.UnsupportedCount}\n" +
                $"  Failed (param error):  {summary.FailedCount}";

            if (summary.UnsupportedCount > 0 || summary.FailedCount > 0)
            {
                List<string> problemLines = summary.Details
                    .Where(r => r.Outcome == CoordWriteOutcome.Unsupported
                             || r.Outcome == CoordWriteOutcome.Failed)
                    .Select(r => $"  [{r.Outcome}] Id={r.ElementId}: {r.DiagnosticMessage}")
                    .ToList();

                string detail = problemLines.Count <= 20
                    ? string.Join("\n", problemLines)
                    : string.Join("\n", problemLines.Take(20))
                      + $"\n  ... and {problemLines.Count - 20} more.";

                body += $"\n\nProblematic elements:\n{detail}";
            }

            RevitTaskDialog.Show("ArcTool — Coordinate Batch", body);
            return Result.Succeeded;
        }

        private static FamilyInstance? FindFirstSupportedFamilyInstance(RevitDocument doc, CoordTriggerFilter triggerFilter)
        {
            BuiltInCategory category = CoordV1Scope.GetCategory(triggerFilter);
            return new FilteredElementCollector(doc)
                .OfClass(typeof(FamilyInstance))
                .OfCategory(category)
                .Cast<FamilyInstance>()
                .FirstOrDefault(instance => triggerFilter != CoordTriggerFilter.DetailItems
                    || CoordinateDetailItemRegistryService.IsRegisteredType(doc, instance));
        }
    }
}
