#nullable enable
using System;
using System.Collections.Generic;
using System.Linq;
using ArcTool.Core.Models;
using Autodesk.Revit.DB;
using RevitView = Autodesk.Revit.DB.View;

namespace ArcTool.Core.Services
{
    /// <summary>
    /// Production read-only Grid candidate collector for Quick Dimension.
    /// This service performs no transactions and never creates dimensions.
    /// </summary>
    public static class QuickDimensionGridCandidateCollector
    {
        /// <summary>
        /// Collects straight Grid candidates visible in the supplied view.
        /// The collector uses true 2D segment/dimension-line intersection instead of midpoint projection.
        /// </summary>
        public static QuickDimensionReadOnlyResult CollectGridCandidates(
            Document doc,
            RevitView view,
            QuickDimensionLineContext lineContext,
            QuickDimensionOptions? options = null)
        {
            if (doc == null) throw new ArgumentNullException(nameof(doc));
            if (view == null) throw new ArgumentNullException(nameof(view));
            if (lineContext == null) throw new ArgumentNullException(nameof(lineContext));

            QuickDimensionOptions effectiveOptions = options ?? QuickDimensionOptions.Default;
            var diagnostics = new List<QuickDimensionDiagnostic>();
            var rawCandidates = new List<QuickDimensionCandidate>();
            int collectedCount = 0;

            if (!effectiveOptions.IncludesSource(QuickDimensionSourceType.Grid))
            {
                diagnostics.Add(new QuickDimensionDiagnostic(
                    QuickDimensionDiagnosticSeverity.Info,
                    QuickDimensionRejectedReason.None,
                    "Grid collection is disabled by Quick Dimension options.",
                    sourceType: QuickDimensionSourceType.Grid));

                return BuildResult(lineContext, rawCandidates, diagnostics, collectedCount, effectiveOptions);
            }

            if (!IsSupportedPlanView(view))
            {
                diagnostics.Add(new QuickDimensionDiagnostic(
                    QuickDimensionDiagnosticSeverity.Error,
                    QuickDimensionRejectedReason.UnsupportedView,
                    $"Quick Dimension Grid collection supports active plan views only. Current view type: {view.ViewType}.",
                    view.Id,
                    QuickDimensionSourceType.Grid,
                    view.Name));

                return BuildResult(lineContext, rawCandidates, diagnostics, collectedCount, effectiveOptions);
            }

            List<Grid> grids = new FilteredElementCollector(doc, view.Id)
                .OfClass(typeof(Grid))
                .Cast<Grid>()
                .ToList();

            collectedCount = grids.Count;

            foreach (Grid grid in grids)
            {
                if (grid?.IsValidObject != true)
                {
                    diagnostics.Add(new QuickDimensionDiagnostic(
                        QuickDimensionDiagnosticSeverity.Warning,
                        QuickDimensionRejectedReason.InvalidGeometry,
                        "Skipped an invalid Grid object returned by the active-view collector.",
                        sourceType: QuickDimensionSourceType.Grid));
                    continue;
                }

                TryCollectGridCandidate(grid, lineContext, effectiveOptions, rawCandidates, diagnostics);
            }

            return BuildResult(lineContext, rawCandidates, diagnostics, collectedCount, effectiveOptions);
        }

        private static void TryCollectGridCandidate(
            Grid grid,
            QuickDimensionLineContext lineContext,
            QuickDimensionOptions options,
            List<QuickDimensionCandidate> candidates,
            List<QuickDimensionDiagnostic> diagnostics)
        {
            string displayName = BuildGridDisplayName(grid);

            try
            {
                Curve gridCurve = grid.Curve;
                if (grid.IsCurved || gridCurve is not Line)
                {
                    diagnostics.Add(new QuickDimensionDiagnostic(
                        QuickDimensionDiagnosticSeverity.Warning,
                        QuickDimensionRejectedReason.ArcGridUnsupported,
                        $"Skipped Grid '{displayName}' because curved grids are outside the Quick Dimension MVP scope.",
                        grid.Id,
                        QuickDimensionSourceType.Grid,
                        displayName));
                    return;
                }

                if (!QuickDimensionGeometryService.TryGetStraightCurveEndpoints(gridCurve, out XYZ startPoint, out XYZ endPoint))
                {
                    diagnostics.Add(new QuickDimensionDiagnostic(
                        QuickDimensionDiagnosticSeverity.Warning,
                        QuickDimensionRejectedReason.InvalidGeometry,
                        $"Skipped Grid '{displayName}' because its curve endpoints are not valid finite points.",
                        grid.Id,
                        QuickDimensionSourceType.Grid,
                        displayName));
                    return;
                }

                if (!QuickDimensionGeometryService.TryGetPlanarDirection(startPoint, endPoint, options.ProjectionTolerance, out XYZ gridDirection))
                {
                    diagnostics.Add(new QuickDimensionDiagnostic(
                        QuickDimensionDiagnosticSeverity.Warning,
                        QuickDimensionRejectedReason.InvalidGeometry,
                        $"Skipped Grid '{displayName}' because its curve does not define a valid plan-view direction.",
                        grid.Id,
                        QuickDimensionSourceType.Grid,
                        displayName));
                    return;
                }

                if (QuickDimensionGeometryService.IsNearlyParallel(lineContext.Direction, gridDirection, options.ProjectionTolerance))
                {
                    diagnostics.Add(new QuickDimensionDiagnostic(
                        QuickDimensionDiagnosticSeverity.Warning,
                        QuickDimensionRejectedReason.ParallelToDimensionLine,
                        $"Skipped Grid '{displayName}' because it is parallel to the picked dimension line.",
                        grid.Id,
                        QuickDimensionSourceType.Grid,
                        displayName));
                    return;
                }

                if (!QuickDimensionGeometryService.TryIntersectSegmentWithDimensionLine2D(
                    lineContext,
                    startPoint,
                    endPoint,
                    options.ProjectionTolerance,
                    out XYZ hitPoint,
                    out double parameterOnDimensionLine))
                {
                    diagnostics.Add(new QuickDimensionDiagnostic(
                        QuickDimensionDiagnosticSeverity.Warning,
                        QuickDimensionRejectedReason.OutsidePickedSpan,
                        $"Skipped Grid '{displayName}' because it does not intersect the picked dimension span.",
                        grid.Id,
                        QuickDimensionSourceType.Grid,
                        displayName));
                    return;
                }

                Reference gridReference = new Reference(grid);
                candidates.Add(new QuickDimensionCandidate(
                    grid.Id,
                    QuickDimensionSourceType.Grid,
                    displayName,
                    gridReference,
                    QuickDimensionReferenceStrategy.GridElementReference,
                    hitPoint,
                    parameterOnDimensionLine));
            }
            catch (Exception ex)
            {
                diagnostics.Add(new QuickDimensionDiagnostic(
                    QuickDimensionDiagnosticSeverity.Warning,
                    QuickDimensionRejectedReason.CollectorException,
                    $"Skipped Grid '{displayName}' because the collector caught an API exception: {ex.Message}",
                    grid.Id,
                    QuickDimensionSourceType.Grid,
                    displayName));
            }
        }

        private static QuickDimensionReadOnlyResult BuildResult(
            QuickDimensionLineContext lineContext,
            IReadOnlyList<QuickDimensionCandidate> rawCandidates,
            List<QuickDimensionDiagnostic> diagnostics,
            int collectedCount,
            QuickDimensionOptions options)
        {
            IReadOnlyList<QuickDimensionCandidate> candidates = QuickDimensionGeometryService
                .DeduplicateCandidates(rawCandidates, options.DuplicateTolerance);

            AddDuplicateDiagnostics(rawCandidates, candidates, diagnostics);
            AddAcceptedDiagnostics(candidates, diagnostics);

            var summaries = new[]
            {
                new QuickDimensionSourceSummary(
                    QuickDimensionSourceType.Grid,
                    collectedCount,
                    candidates.Count,
                    Math.Max(0, collectedCount - candidates.Count))
            };

            return new QuickDimensionReadOnlyResult(
                lineContext,
                candidates,
                diagnostics,
                summaries,
                options);
        }

        private static void AddDuplicateDiagnostics(
            IReadOnlyList<QuickDimensionCandidate> rawCandidates,
            IReadOnlyList<QuickDimensionCandidate> deduplicatedCandidates,
            List<QuickDimensionDiagnostic> diagnostics)
        {
            if (rawCandidates.Count == deduplicatedCandidates.Count)
            {
                return;
            }

            foreach (QuickDimensionCandidate rawCandidate in rawCandidates)
            {
                bool kept = deduplicatedCandidates.Any(candidate => ReferenceEquals(candidate, rawCandidate));
                if (kept)
                {
                    continue;
                }

                diagnostics.Add(new QuickDimensionDiagnostic(
                    QuickDimensionDiagnosticSeverity.Warning,
                    QuickDimensionRejectedReason.DuplicateCandidate,
                    $"Removed duplicate Grid candidate '{rawCandidate.DisplayName}' during source-aware deduplication.",
                    rawCandidate.ElementId,
                    QuickDimensionSourceType.Grid,
                    rawCandidate.DisplayName));
            }
        }

        private static void AddAcceptedDiagnostics(
            IEnumerable<QuickDimensionCandidate> candidates,
            List<QuickDimensionDiagnostic> diagnostics)
        {
            foreach (QuickDimensionCandidate candidate in candidates)
            {
                diagnostics.Add(new QuickDimensionDiagnostic(
                    QuickDimensionDiagnosticSeverity.Info,
                    QuickDimensionRejectedReason.None,
                    $"Accepted Grid candidate '{candidate.DisplayName}' at parameter {candidate.ParameterOnDimensionLine:0.####}.",
                    candidate.ElementId,
                    QuickDimensionSourceType.Grid,
                    candidate.DisplayName));
            }
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

        private static string BuildGridDisplayName(Grid grid)
        {
            string name = grid.Name;
            return string.IsNullOrWhiteSpace(name)
                ? $"Grid {grid.Id.Value}"
                : $"Grid: {name}";
        }
    }
}
