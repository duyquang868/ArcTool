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
    /// Production read-only Quick Dimension engine for Phase 2.6.
    /// This service merges source collectors, performs final ordering/deduplication, and returns diagnostics only.
    /// It performs no transactions and never creates Revit dimensions.
    /// </summary>
    public static class QuickDimensionReadOnlyEngine
    {
        private static readonly QuickDimensionSourceType[] SourceOrder =
        {
            QuickDimensionSourceType.Grid,
            QuickDimensionSourceType.Wall,
            QuickDimensionSourceType.Door,
            QuickDimensionSourceType.Window
        };

        /// <summary>
        /// Collects, merges, sorts, deduplicates, and summarizes all enabled Quick Dimension MVP sources.
        /// The returned result is read-only and is the handoff boundary for the future Phase 3 dimension creation layer.
        /// </summary>
        public static QuickDimensionReadOnlyResult CollectCandidates(
            Document doc,
            RevitView view,
            QuickDimensionLineContext lineContext,
            QuickDimensionOptions? options = null)
        {
            if (doc == null) throw new ArgumentNullException(nameof(doc));
            if (view == null) throw new ArgumentNullException(nameof(view));
            if (lineContext == null) throw new ArgumentNullException(nameof(lineContext));

            QuickDimensionOptions effectiveOptions = options ?? QuickDimensionOptions.Default;
            var rawCandidates = new List<QuickDimensionCandidate>();
            var diagnostics = new List<QuickDimensionDiagnostic>();
            Dictionary<QuickDimensionSourceType, int> collectedCounts = CreateCollectedCountMap();

            if (!IsSupportedPlanView(view))
            {
                diagnostics.Add(new QuickDimensionDiagnostic(
                    QuickDimensionDiagnosticSeverity.Error,
                    QuickDimensionRejectedReason.UnsupportedView,
                    $"Quick Dimension read-only engine supports active plan views only. Current view type: {view.ViewType}.",
                    view.Id,
                    displayName: view.Name));

                return BuildFinalResult(lineContext, rawCandidates, diagnostics, collectedCounts, effectiveOptions);
            }

            // Wall-axis projection model (ADR-2026-06-11): the selected wall IS the axis. Gather references only
            // from that wall (its end faces + hosted Door/Window jambs), projected onto the axis. Grid and other
            // walls are intentionally excluded from this main flow.
            if (lineContext.IsWallAxis)
            {
                CollectWallAxisCandidates(doc, view, lineContext, effectiveOptions, rawCandidates, diagnostics, collectedCounts);
                return BuildFinalResult(lineContext, rawCandidates, diagnostics, collectedCounts, effectiveOptions);
            }

            if (effectiveOptions.IncludeGrids)
            {
                AddCollectorResult(
                    QuickDimensionGridCandidateCollector.CollectGridCandidates(doc, view, lineContext, effectiveOptions),
                    rawCandidates,
                    diagnostics,
                    collectedCounts);
            }
            else
            {
                AddDisabledDiagnostic(diagnostics, QuickDimensionSourceType.Grid);
            }

            if (effectiveOptions.IncludeWalls)
            {
                AddCollectorResult(
                    QuickDimensionWallCandidateCollector.CollectWallCandidates(doc, view, lineContext, effectiveOptions),
                    rawCandidates,
                    diagnostics,
                    collectedCounts);
            }
            else
            {
                AddDisabledDiagnostic(diagnostics, QuickDimensionSourceType.Wall);
            }

            if (effectiveOptions.IncludeDoors || effectiveOptions.IncludeWindows)
            {
                AddCollectorResult(
                    QuickDimensionDoorWindowCandidateCollector.CollectDoorWindowCandidates(doc, view, lineContext, effectiveOptions),
                    rawCandidates,
                    diagnostics,
                    collectedCounts);
            }
            else
            {
                AddDisabledDiagnostic(diagnostics, QuickDimensionSourceType.Door);
                AddDisabledDiagnostic(diagnostics, QuickDimensionSourceType.Window);
            }

            return BuildFinalResult(lineContext, rawCandidates, diagnostics, collectedCounts, effectiveOptions);
        }

        /// <summary>
        /// Wall-axis projection model orchestration: resolves the selected wall from the context, then merges
        /// its end-face anchors and its hosted Door/Window jambs, all projected onto the wall axis.
        /// </summary>
        private static void CollectWallAxisCandidates(
            Document doc,
            RevitView view,
            QuickDimensionLineContext lineContext,
            QuickDimensionOptions effectiveOptions,
            List<QuickDimensionCandidate> rawCandidates,
            List<QuickDimensionDiagnostic> diagnostics,
            Dictionary<QuickDimensionSourceType, int> collectedCounts)
        {
            // Grid is excluded from the wall-axis main flow by design.
            AddDisabledDiagnostic(diagnostics, QuickDimensionSourceType.Grid);

            ElementId? wallId = lineContext.SourceWallId;
            if (wallId == null || doc.GetElement(wallId) is not Wall selectedWall || !selectedWall.IsValidObject)
            {
                diagnostics.Add(new QuickDimensionDiagnostic(
                    QuickDimensionDiagnosticSeverity.Error,
                    QuickDimensionRejectedReason.InvalidGeometry,
                    "Quick Dimension wall-axis model could not resolve the selected wall from the axis context.",
                    wallId));
                return;
            }

            if (effectiveOptions.IncludeWalls)
            {
                AddCollectorResult(
                    QuickDimensionWallCandidateCollector.CollectSelectedWallEndAnchors(doc, lineContext, selectedWall, effectiveOptions),
                    rawCandidates,
                    diagnostics,
                    collectedCounts);
            }
            else
            {
                AddDisabledDiagnostic(diagnostics, QuickDimensionSourceType.Wall);
            }

            if (effectiveOptions.IncludeDoors || effectiveOptions.IncludeWindows)
            {
                AddCollectorResult(
                    QuickDimensionDoorWindowCandidateCollector.CollectOpeningsAlongWallAxis(doc, view, lineContext, selectedWall.Id, effectiveOptions),
                    rawCandidates,
                    diagnostics,
                    collectedCounts);
            }
            else
            {
                AddDisabledDiagnostic(diagnostics, QuickDimensionSourceType.Door);
                AddDisabledDiagnostic(diagnostics, QuickDimensionSourceType.Window);
            }
        }

        private static QuickDimensionReadOnlyResult BuildFinalResult(
            QuickDimensionLineContext lineContext,
            IReadOnlyList<QuickDimensionCandidate> rawCandidates,
            List<QuickDimensionDiagnostic> diagnostics,
            IReadOnlyDictionary<QuickDimensionSourceType, int> collectedCounts,
            QuickDimensionOptions options)
        {
            IReadOnlyList<QuickDimensionCandidate> sourceAwareCandidates = QuickDimensionGeometryService
                .DeduplicateCandidates(rawCandidates, options.DuplicateTolerance);
            IReadOnlyList<QuickDimensionCandidate> finalCandidates = RemoveDuplicateStations(
                sourceAwareCandidates,
                options.DuplicateTolerance,
                diagnostics);

            AddEngineSummaryDiagnostic(rawCandidates.Count, finalCandidates.Count, diagnostics);
            AddDuplicateDiagnostics(rawCandidates, sourceAwareCandidates, diagnostics);
            AddAcceptedDiagnostics(finalCandidates, diagnostics);
            AddMinimumReferenceDiagnostic(finalCandidates, diagnostics);

            IReadOnlyList<QuickDimensionSourceSummary> summaries = BuildSourceSummaries(finalCandidates, collectedCounts);

            return new QuickDimensionReadOnlyResult(
                lineContext,
                finalCandidates,
                diagnostics,
                summaries,
                options);
        }

        private static void AddCollectorResult(
            QuickDimensionReadOnlyResult collectorResult,
            List<QuickDimensionCandidate> rawCandidates,
            List<QuickDimensionDiagnostic> diagnostics,
            Dictionary<QuickDimensionSourceType, int> collectedCounts)
        {
            if (collectorResult == null) throw new ArgumentNullException(nameof(collectorResult));

            rawCandidates.AddRange(collectorResult.Candidates);

            foreach (QuickDimensionDiagnostic diagnostic in collectorResult.Diagnostics)
            {
                if (!IsCollectorAcceptedDiagnostic(diagnostic))
                {
                    diagnostics.Add(diagnostic);
                }
            }

            foreach (QuickDimensionSourceSummary summary in collectorResult.SourceSummaries)
            {
                if (collectedCounts.ContainsKey(summary.SourceType))
                {
                    collectedCounts[summary.SourceType] += summary.CollectedCount;
                }
            }
        }

        private static bool IsCollectorAcceptedDiagnostic(QuickDimensionDiagnostic diagnostic)
        {
            if (diagnostic == null)
            {
                return false;
            }

            return diagnostic.Reason == QuickDimensionRejectedReason.None
                && diagnostic.Message.StartsWith("Accepted ", StringComparison.Ordinal);
        }

        private static void AddEngineSummaryDiagnostic(
            int rawCandidateCount,
            int finalCandidateCount,
            List<QuickDimensionDiagnostic> diagnostics)
        {
            diagnostics.Add(new QuickDimensionDiagnostic(
                QuickDimensionDiagnosticSeverity.Info,
                QuickDimensionRejectedReason.None,
                $"Quick Dimension read-only engine merged {rawCandidateCount} collector candidates into {finalCandidateCount} final ordered candidates."));
        }

        private static void AddDuplicateDiagnostics(
            IReadOnlyList<QuickDimensionCandidate> rawCandidates,
            IReadOnlyList<QuickDimensionCandidate> finalCandidates,
            List<QuickDimensionDiagnostic> diagnostics)
        {
            if (rawCandidates.Count == finalCandidates.Count)
            {
                return;
            }

            foreach (QuickDimensionCandidate rawCandidate in rawCandidates)
            {
                bool kept = finalCandidates.Any(candidate => ReferenceEquals(candidate, rawCandidate));
                if (kept)
                {
                    continue;
                }

                diagnostics.Add(new QuickDimensionDiagnostic(
                    QuickDimensionDiagnosticSeverity.Warning,
                    QuickDimensionRejectedReason.DuplicateCandidate,
                    $"Removed duplicate {rawCandidate.SourceType} candidate '{rawCandidate.DisplayName}' during final read-only engine deduplication.",
                    rawCandidate.ElementId,
                    rawCandidate.SourceType,
                    rawCandidate.DisplayName));
            }
        }

        private static void AddAcceptedDiagnostics(
            IEnumerable<QuickDimensionCandidate> finalCandidates,
            List<QuickDimensionDiagnostic> diagnostics)
        {
            foreach (QuickDimensionCandidate candidate in finalCandidates)
            {
                diagnostics.Add(new QuickDimensionDiagnostic(
                    QuickDimensionDiagnosticSeverity.Info,
                    QuickDimensionRejectedReason.None,
                    $"Accepted final {candidate.SourceType} candidate '{candidate.DisplayName}' at parameter {candidate.ParameterOnDimensionLine:0.####}. Strategy: {candidate.ReferenceStrategy}.",
                    candidate.ElementId,
                    candidate.SourceType,
                    candidate.DisplayName));
            }
        }

        private static void AddMinimumReferenceDiagnostic(
            IReadOnlyList<QuickDimensionCandidate> finalCandidates,
            List<QuickDimensionDiagnostic> diagnostics)
        {
            if (finalCandidates.Count >= 2)
            {
                return;
            }

            diagnostics.Add(new QuickDimensionDiagnostic(
                QuickDimensionDiagnosticSeverity.Warning,
                QuickDimensionRejectedReason.MissingReference,
                $"Quick Dimension needs at least 2 final candidates to create a chain dimension later; current final candidate count is {finalCandidates.Count}."));
        }

        /// <summary>
        /// Removes candidates that share the same projected station on the dimension line after source-aware
        /// dedupe. Zero-length dimension segments are unsafe for chain dimensioning, so only the first sorted
        /// candidate at each station is kept and later collisions are recorded as DuplicateStation diagnostics.
        /// </summary>
        private static IReadOnlyList<QuickDimensionCandidate> RemoveDuplicateStations(
            IReadOnlyList<QuickDimensionCandidate> candidates,
            double duplicateTolerance,
            List<QuickDimensionDiagnostic> diagnostics)
        {
            IReadOnlyList<QuickDimensionCandidate> sortedCandidates = QuickDimensionGeometryService.SortByDimensionParameter(candidates);
            var result = new List<QuickDimensionCandidate>();

            foreach (QuickDimensionCandidate candidate in sortedCandidates)
            {
                QuickDimensionCandidate? keptAtSameStation = result
                    .FirstOrDefault(kept => Math.Abs(kept.ParameterOnDimensionLine - candidate.ParameterOnDimensionLine) <= duplicateTolerance);

                if (keptAtSameStation != null)
                {
                    double stationMm = UnitUtils.ConvertFromInternalUnits(candidate.ParameterOnDimensionLine, UnitTypeId.Millimeters);
                    diagnostics.Add(new QuickDimensionDiagnostic(
                        QuickDimensionDiagnosticSeverity.Warning,
                        QuickDimensionRejectedReason.DuplicateStation,
                        $"Removed {candidate.SourceType} candidate '{candidate.DisplayName}' at station {stationMm:0.##} mm because it shares that station with '{keptAtSameStation.DisplayName}' ({keptAtSameStation.SourceType}). Chain dimensions cannot use zero-length segments.",
                        candidate.ElementId,
                        candidate.SourceType,
                        candidate.DisplayName));
                    continue;
                }

                result.Add(candidate);
            }

            return result.AsReadOnly();
        }

        private static IReadOnlyList<QuickDimensionSourceSummary> BuildSourceSummaries(
            IReadOnlyList<QuickDimensionCandidate> finalCandidates,
            IReadOnlyDictionary<QuickDimensionSourceType, int> collectedCounts)
        {
            var summaries = new List<QuickDimensionSourceSummary>();

            foreach (QuickDimensionSourceType sourceType in SourceOrder)
            {
                int collectedCount = collectedCounts.TryGetValue(sourceType, out int value) ? value : 0;
                int acceptedElementCount = CountAcceptedElements(finalCandidates, sourceType);

                summaries.Add(new QuickDimensionSourceSummary(
                    sourceType,
                    collectedCount,
                    acceptedElementCount,
                    Math.Max(0, collectedCount - acceptedElementCount)));
            }

            return summaries.AsReadOnly();
        }

        private static int CountAcceptedElements(
            IEnumerable<QuickDimensionCandidate> candidates,
            QuickDimensionSourceType sourceType)
        {
            return candidates
                .Where(candidate => candidate.SourceType == sourceType)
                .Select(candidate => candidate.ElementValue)
                .Distinct()
                .Count();
        }

        private static Dictionary<QuickDimensionSourceType, int> CreateCollectedCountMap()
        {
            var result = new Dictionary<QuickDimensionSourceType, int>();
            foreach (QuickDimensionSourceType sourceType in SourceOrder)
            {
                result[sourceType] = 0;
            }

            return result;
        }

        private static void AddDisabledDiagnostic(
            List<QuickDimensionDiagnostic> diagnostics,
            QuickDimensionSourceType sourceType)
        {
            diagnostics.Add(new QuickDimensionDiagnostic(
                QuickDimensionDiagnosticSeverity.Info,
                QuickDimensionRejectedReason.None,
                $"{sourceType} collection is disabled by Quick Dimension options.",
                sourceType: sourceType));
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
