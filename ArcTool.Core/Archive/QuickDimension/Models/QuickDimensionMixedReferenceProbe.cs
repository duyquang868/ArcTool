using Autodesk.Revit.DB;

namespace ArcTool.Core.Archive.QuickDimension.Models
{
    /// <summary>
    /// Source type for mixed dimension reference candidates.
    /// </summary>
    public enum QuickDimensionMixedSourceType
    {
        Grid,
        Wall
    }

    /// <summary>
    /// Unified candidate for mixed Grid + Wall dimension reference testing.
    /// Session 1.3 uses this to merge and sort candidates from different source types.
    /// </summary>
    public sealed class QuickDimensionMixedCandidate
    {
        public QuickDimensionMixedCandidate(
            ElementId elementId,
            QuickDimensionMixedSourceType sourceType,
            string displayName,
            double parameterOnDimensionLine,
            Reference reference)
        {
            ElementId = elementId;
            SourceType = sourceType;
            DisplayName = displayName ?? string.Empty;
            ParameterOnDimensionLine = parameterOnDimensionLine;
            Reference = reference;
        }

        /// <summary>
        /// The element ID (Grid or Wall).
        /// </summary>
        public ElementId ElementId { get; }

        /// <summary>
        /// Whether this candidate is a Grid or Wall.
        /// </summary>
        public QuickDimensionMixedSourceType SourceType { get; }

        /// <summary>
        /// Display name for diagnostics (Grid name or Wall type name).
        /// </summary>
        public string DisplayName { get; }

        /// <summary>
        /// Projection parameter along the dimension line for sorting.
        /// </summary>
        public double ParameterOnDimensionLine { get; }

        /// <summary>
        /// The reference to use for NewDimension.
        /// For Grid: new Reference(grid).
        /// For Wall: HostObjectUtils.GetSideFaces() closest face.
        /// </summary>
        public Reference Reference { get; }
    }

    /// <summary>
    /// Test scenario for mixed reference array probing.
    /// </summary>
    public enum QuickDimensionMixedTestScenario
    {
        /// <summary>
        /// References sorted by position along dimension line (expected to work).
        /// </summary>
        SortedByPosition,

        /// <summary>
        /// References in reverse order (test if Revit rejects or auto-sorts).
        /// </summary>
        ReversedOrder,

        /// <summary>
        /// Grids only (baseline comparison).
        /// </summary>
        GridsOnly,

        /// <summary>
        /// Walls only (baseline comparison).
        /// </summary>
        WallsOnly
    }

    /// <summary>
    /// Result of testing a single mixed reference scenario against NewDimension.
    /// </summary>
    public sealed class QuickDimensionMixedScenarioResult
    {
        public QuickDimensionMixedScenarioResult(
            QuickDimensionMixedTestScenario scenario,
            bool succeeded,
            int gridReferenceCount,
            int wallReferenceCount,
            string message)
        {
            Scenario = scenario;
            Succeeded = succeeded;
            GridReferenceCount = gridReferenceCount;
            WallReferenceCount = wallReferenceCount;
            Message = message ?? string.Empty;
        }

        public QuickDimensionMixedTestScenario Scenario { get; }
        public bool Succeeded { get; }
        public int GridReferenceCount { get; }
        public int WallReferenceCount { get; }
        public int TotalReferenceCount => GridReferenceCount + WallReferenceCount;
        public string Message { get; }
    }

    /// <summary>
    /// Summary of the mixed reference probe session.
    /// </summary>
    public sealed class QuickDimensionMixedProbeSummary
    {
        public QuickDimensionMixedProbeSummary(
            int collectedGridCount,
            int collectedWallCount,
            int acceptedGridCount,
            int acceptedWallCount,
            int skippedArcGridCount,
            int skippedParallelGridCount,
            int skippedCurtainWallCount,
            int skippedParallelWallCount,
            int skippedNoFaceReferenceCount,
            QuickDimensionMixedScenarioResult sortedResult,
            QuickDimensionMixedScenarioResult reversedResult,
            QuickDimensionMixedScenarioResult gridsOnlyResult,
            QuickDimensionMixedScenarioResult wallsOnlyResult)
        {
            CollectedGridCount = collectedGridCount;
            CollectedWallCount = collectedWallCount;
            AcceptedGridCount = acceptedGridCount;
            AcceptedWallCount = acceptedWallCount;
            SkippedArcGridCount = skippedArcGridCount;
            SkippedParallelGridCount = skippedParallelGridCount;
            SkippedCurtainWallCount = skippedCurtainWallCount;
            SkippedParallelWallCount = skippedParallelWallCount;
            SkippedNoFaceReferenceCount = skippedNoFaceReferenceCount;
            SortedResult = sortedResult;
            ReversedResult = reversedResult;
            GridsOnlyResult = gridsOnlyResult;
            WallsOnlyResult = wallsOnlyResult;
        }

        // Grid collection stats
        public int CollectedGridCount { get; }
        public int AcceptedGridCount { get; }
        public int SkippedArcGridCount { get; }
        public int SkippedParallelGridCount { get; }

        // Wall collection stats
        public int CollectedWallCount { get; }
        public int AcceptedWallCount { get; }
        public int SkippedCurtainWallCount { get; }
        public int SkippedParallelWallCount { get; }
        public int SkippedNoFaceReferenceCount { get; }

        // Total accepted
        public int TotalAcceptedCount => AcceptedGridCount + AcceptedWallCount;

        // Scenario results
        public QuickDimensionMixedScenarioResult SortedResult { get; }
        public QuickDimensionMixedScenarioResult ReversedResult { get; }
        public QuickDimensionMixedScenarioResult GridsOnlyResult { get; }
        public QuickDimensionMixedScenarioResult WallsOnlyResult { get; }

        /// <summary>
        /// True if mixed Grid + Wall references work in sorted order.
        /// This is the primary success criterion for Session 1.3.
        /// </summary>
        public bool MixedReferencesWork => SortedResult?.Succeeded == true;
    }
}
