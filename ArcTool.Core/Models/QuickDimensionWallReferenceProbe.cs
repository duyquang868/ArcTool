using Autodesk.Revit.DB;

namespace ArcTool.Core.Models
{
    /// <summary>
    /// Reference strategy options for wall face dimensioning.
    /// Session 1.2 tests which strategy produces valid references for NewDimension.
    /// </summary>
    public enum QuickDimensionWallReferenceStrategy
    {
        /// <summary>
        /// Use HostObjectUtils.GetSideFaces() to get face references directly.
        /// This is the recommended Revit API approach for host objects.
        /// </summary>
        HostObjectUtilsSideFaces,

        /// <summary>
        /// Use Options.ComputeReferences = true with geometry extraction,
        /// then get Face.Reference from planar faces.
        /// </summary>
        GeometryComputeReferences
    }

    /// <summary>
    /// Represents a wall candidate for dimension reference testing.
    /// </summary>
    public sealed class QuickDimensionWallCandidate
    {
        public QuickDimensionWallCandidate(
            ElementId wallId,
            string wallTypeName,
            double parameterOnDimensionLine,
            Reference sideFaceReference,
            Reference geometryFaceReference)
        {
            WallId = wallId;
            WallTypeName = wallTypeName ?? string.Empty;
            ParameterOnDimensionLine = parameterOnDimensionLine;
            SideFaceReference = sideFaceReference;
            GeometryFaceReference = geometryFaceReference;
        }

        /// <summary>
        /// The wall element ID.
        /// </summary>
        public ElementId WallId { get; }

        /// <summary>
        /// The wall type name for diagnostic display.
        /// </summary>
        public string WallTypeName { get; }

        /// <summary>
        /// Projection parameter along the dimension line for sorting.
        /// </summary>
        public double ParameterOnDimensionLine { get; }

        /// <summary>
        /// Reference obtained via HostObjectUtils.GetSideFaces().
        /// May be null if the API call failed.
        /// </summary>
        public Reference SideFaceReference { get; }

        /// <summary>
        /// Reference obtained via Options.ComputeReferences + Face.Reference.
        /// May be null if geometry extraction failed.
        /// </summary>
        public Reference GeometryFaceReference { get; }
    }

    /// <summary>
    /// Result of testing a single reference strategy against NewDimension.
    /// </summary>
    public sealed class QuickDimensionWallStrategyProbeResult
    {
        public QuickDimensionWallStrategyProbeResult(
            QuickDimensionWallReferenceStrategy strategy,
            bool succeeded,
            int referenceCount,
            string message)
        {
            Strategy = strategy;
            Succeeded = succeeded;
            ReferenceCount = referenceCount;
            Message = message ?? string.Empty;
        }

        public QuickDimensionWallReferenceStrategy Strategy { get; }
        public bool Succeeded { get; }
        public int ReferenceCount { get; }
        public string Message { get; }
    }

    /// <summary>
    /// Summary of the wall reference probe session.
    /// </summary>
    public sealed class QuickDimensionWallProbeSummary
    {
        public QuickDimensionWallProbeSummary(
            int collectedWallCount,
            int acceptedWallCount,
            int skippedCurtainWallCount,
            int skippedParallelWallCount,
            int skippedNoFaceReferenceCount,
            QuickDimensionWallStrategyProbeResult sideFacesResult,
            QuickDimensionWallStrategyProbeResult geometryResult)
        {
            CollectedWallCount = collectedWallCount;
            AcceptedWallCount = acceptedWallCount;
            SkippedCurtainWallCount = skippedCurtainWallCount;
            SkippedParallelWallCount = skippedParallelWallCount;
            SkippedNoFaceReferenceCount = skippedNoFaceReferenceCount;
            SideFacesResult = sideFacesResult;
            GeometryResult = geometryResult;
        }

        /// <summary>
        /// Total walls collected from the view.
        /// </summary>
        public int CollectedWallCount { get; }

        /// <summary>
        /// Walls accepted for dimension testing (intersect dimension line, have valid references).
        /// </summary>
        public int AcceptedWallCount { get; }

        /// <summary>
        /// Curtain walls skipped (not supported in V1).
        /// </summary>
        public int SkippedCurtainWallCount { get; }

        /// <summary>
        /// Walls skipped because they are parallel to the dimension line.
        /// </summary>
        public int SkippedParallelWallCount { get; }

        /// <summary>
        /// Walls skipped because no face reference could be obtained.
        /// </summary>
        public int SkippedNoFaceReferenceCount { get; }

        /// <summary>
        /// Result of testing HostObjectUtils.GetSideFaces() strategy.
        /// </summary>
        public QuickDimensionWallStrategyProbeResult SideFacesResult { get; }

        /// <summary>
        /// Result of testing Options.ComputeReferences + Face.Reference strategy.
        /// </summary>
        public QuickDimensionWallStrategyProbeResult GeometryResult { get; }
    }
}
