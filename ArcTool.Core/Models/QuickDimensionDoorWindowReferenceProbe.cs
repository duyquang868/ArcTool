using Autodesk.Revit.DB;

namespace ArcTool.Core.Models
{
    /// <summary>
    /// Reference strategy for Door/Window opening dimensions.
    /// Session 1.4 spike tests all three strategies to determine which works with NewDimension.
    /// </summary>
    public enum QuickDimensionDoorWindowReferenceStrategy
    {
        /// <summary>
        /// Uses FamilyInstance.GetReferences(FamilyInstanceReferenceType.Left/Right).
        /// Requires family to have properly named reference planes.
        /// </summary>
        FamilyInstanceReferences,

        /// <summary>
        /// Uses Options.ComputeReferences = true and extracts Face.Reference from geometry.
        /// General approach that doesn't depend on family definition.
        /// </summary>
        GeometryComputeReferences,

        /// <summary>
        /// Uses Wall.FindInserts() to get hosted openings, then extracts opening cut geometry.
        /// Doesn't depend on family definition but requires host wall context.
        /// </summary>
        HostWallOpeningGeometry
    }

    /// <summary>
    /// Source type for Door/Window candidates.
    /// </summary>
    public enum QuickDimensionDoorWindowSourceType
    {
        Door,
        Window
    }

    /// <summary>
    /// Represents a single Door or Window opening candidate for dimensioning.
    /// Each candidate may have references from multiple strategies.
    /// </summary>
    public sealed class QuickDimensionDoorWindowCandidate
    {
        public QuickDimensionDoorWindowCandidate(
            ElementId elementId,
            QuickDimensionDoorWindowSourceType sourceType,
            string familyName,
            string typeName,
            double parameterOnDimensionLine,
            ElementId hostWallId)
        {
            ElementId = elementId;
            SourceType = sourceType;
            FamilyName = familyName ?? string.Empty;
            TypeName = typeName ?? string.Empty;
            ParameterOnDimensionLine = parameterOnDimensionLine;
            HostWallId = hostWallId;
        }

        public ElementId ElementId { get; }
        public QuickDimensionDoorWindowSourceType SourceType { get; }
        public string FamilyName { get; }
        public string TypeName { get; }
        public double ParameterOnDimensionLine { get; }
        public ElementId HostWallId { get; }

        // References from different strategies - populated by service
        public Reference LeftFamilyInstanceReference { get; set; }
        public Reference RightFamilyInstanceReference { get; set; }
        public Reference LeftGeometryReference { get; set; }
        public Reference RightGeometryReference { get; set; }
        public Reference LeftOpeningReference { get; set; }
        public Reference RightOpeningReference { get; set; }

        /// <summary>
        /// Gets the left reference for the specified strategy, or null if not available.
        /// </summary>
        public Reference GetLeftReference(QuickDimensionDoorWindowReferenceStrategy strategy)
        {
            return strategy switch
            {
                QuickDimensionDoorWindowReferenceStrategy.FamilyInstanceReferences => LeftFamilyInstanceReference,
                QuickDimensionDoorWindowReferenceStrategy.GeometryComputeReferences => LeftGeometryReference,
                QuickDimensionDoorWindowReferenceStrategy.HostWallOpeningGeometry => LeftOpeningReference,
                _ => null
            };
        }

        /// <summary>
        /// Gets the right reference for the specified strategy, or null if not available.
        /// </summary>
        public Reference GetRightReference(QuickDimensionDoorWindowReferenceStrategy strategy)
        {
            return strategy switch
            {
                QuickDimensionDoorWindowReferenceStrategy.FamilyInstanceReferences => RightFamilyInstanceReference,
                QuickDimensionDoorWindowReferenceStrategy.GeometryComputeReferences => RightGeometryReference,
                QuickDimensionDoorWindowReferenceStrategy.HostWallOpeningGeometry => RightOpeningReference,
                _ => null
            };
        }

        /// <summary>
        /// Returns true if this candidate has at least one valid reference pair for the strategy.
        /// </summary>
        public bool HasReferences(QuickDimensionDoorWindowReferenceStrategy strategy)
        {
            return GetLeftReference(strategy) != null || GetRightReference(strategy) != null;
        }
    }

    /// <summary>
    /// Result of testing a single reference strategy against NewDimension.
    /// </summary>
    public sealed class QuickDimensionDoorWindowStrategyResult
    {
        public QuickDimensionDoorWindowStrategyResult(
            QuickDimensionDoorWindowReferenceStrategy strategy,
            bool succeeded,
            int totalCandidates,
            int candidatesWithReferences,
            int referencesUsed,
            string message)
        {
            Strategy = strategy;
            Succeeded = succeeded;
            TotalCandidates = totalCandidates;
            CandidatesWithReferences = candidatesWithReferences;
            ReferencesUsed = referencesUsed;
            Message = message ?? string.Empty;
        }

        public QuickDimensionDoorWindowReferenceStrategy Strategy { get; }
        public bool Succeeded { get; }
        public int TotalCandidates { get; }
        public int CandidatesWithReferences { get; }
        public int ReferencesUsed { get; }
        public string Message { get; }
    }

    /// <summary>
    /// Summary of the Door/Window reference spike probe.
    /// </summary>
    public sealed class QuickDimensionDoorWindowProbeSummary
    {
        public QuickDimensionDoorWindowProbeSummary(
            int collectedDoorCount,
            int collectedWindowCount,
            int acceptedDoorCount,
            int acceptedWindowCount,
            int skippedNonHostedCount,
            int skippedParallelCount,
            int skippedOutsideSpanCount,
            QuickDimensionDoorWindowStrategyResult familyInstanceResult,
            QuickDimensionDoorWindowStrategyResult geometryResult,
            QuickDimensionDoorWindowStrategyResult openingGeometryResult)
        {
            CollectedDoorCount = collectedDoorCount;
            CollectedWindowCount = collectedWindowCount;
            AcceptedDoorCount = acceptedDoorCount;
            AcceptedWindowCount = acceptedWindowCount;
            SkippedNonHostedCount = skippedNonHostedCount;
            SkippedParallelCount = skippedParallelCount;
            SkippedOutsideSpanCount = skippedOutsideSpanCount;
            FamilyInstanceResult = familyInstanceResult;
            GeometryResult = geometryResult;
            OpeningGeometryResult = openingGeometryResult;
        }

        public int CollectedDoorCount { get; }
        public int CollectedWindowCount { get; }
        public int AcceptedDoorCount { get; }
        public int AcceptedWindowCount { get; }
        public int SkippedNonHostedCount { get; }
        public int SkippedParallelCount { get; }
        public int SkippedOutsideSpanCount { get; }

        public int TotalCollected => CollectedDoorCount + CollectedWindowCount;
        public int TotalAccepted => AcceptedDoorCount + AcceptedWindowCount;

        public QuickDimensionDoorWindowStrategyResult FamilyInstanceResult { get; }
        public QuickDimensionDoorWindowStrategyResult GeometryResult { get; }
        public QuickDimensionDoorWindowStrategyResult OpeningGeometryResult { get; }

        /// <summary>
        /// Returns true if at least one strategy succeeded.
        /// </summary>
        public bool AnyStrategySucceeded =>
            (FamilyInstanceResult?.Succeeded ?? false) ||
            (GeometryResult?.Succeeded ?? false) ||
            (OpeningGeometryResult?.Succeeded ?? false);

        /// <summary>
        /// Returns the best strategy (first one that succeeded), or null if none succeeded.
        /// </summary>
        public QuickDimensionDoorWindowReferenceStrategy? BestStrategy
        {
            get
            {
                // Priority order: FamilyInstance > Geometry > OpeningGeometry
                if (FamilyInstanceResult?.Succeeded == true)
                    return QuickDimensionDoorWindowReferenceStrategy.FamilyInstanceReferences;
                if (GeometryResult?.Succeeded == true)
                    return QuickDimensionDoorWindowReferenceStrategy.GeometryComputeReferences;
                if (OpeningGeometryResult?.Succeeded == true)
                    return QuickDimensionDoorWindowReferenceStrategy.HostWallOpeningGeometry;
                return null;
            }
        }
    }
}
