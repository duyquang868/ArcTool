using Autodesk.Revit.DB;

namespace ArcTool.Core.Models
{
    /// <summary>
    /// Source type for full mixed dimension reference candidates.
    /// Session 1.5 extends Session 1.3 to include Door and Window sources.
    /// </summary>
    public enum QuickDimensionFullMixedSourceType
    {
        Grid,
        Wall,
        Door,
        Window
    }

    /// <summary>
    /// Unified candidate for full mixed Grid + Wall + Door + Window dimension reference testing.
    /// Session 1.5 uses this to merge and sort candidates from all four source types.
    /// </summary>
    public sealed class QuickDimensionFullMixedCandidate
    {
        public QuickDimensionFullMixedCandidate(
            ElementId elementId,
            QuickDimensionFullMixedSourceType sourceType,
            string displayName,
            double parameterOnDimensionLine,
            Reference reference,
            ElementId hostWallId = null)
        {
            ElementId = elementId;
            SourceType = sourceType;
            DisplayName = displayName ?? string.Empty;
            ParameterOnDimensionLine = parameterOnDimensionLine;
            Reference = reference;
            HostWallId = hostWallId;
        }

        /// <summary>
        /// The element ID (Grid, Wall, Door, or Window).
        /// </summary>
        public ElementId ElementId { get; }

        /// <summary>
        /// The source type of this candidate.
        /// </summary>
        public QuickDimensionFullMixedSourceType SourceType { get; }

        /// <summary>
        /// Display name for diagnostics (Grid name, Wall type name, or Door/Window family:type).
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
        /// For Door/Window: FamilyInstance.GetReferences(Left/Right) or HostWallOpeningGeometry fallback.
        /// </summary>
        public Reference Reference { get; }

        /// <summary>
        /// Host wall ID for Door/Window candidates. Null for Grid/Wall.
        /// </summary>
        public ElementId HostWallId { get; }

        /// <summary>
        /// True if this is a Door or Window candidate.
        /// </summary>
        public bool IsOpening => SourceType == QuickDimensionFullMixedSourceType.Door ||
                                  SourceType == QuickDimensionFullMixedSourceType.Window;
    }

    /// <summary>
    /// Result of testing the full mixed reference array against NewDimension.
    /// </summary>
    public sealed class QuickDimensionFullMixedTestResult
    {
        public QuickDimensionFullMixedTestResult(
            bool succeeded,
            int gridCount,
            int wallCount,
            int doorCount,
            int windowCount,
            int totalReferences,
            string message)
        {
            Succeeded = succeeded;
            GridCount = gridCount;
            WallCount = wallCount;
            DoorCount = doorCount;
            WindowCount = windowCount;
            TotalReferences = totalReferences;
            Message = message ?? string.Empty;
        }

        public bool Succeeded { get; }
        public int GridCount { get; }
        public int WallCount { get; }
        public int DoorCount { get; }
        public int WindowCount { get; }
        public int TotalReferences { get; }
        public string Message { get; }

        /// <summary>
        /// True if all four source types are represented.
        /// </summary>
        public bool HasAllSourceTypes => GridCount > 0 && WallCount > 0 && DoorCount > 0 && WindowCount > 0;
    }

    /// <summary>
    /// Summary of the full mixed reference probe session.
    /// </summary>
    public sealed class QuickDimensionFullMixedProbeSummary
    {
        public QuickDimensionFullMixedProbeSummary(
            int collectedGridCount,
            int collectedWallCount,
            int collectedDoorCount,
            int collectedWindowCount,
            int acceptedGridCount,
            int acceptedWallCount,
            int acceptedDoorCount,
            int acceptedWindowCount,
            int skippedArcGridCount,
            int skippedParallelGridCount,
            int skippedCurtainWallCount,
            int skippedParallelWallCount,
            int skippedNoFaceReferenceCount,
            int skippedNonHostedCount,
            int skippedParallelOpeningCount,
            int skippedOutsideSpanCount,
            int skippedNoOpeningReferenceCount,
            QuickDimensionFullMixedTestResult fullMixedResult,
            QuickDimensionFullMixedTestResult gridsOnlyResult,
            QuickDimensionFullMixedTestResult wallsOnlyResult,
            QuickDimensionFullMixedTestResult openingsOnlyResult,
            QuickDimensionFullMixedTestResult gridWallResult,
            QuickDimensionFullMixedTestResult wallOpeningResult)
        {
            // Collection counts
            CollectedGridCount = collectedGridCount;
            CollectedWallCount = collectedWallCount;
            CollectedDoorCount = collectedDoorCount;
            CollectedWindowCount = collectedWindowCount;

            // Accepted counts
            AcceptedGridCount = acceptedGridCount;
            AcceptedWallCount = acceptedWallCount;
            AcceptedDoorCount = acceptedDoorCount;
            AcceptedWindowCount = acceptedWindowCount;

            // Skip counts - Grid
            SkippedArcGridCount = skippedArcGridCount;
            SkippedParallelGridCount = skippedParallelGridCount;

            // Skip counts - Wall
            SkippedCurtainWallCount = skippedCurtainWallCount;
            SkippedParallelWallCount = skippedParallelWallCount;
            SkippedNoFaceReferenceCount = skippedNoFaceReferenceCount;

            // Skip counts - Door/Window
            SkippedNonHostedCount = skippedNonHostedCount;
            SkippedParallelOpeningCount = skippedParallelOpeningCount;
            SkippedOutsideSpanCount = skippedOutsideSpanCount;
            SkippedNoOpeningReferenceCount = skippedNoOpeningReferenceCount;

            // Test results
            FullMixedResult = fullMixedResult;
            GridsOnlyResult = gridsOnlyResult;
            WallsOnlyResult = wallsOnlyResult;
            OpeningsOnlyResult = openingsOnlyResult;
            GridWallResult = gridWallResult;
            WallOpeningResult = wallOpeningResult;
        }

        // Collection counts
        public int CollectedGridCount { get; }
        public int CollectedWallCount { get; }
        public int CollectedDoorCount { get; }
        public int CollectedWindowCount { get; }

        // Accepted counts
        public int AcceptedGridCount { get; }
        public int AcceptedWallCount { get; }
        public int AcceptedDoorCount { get; }
        public int AcceptedWindowCount { get; }

        // Skip counts - Grid
        public int SkippedArcGridCount { get; }
        public int SkippedParallelGridCount { get; }

        // Skip counts - Wall
        public int SkippedCurtainWallCount { get; }
        public int SkippedParallelWallCount { get; }
        public int SkippedNoFaceReferenceCount { get; }

        // Skip counts - Door/Window
        public int SkippedNonHostedCount { get; }
        public int SkippedParallelOpeningCount { get; }
        public int SkippedOutsideSpanCount { get; }
        public int SkippedNoOpeningReferenceCount { get; }

        // Totals
        public int TotalCollected => CollectedGridCount + CollectedWallCount + CollectedDoorCount + CollectedWindowCount;
        public int TotalAccepted => AcceptedGridCount + AcceptedWallCount + AcceptedDoorCount + AcceptedWindowCount;
        public int TotalOpeningsCollected => CollectedDoorCount + CollectedWindowCount;
        public int TotalOpeningsAccepted => AcceptedDoorCount + AcceptedWindowCount;

        // Test results
        public QuickDimensionFullMixedTestResult FullMixedResult { get; }
        public QuickDimensionFullMixedTestResult GridsOnlyResult { get; }
        public QuickDimensionFullMixedTestResult WallsOnlyResult { get; }
        public QuickDimensionFullMixedTestResult OpeningsOnlyResult { get; }
        public QuickDimensionFullMixedTestResult GridWallResult { get; }
        public QuickDimensionFullMixedTestResult WallOpeningResult { get; }

        /// <summary>
        /// True if the full mixed test (Grid + Wall + Door + Window) succeeded.
        /// This is the primary success criterion for Session 1.5.
        /// </summary>
        public bool FullMixedReferencesWork => FullMixedResult?.Succeeded == true;

        /// <summary>
        /// True if all four source types were present in the full mixed test.
        /// </summary>
        public bool AllSourceTypesPresent => FullMixedResult?.HasAllSourceTypes == true;
    }
}
