#nullable enable
using System;
using System.Collections.Generic;
using System.Linq;
using Autodesk.Revit.DB;

namespace ArcTool.Core.Archive.QuickDimension.Models
{
    /// <summary>
    /// Source categories supported by Quick Dimension MVP.
    /// A domain-prefixed enum prevents namespace collisions and keeps collector output source-aware.
    /// </summary>
    public enum QuickDimensionSourceType
    {
        /// <summary>
        /// Straight Revit Grid source.
        /// </summary>
        Grid = 0,

        /// <summary>
        /// Straight non-curtain Wall boundary source.
        /// </summary>
        Wall = 1,

        /// <summary>
        /// Wall-hosted Door opening source.
        /// </summary>
        Door = 2,

        /// <summary>
        /// Wall-hosted Window opening source.
        /// </summary>
        Window = 3
    }

    /// <summary>
    /// Production reference strategies proven by Phase 1 Quick Dimension spikes.
    /// The read-only engine records strategy provenance without creating dimensions.
    /// </summary>
    public enum QuickDimensionReferenceStrategy
    {
        /// <summary>
        /// Grid reference created from the Grid element itself.
        /// </summary>
        GridElementReference = 0,

        /// <summary>
        /// Wall face reference selected from HostObjectUtils.GetSideFaces() in the legacy cross-cutting path.
        /// </summary>
        WallSideFace = 1,

        /// <summary>
        /// Door/Window reference from FamilyInstance.GetReferences(Left/Right).
        /// </summary>
        FamilyInstanceLeftRight = 2,

        /// <summary>
        /// Door/Window fallback reference extracted from host-wall opening geometry.
        /// </summary>
        HostWallOpeningGeometry = 3,

        /// <summary>
        /// Selected host wall end-face reference extracted from wall geometry for the wall-axis projection model.
        /// </summary>
        WallEndFace = 4
    }

    /// <summary>
    /// Diagnostic severity for explaining accepted, skipped, and failed Quick Dimension work.
    /// </summary>
    public enum QuickDimensionDiagnosticSeverity
    {
        /// <summary>
        /// Informational diagnostic, usually for accepted candidates or harmless context.
        /// </summary>
        Info = 0,

        /// <summary>
        /// Expected unsupported or skipped condition that should be visible in summaries.
        /// </summary>
        Warning = 1,

        /// <summary>
        /// Invalid input or isolated API failure that prevents reliable engine output.
        /// </summary>
        Error = 2
    }

    /// <summary>
    /// Source-aware reason codes for Quick Dimension diagnostics.
    /// Explicit reasons prevent silent collector behavior and keep Phase 2 output auditable.
    /// </summary>
    public enum QuickDimensionRejectedReason
    {
        /// <summary>
        /// No rejection; used by informational diagnostics.
        /// </summary>
        None = 0,

        /// <summary>
        /// Active view is not supported by the MVP.
        /// </summary>
        UnsupportedView = 1,

        /// <summary>
        /// Picked points cannot define a valid dimension line.
        /// </summary>
        InvalidDimensionLine = 2,

        /// <summary>
        /// Arc grids are intentionally excluded from MVP.
        /// </summary>
        ArcGridUnsupported = 3,

        /// <summary>
        /// Arc walls are intentionally excluded from MVP.
        /// </summary>
        ArcWallUnsupported = 4,

        /// <summary>
        /// Element direction is incompatible with the picked dimension line.
        /// </summary>
        ParallelToDimensionLine = 5,

        /// <summary>
        /// Element does not intersect the picked span.
        /// </summary>
        OutsidePickedSpan = 6,

        /// <summary>
        /// Curtain walls are intentionally excluded from MVP wall-boundary support.
        /// </summary>
        CurtainWallUnsupported = 7,

        /// <summary>
        /// Door or Window is not hosted by a Wall.
        /// </summary>
        NonWallHostedOpening = 8,

        /// <summary>
        /// No valid Revit Reference could be produced for the candidate.
        /// </summary>
        MissingReference = 9,

        /// <summary>
        /// Candidate was removed by source-aware deduplication.
        /// </summary>
        DuplicateCandidate = 10,

        /// <summary>
        /// Element geometry was missing, invalid, or unusable for the current collector.
        /// </summary>
        InvalidGeometry = 11,

        /// <summary>
        /// Element category is outside the current Quick Dimension MVP scope.
        /// </summary>
        UnsupportedCategory = 12,

        /// <summary>
        /// Collector caught an isolated API exception and skipped the element safely.
        /// </summary>
        CollectorException = 13,

        /// <summary>
        /// Candidate shared a projected station with another candidate and was removed from the final dimension-ready chain.
        /// </summary>
        DuplicateStation = 14
    }

    /// <summary>
    /// Runtime options for the Quick Dimension read-only engine.
    /// Defaults match the MVP scope: Grid, Wall, Door, and Window enabled, with opening fallback enabled.
    /// </summary>
    public sealed class QuickDimensionOptions
    {
        /// <summary>
        /// Default MVP options used when the caller does not expose a settings UI.
        /// </summary>
        public static QuickDimensionOptions Default { get; } = new QuickDimensionOptions();

        public QuickDimensionOptions(
            bool includeGrids = true,
            bool includeWalls = true,
            bool includeDoors = true,
            bool includeWindows = true,
            bool enableHostWallOpeningFallback = true,
            double projectionTolerance = 1e-4,
            double duplicateTolerance = 1e-4,
            double minimumDimensionLineLength = 1e-6,
            double wallEndStationTolerance = 0.0033)
        {
            if (projectionTolerance <= 0.0)
            {
                throw new ArgumentOutOfRangeException(nameof(projectionTolerance), "Projection tolerance must be positive.");
            }

            if (duplicateTolerance <= 0.0)
            {
                throw new ArgumentOutOfRangeException(nameof(duplicateTolerance), "Duplicate tolerance must be positive.");
            }

            if (minimumDimensionLineLength <= 0.0)
            {
                throw new ArgumentOutOfRangeException(nameof(minimumDimensionLineLength), "Minimum dimension line length must be positive.");
            }

            if (wallEndStationTolerance <= 0.0)
            {
                throw new ArgumentOutOfRangeException(nameof(wallEndStationTolerance), "Wall end station tolerance must be positive.");
            }

            IncludeGrids = includeGrids;
            IncludeWalls = includeWalls;
            IncludeDoors = includeDoors;
            IncludeWindows = includeWindows;
            EnableHostWallOpeningFallback = enableHostWallOpeningFallback;
            ProjectionTolerance = projectionTolerance;
            DuplicateTolerance = duplicateTolerance;
            MinimumDimensionLineLength = minimumDimensionLineLength;
            WallEndStationTolerance = wallEndStationTolerance;
        }

        public bool IncludeGrids { get; }
        public bool IncludeWalls { get; }
        public bool IncludeDoors { get; }
        public bool IncludeWindows { get; }
        public bool EnableHostWallOpeningFallback { get; }
        public double ProjectionTolerance { get; }
        public double DuplicateTolerance { get; }
        public double MinimumDimensionLineLength { get; }
        public double WallEndStationTolerance { get; }

        /// <summary>
        /// Returns true when the requested source type is enabled for collection.
        /// </summary>
        public bool IncludesSource(QuickDimensionSourceType sourceType)
        {
            return sourceType switch
            {
                QuickDimensionSourceType.Grid => IncludeGrids,
                QuickDimensionSourceType.Wall => IncludeWalls,
                QuickDimensionSourceType.Door => IncludeDoors,
                QuickDimensionSourceType.Window => IncludeWindows,
                _ => false
            };
        }
    }

    /// <summary>
    /// Normalized context for the dimension axis used by Quick Dimension.
    /// Two construction modes are supported:
    /// (1) the legacy picked-two-point cross-cutting line (kept for the optional/legacy path), and
    /// (2) the wall-axis projection model (ADR-2026-06-11): the selected host wall's straight
    /// LocationCurve is the axis, and a side sign records which side of the wall the dimension goes on.
    /// This keeps later collectors from recalculating direction, length, span, and projection rules independently.
    /// </summary>
    public sealed class QuickDimensionLineContext
    {
        private QuickDimensionLineContext(
            XYZ firstPoint,
            XYZ secondPoint,
            Line line,
            XYZ direction,
            double length,
            bool isWallAxis,
            ElementId? sourceWallId,
            int sideSign)
        {
            FirstPoint = firstPoint;
            SecondPoint = secondPoint;
            Line = line;
            Direction = direction;
            Length = length;
            IsWallAxis = isWallAxis;
            SourceWallId = sourceWallId;
            SideSign = sideSign;
        }

        public XYZ FirstPoint { get; }
        public XYZ SecondPoint { get; }
        public Line Line { get; }
        public XYZ Direction { get; }
        public double Length { get; }

        /// <summary>
        /// True when this context was built from a selected wall's LocationCurve (wall-axis projection model).
        /// </summary>
        public bool IsWallAxis { get; }

        /// <summary>
        /// The selected host wall whose LocationCurve defines the axis, or null for the legacy picked-line model.
        /// </summary>
        public ElementId? SourceWallId { get; }

        /// <summary>
        /// Placement side sign for the wall-axis model: +1 = left of the axis direction, -1 = right, 0 = unspecified.
        /// "Left" is the +90 degree (counter-clockwise) planar normal of <see cref="Direction"/>.
        /// </summary>
        public int SideSign { get; }

        /// <summary>
        /// Planar (XY) unit normal pointing to the placement side, or null when SideSign is unspecified.
        /// </summary>
        public XYZ? SideNormal
        {
            get
            {
                if (SideSign == 0)
                {
                    return null;
                }

                // Left (+1) is the counter-clockwise 90-degree rotation of the planar direction.
                XYZ left = new XYZ(-Direction.Y, Direction.X, 0.0);
                return SideSign > 0 ? left : left.Negate();
            }
        }

        /// <summary>
        /// Creates a validated line context from two picked points (legacy cross-cutting model).
        /// </summary>
        public static QuickDimensionLineContext Create(XYZ firstPoint, XYZ secondPoint, double minimumLength)
        {
            if (firstPoint == null) throw new ArgumentNullException(nameof(firstPoint));
            if (secondPoint == null) throw new ArgumentNullException(nameof(secondPoint));
            if (minimumLength <= 0.0)
            {
                throw new ArgumentOutOfRangeException(nameof(minimumLength), "Minimum dimension line length must be positive.");
            }

            double length = firstPoint.DistanceTo(secondPoint);
            if (length < minimumLength)
            {
                throw new InvalidOperationException("The two picked points are too close to define a Quick Dimension line.");
            }

            Line line = Line.CreateBound(firstPoint, secondPoint);
            XYZ direction = (secondPoint - firstPoint).Normalize();
            return new QuickDimensionLineContext(firstPoint, secondPoint, line, direction, length, false, null, 0);
        }

        /// <summary>
        /// Creates a validated axis context from a selected host wall's straight LocationCurve (wall-axis projection model).
        /// The wall curve endpoints define FirstPoint/SecondPoint and the axis direction, even for skewed walls.
        /// <paramref name="sidePickPoint"/> sets the placement side: the point's signed offset from the axis
        /// determines whether the dimension goes on the left (+1) or right (-1) of the axis direction.
        /// </summary>
        public static QuickDimensionLineContext CreateFromWallAxis(
            ElementId wallId,
            XYZ wallStartPoint,
            XYZ wallEndPoint,
            XYZ? sidePickPoint,
            double minimumLength)
        {
            if (wallId == null) throw new ArgumentNullException(nameof(wallId));
            if (wallStartPoint == null) throw new ArgumentNullException(nameof(wallStartPoint));
            if (wallEndPoint == null) throw new ArgumentNullException(nameof(wallEndPoint));
            if (minimumLength <= 0.0)
            {
                throw new ArgumentOutOfRangeException(nameof(minimumLength), "Minimum dimension line length must be positive.");
            }

            double deltaX = wallEndPoint.X - wallStartPoint.X;
            double deltaY = wallEndPoint.Y - wallStartPoint.Y;
            double planarLength = Math.Sqrt((deltaX * deltaX) + (deltaY * deltaY));
            if (planarLength < minimumLength)
            {
                throw new InvalidOperationException("The selected wall is too short to define a Quick Dimension axis.");
            }

            Line line = Line.CreateBound(wallStartPoint, wallEndPoint);
            XYZ direction = new XYZ(deltaX / planarLength, deltaY / planarLength, 0.0);

            int sideSign = 0;
            if (sidePickPoint != null)
            {
                // Cross product of axis direction with the offset gives the side; positive = left (CCW normal).
                XYZ offset = sidePickPoint - wallStartPoint;
                double cross = (direction.X * offset.Y) - (direction.Y * offset.X);
                if (cross > 0.0)
                {
                    sideSign = 1;
                }
                else if (cross < 0.0)
                {
                    sideSign = -1;
                }
            }

            return new QuickDimensionLineContext(wallStartPoint, wallEndPoint, line, direction, planarLength, true, wallId, sideSign);
        }

        /// <summary>
        /// Projects a point onto the dimension line and returns its signed distance from the first picked point.
        /// </summary>
        public double ProjectParameter(XYZ point)
        {
            if (point == null) throw new ArgumentNullException(nameof(point));
            return (point - FirstPoint).DotProduct(Direction);
        }

        /// <summary>
        /// Evaluates a point on the dimension line at the given parameter from the first picked point.
        /// </summary>
        public XYZ Evaluate(double parameterOnDimensionLine)
        {
            return FirstPoint + (Direction * parameterOnDimensionLine);
        }

        /// <summary>
        /// Returns true when the projected parameter falls inside the picked span within tolerance.
        /// </summary>
        public bool IsInsidePickedSpan(double parameterOnDimensionLine, double tolerance)
        {
            if (tolerance < 0.0)
            {
                throw new ArgumentOutOfRangeException(nameof(tolerance), "Tolerance cannot be negative.");
            }

            return parameterOnDimensionLine >= -tolerance && parameterOnDimensionLine <= Length + tolerance;
        }
    }

    /// <summary>
    /// A valid read-only Quick Dimension candidate with a stable Revit Reference and hit position.
    /// The model intentionally stores ElementId and Reference, not live Element objects.
    /// </summary>
    public sealed class QuickDimensionCandidate
    {
        public QuickDimensionCandidate(
            ElementId elementId,
            QuickDimensionSourceType sourceType,
            string displayName,
            Reference reference,
            QuickDimensionReferenceStrategy referenceStrategy,
            XYZ hitPoint,
            double parameterOnDimensionLine,
            ElementId? hostElementId = null,
            string? familyName = null,
            string? typeName = null)
        {
            if (elementId == null) throw new ArgumentNullException(nameof(elementId));
            if (reference == null) throw new ArgumentNullException(nameof(reference));
            if (hitPoint == null) throw new ArgumentNullException(nameof(hitPoint));
            if (double.IsNaN(parameterOnDimensionLine) || double.IsInfinity(parameterOnDimensionLine))
            {
                throw new ArgumentOutOfRangeException(nameof(parameterOnDimensionLine), "Dimension-line parameter must be finite.");
            }

            ElementId = elementId;
            SourceType = sourceType;
            DisplayName = displayName ?? string.Empty;
            Reference = reference;
            ReferenceStrategy = referenceStrategy;
            HitPoint = hitPoint;
            ParameterOnDimensionLine = parameterOnDimensionLine;
            HostElementId = hostElementId;
            FamilyName = familyName ?? string.Empty;
            TypeName = typeName ?? string.Empty;
        }

        public ElementId ElementId { get; }
        public long ElementValue => ElementId.Value;
        public QuickDimensionSourceType SourceType { get; }
        public string DisplayName { get; }
        public Reference Reference { get; }
        public QuickDimensionReferenceStrategy ReferenceStrategy { get; }
        public XYZ HitPoint { get; }
        public double ParameterOnDimensionLine { get; }
        public ElementId? HostElementId { get; }
        public long? HostElementValue => HostElementId?.Value;
        public string FamilyName { get; }
        public string TypeName { get; }

        public bool IsOpening => SourceType == QuickDimensionSourceType.Door || SourceType == QuickDimensionSourceType.Window;
    }

    /// <summary>
    /// A source-aware diagnostic emitted by the read-only Quick Dimension engine.
    /// Diagnostics are the contract for explaining both accepted and rejected elements.
    /// </summary>
    public sealed class QuickDimensionDiagnostic
    {
        public QuickDimensionDiagnostic(
            QuickDimensionDiagnosticSeverity severity,
            QuickDimensionRejectedReason reason,
            string message,
            ElementId? elementId = null,
            QuickDimensionSourceType? sourceType = null,
            string? displayName = null)
        {
            if (string.IsNullOrWhiteSpace(message))
            {
                throw new ArgumentException("Diagnostic message cannot be empty.", nameof(message));
            }

            Severity = severity;
            Reason = reason;
            Message = message;
            ElementId = elementId;
            SourceType = sourceType;
            DisplayName = displayName ?? string.Empty;
        }

        public QuickDimensionDiagnosticSeverity Severity { get; }
        public QuickDimensionRejectedReason Reason { get; }
        public string Message { get; }
        public ElementId? ElementId { get; }
        public long? ElementValue => ElementId?.Value;
        public QuickDimensionSourceType? SourceType { get; }
        public string DisplayName { get; }

        public bool IsRejected => Reason != QuickDimensionRejectedReason.None;
    }

    /// <summary>
    /// Count summary for one Quick Dimension source type.
    /// Collector services can report how many elements were seen, accepted, and rejected without exposing mutable state.
    /// </summary>
    public sealed class QuickDimensionSourceSummary
    {
        public QuickDimensionSourceSummary(
            QuickDimensionSourceType sourceType,
            int collectedCount,
            int acceptedCount,
            int rejectedCount)
        {
            if (collectedCount < 0) throw new ArgumentOutOfRangeException(nameof(collectedCount));
            if (acceptedCount < 0) throw new ArgumentOutOfRangeException(nameof(acceptedCount));
            if (rejectedCount < 0) throw new ArgumentOutOfRangeException(nameof(rejectedCount));

            SourceType = sourceType;
            CollectedCount = collectedCount;
            AcceptedCount = acceptedCount;
            RejectedCount = rejectedCount;
        }

        public QuickDimensionSourceType SourceType { get; }
        public int CollectedCount { get; }
        public int AcceptedCount { get; }
        public int RejectedCount { get; }
    }

    /// <summary>
    /// Final read-only output contract for Phase 2 before any production NewDimension call exists.
    /// Candidates are expected to be sorted by ParameterOnDimensionLine before construction.
    /// </summary>
    public sealed class QuickDimensionCollectionTimingTrace
    {
        public double TotalWallAxisCollectionMilliseconds { get; set; }
        public double WallEndAnchorCollectionMilliseconds { get; set; }
        public double MidRunAggregationMilliseconds { get; set; }
        public double OpeningCollectionMilliseconds { get; set; }
        public double DuplicateStationReductionMilliseconds { get; set; }
    }

    public sealed class QuickDimensionReadOnlyResult
    {
        public QuickDimensionReadOnlyResult(
            QuickDimensionLineContext lineContext,
            IEnumerable<QuickDimensionCandidate> candidates,
            IEnumerable<QuickDimensionDiagnostic> diagnostics,
            IEnumerable<QuickDimensionSourceSummary> sourceSummaries,
            QuickDimensionOptions options,
            QuickDimensionWallAxisAggregationTrace? wallAxisAggregationTrace = null,
            QuickDimensionCollectionTimingTrace? timingTrace = null)
        {
            LineContext = lineContext ?? throw new ArgumentNullException(nameof(lineContext));
            Options = options ?? throw new ArgumentNullException(nameof(options));
            Candidates = new List<QuickDimensionCandidate>(candidates ?? throw new ArgumentNullException(nameof(candidates))).AsReadOnly();
            Diagnostics = new List<QuickDimensionDiagnostic>(diagnostics ?? throw new ArgumentNullException(nameof(diagnostics))).AsReadOnly();
            SourceSummaries = new List<QuickDimensionSourceSummary>(sourceSummaries ?? throw new ArgumentNullException(nameof(sourceSummaries))).AsReadOnly();
            WallAxisAggregationTrace = wallAxisAggregationTrace;
            TimingTrace = timingTrace;
        }

        public QuickDimensionLineContext LineContext { get; }
        public QuickDimensionOptions Options { get; }
        public IReadOnlyList<QuickDimensionCandidate> Candidates { get; }
        public IReadOnlyList<QuickDimensionDiagnostic> Diagnostics { get; }
        public IReadOnlyList<QuickDimensionSourceSummary> SourceSummaries { get; }
        public QuickDimensionWallAxisAggregationTrace? WallAxisAggregationTrace { get; }
        public QuickDimensionCollectionTimingTrace? TimingTrace { get; }

        public int CandidateCount => Candidates.Count;
        public int DiagnosticCount => Diagnostics.Count;
        public bool CanCreateChainDimension => Candidates.Count >= 2
            && CountDistinctStations(Candidates, Options.DuplicateTolerance) == Candidates.Count;
        public QuickDimensionCandidate? FirstCandidate => Candidates.Count > 0 ? Candidates[0] : null;
        public QuickDimensionCandidate? LastCandidate => Candidates.Count > 0 ? Candidates[Candidates.Count - 1] : null;

        public int GridCount => CountCandidates(QuickDimensionSourceType.Grid);
        public int WallCount => CountCandidates(QuickDimensionSourceType.Wall);
        public int DoorCount => CountCandidates(QuickDimensionSourceType.Door);
        public int WindowCount => CountCandidates(QuickDimensionSourceType.Window);

        private int CountCandidates(QuickDimensionSourceType sourceType)
        {
            int count = 0;
            foreach (QuickDimensionCandidate candidate in Candidates)
            {
                if (candidate.SourceType == sourceType)
                {
                    count++;
                }
            }

            return count;
        }

        private static int CountDistinctStations(IEnumerable<QuickDimensionCandidate> candidates, double tolerance)
        {
            var stations = new List<double>();

            foreach (QuickDimensionCandidate candidate in candidates)
            {
                bool exists = stations.Any(station => Math.Abs(station - candidate.ParameterOnDimensionLine) <= tolerance);
                if (!exists)
                {
                    stations.Add(candidate.ParameterOnDimensionLine);
                }
            }

            return stations.Count;
        }
    }
}
