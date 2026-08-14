using System.Collections.Generic;
using System.Linq;
using Autodesk.Revit.DB;

namespace ArcTool.Core.Archive.QuickDimension.Models
{
    /// <summary>
    /// Read-only relationship observed between a visible candidate wall and the selected wall axis.
    /// This enum is diagnostic only; it does not define a dimension station or aggregation rule.
    /// </summary>
    public enum QuickDimensionWallMidRunRelation
    {
        Ignored = 0,
        EndJoinOnly = 1,
        GeometryJoinOnly = 2,
        MidRunCrossing = 3,
        NonJoinedProximity = 4,
        ParallelNonJoined = 5
    }

    /// <summary>
    /// One distinct vertical Edge.Reference hit observed on the selected side line. It keeps values only;
    /// the transient Revit Reference is discarded after its stable identity key is used for deduplication.
    /// </summary>
    public sealed class QuickDimensionWallMidRunReferenceHit
    {
        public QuickDimensionWallMidRunReferenceHit(
            XYZ midpoint,
            double stationOnSelectedAxis,
            double distanceToSideLine,
            bool candidateReferenceNormalAlongAxis,
            bool selectedWallExposesRefAtStation,
            bool selectedReferenceNormalAlongAxis)
        {
            Midpoint = midpoint;
            StationOnSelectedAxis = stationOnSelectedAxis;
            DistanceToSideLine = distanceToSideLine;
            CandidateReferenceNormalAlongAxis = candidateReferenceNormalAlongAxis;
            SelectedWallExposesRefAtStation = selectedWallExposesRefAtStation;
            SelectedReferenceNormalAlongAxis = selectedReferenceNormalAlongAxis;
        }

        public XYZ Midpoint { get; }
        public double StationOnSelectedAxis { get; }
        public double DistanceToSideLine { get; }
        public bool CandidateReferenceNormalAlongAxis { get; }
        public bool SelectedWallExposesRefAtStation { get; }
        public bool SelectedReferenceNormalAlongAxis { get; }
    }

    /// <summary>
    /// Value-only evidence for one visible candidate wall in the Session 2.7 mid-run probe.
    /// No live Revit Element or Reference is retained. ReferenceHits preserves every distinct side-line
    /// vertical-edge fact; it never reduces a T-joint to one representative station.
    /// </summary>
    public sealed class QuickDimensionWallMidRunCandidate
    {
        public QuickDimensionWallMidRunCandidate(
            long candidateWallId,
            string candidateTypeName,
            QuickDimensionWallMidRunRelation relation,
            bool inElementsAtJoinStart,
            bool inElementsAtJoinEnd,
            bool inGeometryJoin,
            bool isPerpendicular,
            bool isParallel,
            double fallbackStationOnSelectedAxis,
            double fallbackDistanceToSideLine,
            string sourceLabel,
            IReadOnlyList<QuickDimensionWallMidRunReferenceHit> referenceHits,
            int acceptedMidRunStationCount)
        {
            CandidateWallId = candidateWallId;
            CandidateTypeName = candidateTypeName ?? string.Empty;
            Relation = relation;
            InElementsAtJoinStart = inElementsAtJoinStart;
            InElementsAtJoinEnd = inElementsAtJoinEnd;
            InGeometryJoin = inGeometryJoin;
            IsPerpendicular = isPerpendicular;
            IsParallel = isParallel;
            FallbackStationOnSelectedAxis = fallbackStationOnSelectedAxis;
            FallbackDistanceToSideLine = fallbackDistanceToSideLine;
            SourceLabel = sourceLabel ?? string.Empty;
            ReferenceHits = referenceHits ?? new List<QuickDimensionWallMidRunReferenceHit>();
            AcceptedMidRunStationCount = acceptedMidRunStationCount;
        }

        public long CandidateWallId { get; }
        public string CandidateTypeName { get; }
        public QuickDimensionWallMidRunRelation Relation { get; }
        public bool InElementsAtJoinStart { get; }
        public bool InElementsAtJoinEnd { get; }
        public bool InGeometryJoin { get; }
        public bool IsPerpendicular { get; }
        public bool IsParallel { get; }
        public double FallbackStationOnSelectedAxis { get; }
        public double FallbackDistanceToSideLine { get; }
        public string SourceLabel { get; }
        public IReadOnlyList<QuickDimensionWallMidRunReferenceHit> ReferenceHits { get; }
        public int AcceptedMidRunStationCount { get; }

        public bool CandidateWallExposesRefAtStation => ReferenceHits.Count > 0;
        public bool CandidateReferenceNormalAlongAxis => ReferenceHits.Any(hit => hit.CandidateReferenceNormalAlongAxis);
        public bool SelectedWallExposesRefAtStation => ReferenceHits.Any(hit => hit.SelectedWallExposesRefAtStation);
        public bool SelectedReferenceNormalAlongAxis => ReferenceHits.Any(hit => hit.SelectedReferenceNormalAlongAxis);
    }

    /// <summary>
    /// Read-only Session 2.7 log-probe output. It preserves join provenance separately and never contains a chain.
    /// </summary>
    public sealed class QuickDimensionWallMidRunProbeResult
    {
        public QuickDimensionWallMidRunProbeResult(
            bool supported,
            long selectedWallId,
            QuickDimensionWallSpikeSide side,
            ShellLayerType? shell,
            double axisLength,
            IReadOnlyList<long> elementsAtJoinStartIds,
            IReadOnlyList<long> elementsAtJoinEndIds,
            IReadOnlyList<long> geometryJoinIds,
            IReadOnlyList<QuickDimensionWallMidRunCandidate> candidates,
            string message)
        {
            Supported = supported;
            SelectedWallId = selectedWallId;
            Side = side;
            Shell = shell;
            AxisLength = axisLength;
            ElementsAtJoinStartIds = elementsAtJoinStartIds ?? new List<long>();
            ElementsAtJoinEndIds = elementsAtJoinEndIds ?? new List<long>();
            GeometryJoinIds = geometryJoinIds ?? new List<long>();
            Candidates = candidates ?? new List<QuickDimensionWallMidRunCandidate>();
            Message = message ?? string.Empty;
        }

        public bool Supported { get; }
        public long SelectedWallId { get; }
        public QuickDimensionWallSpikeSide Side { get; }
        public ShellLayerType? Shell { get; }
        public double AxisLength { get; }
        public IReadOnlyList<long> ElementsAtJoinStartIds { get; }
        public IReadOnlyList<long> ElementsAtJoinEndIds { get; }
        public IReadOnlyList<long> GeometryJoinIds { get; }
        public IReadOnlyList<QuickDimensionWallMidRunCandidate> Candidates { get; }
        public string Message { get; }
    }
}
