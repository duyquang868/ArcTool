#nullable enable
using System.Collections.Generic;
using Autodesk.Revit.DB;

namespace ArcTool.Core.Archive.QuickDimension.Models
{
    /// <summary>
    /// Lightweight, mutable trace for the production wall-axis mid-run aggregator.
    /// It exists ONLY so the read-only XML log can explain why each visible candidate wall was
    /// classified <c>MidRunCrossing</c>, <c>EndJoinOnly</c>, <c>Ignored</c>, etc. The aggregator owns
    /// classification; this trace only records the decision. No transaction and no dimension are created.
    ///
    /// Reference discipline: within the runtime command scope the trace may hold live <see cref="Reference"/>
    /// and <see cref="XYZ"/> values so the XML service can convert them to stable strings and survey points.
    /// It is never stored in the long-lived read-only result contract beyond this diagnostic channel.
    /// </summary>
    public sealed class QuickDimensionWallAxisAggregationTrace
    {
        public bool Supported { get; set; }
        public string Message { get; set; } = string.Empty;
        public long SelectedWallId { get; set; }
        public string SideLabel { get; set; } = string.Empty;
        public int SideSign { get; set; }
        public ShellLayerType? ShellLayer { get; set; }
        public double AxisLength { get; set; }
        public XYZ? SideNormal { get; set; }
        public XYZ? SideLinePoint { get; set; }
        public double ResolvedAnchorMinStation { get; set; }
        public double ResolvedAnchorMaxStation { get; set; }
        public QuickDimensionWallAxisAnchorTrace? StartAnchor { get; set; }
        public QuickDimensionWallAxisAnchorTrace? FinishAnchor { get; set; }
        public List<long> ElementsAtJoinStartIds { get; } = new List<long>();
        public List<long> ElementsAtJoinEndIds { get; } = new List<long>();
        public List<long> GeometryJoinIds { get; } = new List<long>();
        public List<QuickDimensionWallAxisCandidateTrace> Candidates { get; } = new List<QuickDimensionWallAxisCandidateTrace>();
    }

    /// <summary>
    /// One resolved wall-end anchor snapshot captured from the spike anchor result for the XML log.
    /// </summary>
    public sealed class QuickDimensionWallAxisAnchorTrace
    {
        public string Label { get; set; } = string.Empty;
        public double StationOnWallAxis { get; set; }
        public XYZ? Point { get; set; }
        public Reference? EdgeReference { get; set; }
        public bool HasReference => EdgeReference != null;
    }

    /// <summary>
    /// One visible candidate wall observed along the selected wall axis, with its classification and provenance.
    /// </summary>
    public sealed class QuickDimensionWallAxisCandidateTrace
    {
        public long CandidateWallId { get; set; }
        public string CandidateTypeName { get; set; } = string.Empty;
        public QuickDimensionWallMidRunRelation Relation { get; set; }
        public bool InElementsAtJoinStart { get; set; }
        public bool InElementsAtJoinEnd { get; set; }
        public bool InGeometryJoin { get; set; }
        public bool IsPerpendicular { get; set; }
        public bool IsParallel { get; set; }
        public int ReferenceHitCount { get; set; }
        public int AcceptedMidRunStationCount { get; set; }
        public double FallbackStationOnSelectedAxis { get; set; }
        public double FallbackDistanceToSideLine { get; set; }
        public string RejectedReason { get; set; } = string.Empty;
        public List<QuickDimensionWallAxisReferenceHitTrace> ReferenceHits { get; } = new List<QuickDimensionWallAxisReferenceHitTrace>();
    }

    /// <summary>
    /// One distinct vertical-edge reference hit observed on the selected side line, with accept/reject provenance.
    /// </summary>
    public sealed class QuickDimensionWallAxisReferenceHitTrace
    {
        public int Index { get; set; }
        public double StationOnSelectedAxis { get; set; }
        public double DistanceToSideLine { get; set; }
        public bool CandidateReferenceNormalAlongAxis { get; set; }
        public bool SelectedWallExposesRefAtStation { get; set; }
        public bool SelectedReferenceNormalAlongAxis { get; set; }
        public bool Accepted { get; set; }
        public string RejectedReason { get; set; } = string.Empty;
        public Reference? EdgeReference { get; set; }
        public XYZ? Point { get; set; }
    }
}
