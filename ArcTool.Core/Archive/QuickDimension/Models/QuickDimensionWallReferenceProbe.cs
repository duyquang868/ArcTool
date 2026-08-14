using Autodesk.Revit.DB;

namespace ArcTool.Core.Archive.QuickDimension.Models
{
    /// <summary>
    /// Placement side of the Quick Dimension wall spike relative to the selected wall axis.
    /// Left is the +90 degree (counter-clockwise) planar normal of the wall LocationCurve direction.
    /// </summary>
    public enum QuickDimensionWallSpikeSide
    {
        Unspecified = 0,
        Left = 1,
        Right = 2
    }

    /// <summary>
    /// A single wall end anchor discovered by the wall-axis spike.
    /// Anchors are the two vertical edges on the selected side face at min and max projected station along the wall axis.
    /// </summary>
    public sealed class QuickDimensionWallSpikeAnchor
    {
        public QuickDimensionWallSpikeAnchor(
            string label,
            Reference edgeReference,
            XYZ midpoint,
            double parameterOnWallAxis)
        {
            Label = label ?? string.Empty;
            EdgeReference = edgeReference;
            Midpoint = midpoint;
            ParameterOnWallAxis = parameterOnWallAxis;
        }

        public string Label { get; }
        public Reference EdgeReference { get; }
        public XYZ Midpoint { get; }
        public double ParameterOnWallAxis { get; }
    }

    /// <summary>
    /// A single boundary-candidate corner point captured for the XML smoke log.
    /// Mirrors the internal boundary-candidate model the spike uses so the log audits the real logic.
    /// </summary>
    public sealed class QuickDimensionWallSpikeCornerProbePoint
    {
        public QuickDimensionWallSpikeCornerProbePoint(
            XYZ point,
            double parameterOnSelectedWallAxis,
            long sourceWallId,
            string source)
        {
            Point = point;
            ParameterOnSelectedWallAxis = parameterOnSelectedWallAxis;
            SourceWallId = sourceWallId;
            Source = source ?? string.Empty;
        }

        /// <summary>Model-space point of the candidate corner in Revit internal units.</summary>
        public XYZ Point { get; }

        /// <summary>Projected station of the corner along the selected wall axis, in internal units.</summary>
        public double ParameterOnSelectedWallAxis { get; }

        /// <summary>ElementId value of the wall that owns this corner candidate.</summary>
        public long SourceWallId { get; }

        /// <summary>Origin tag of the candidate (vertical-edge, horizontal-endpoint, ...).</summary>
        public string Source { get; }
    }

    /// <summary>
    /// Final read-only summary of the wall-axis spike for a single selected wall + side pick.
    /// This spike collects only two wall anchors; it never creates a Revit dimension.
    /// </summary>
    public sealed class QuickDimensionWallSpikeResult
    {
        public QuickDimensionWallSpikeResult(
            bool succeeded,
            ElementId wallId,
            string wallTypeName,
            double wallAxisLength,
            QuickDimensionWallSpikeSide side,
            ShellLayerType? selectedShellLayer,
            int totalVerticalEdgesOnSide,
            QuickDimensionWallSpikeAnchor startAnchor,
            QuickDimensionWallSpikeAnchor finishAnchor,
            string message)
        {
            Succeeded = succeeded;
            WallId = wallId;
            WallTypeName = wallTypeName ?? string.Empty;
            WallAxisLength = wallAxisLength;
            Side = side;
            SelectedShellLayer = selectedShellLayer;
            TotalVerticalEdgesOnSide = totalVerticalEdgesOnSide;
            StartAnchor = startAnchor;
            FinishAnchor = finishAnchor;
            Message = message ?? string.Empty;
        }

        public bool Succeeded { get; }
        public ElementId WallId { get; }
        public string WallTypeName { get; }
        public double WallAxisLength { get; }
        public QuickDimensionWallSpikeSide Side { get; }
        public ShellLayerType? SelectedShellLayer { get; }
        public int TotalVerticalEdgesOnSide { get; }
        public QuickDimensionWallSpikeAnchor StartAnchor { get; }
        public QuickDimensionWallSpikeAnchor FinishAnchor { get; }
        public string Message { get; }
    }
}
