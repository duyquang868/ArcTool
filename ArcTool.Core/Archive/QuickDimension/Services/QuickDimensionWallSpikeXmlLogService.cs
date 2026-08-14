using System;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Xml.Linq;
using ArcTool.Core.Archive.QuickDimension.Models;
using Autodesk.Revit.DB;
using RevitView = Autodesk.Revit.DB.View;

namespace ArcTool.Core.Archive.QuickDimension.Services
{
    /// <summary>
    /// Writes a read-only XML smoke log for the isolated Quick Dimension wall-reference spike.
    /// The log intentionally uses the same boundary-candidate collector as the spike logic so
    /// failed Left/Right smoke cases can be audited against the exact candidate model in code.
    /// </summary>
    public static class QuickDimensionWallSpikeXmlLogService
    {
        private const string DateTimeFormat = "yyyy-MM-ddTHH:mm:ss.fffK";

        public static string WriteWallSpikeLog(
            Document doc,
            Wall selectedWall,
            XYZ sidePickPoint,
            QuickDimensionWallSpikeResult probeResult)
        {
            if (doc == null) throw new ArgumentNullException(nameof(doc));
            if (selectedWall == null) throw new ArgumentNullException(nameof(selectedWall));
            if (sidePickPoint == null) throw new ArgumentNullException(nameof(sidePickPoint));
            if (probeResult == null) throw new ArgumentNullException(nameof(probeResult));

            if (string.IsNullOrWhiteSpace(doc.PathName))
            {
                throw new InvalidOperationException("The Revit document must be saved before writing the wall spike XML log next to the .rvt file.");
            }

            if (selectedWall.Location is not LocationCurve locationCurve || locationCurve.Curve is not Line wallLine)
            {
                throw new InvalidOperationException("The selected wall does not expose a straight LocationCurve line for wall spike XML logging.");
            }

            XYZ wallStart = wallLine.GetEndPoint(0);
            XYZ wallEnd = wallLine.GetEndPoint(1);
            XYZ wallDirection = GetPlanarDirection(wallStart, wallEnd);

            ShellLayerType selectedShellLayer = probeResult.SelectedShellLayer ?? ShellLayerType.Exterior;
            IReadOnlyList<QuickDimensionWallSpikeCornerProbePoint> selectedCorners =
                QuickDimensionWallReferenceProbeService.CollectBoundaryCornerPointsForLog(
                    selectedWall,
                    wallStart,
                    wallDirection,
                    selectedShellLayer,
                    includeBothShells: false,
                    out string selectedFailureReason);

            List<Wall> joinedWalls = QuickDimensionWallReferenceProbeService.CollectJoinedWalls(selectedWall);

            XElement root = new XElement(
                "QuickDimensionWallSpikeLog",
                new XAttribute("createdAt", DateTimeOffset.Now.ToString(DateTimeFormat, CultureInfo.InvariantCulture)),
                new XAttribute("coordinateBasis", "Survey coordinates from Document.ActiveProjectLocation.GetProjectPosition"),
                new XAttribute("coordinateUnits", "meters primary; millimeters included as *_mm attributes"),
                new XAttribute("northField", "ProjectPosition.NorthSouth"),
                new XAttribute("eastField", "ProjectPosition.EastWest"),
                new XAttribute("cornerDefinition", "Current spike boundary candidates from side-face edges: vertical-edge XY midpoints and horizontal-endpoints"),
                new XElement(
                    "ProbeResult",
                    new XAttribute("succeeded", probeResult.Succeeded),
                    new XAttribute("side", probeResult.Side),
                    new XAttribute("selectedShellLayer", probeResult.SelectedShellLayer?.ToString() ?? string.Empty),
                    new XAttribute("message", probeResult.Message ?? string.Empty),
                    BuildAnchorElement(doc, "StartAnchor", probeResult.StartAnchor),
                    BuildAnchorElement(doc, "FinishAnchor", probeResult.FinishAnchor)),
                new XElement(
                    "SelectedWall",
                    new XAttribute("id", selectedWall.Id.Value),
                    new XAttribute("typeName", selectedWall.WallType?.Name ?? string.Empty),
                    new XAttribute("cornerCount", selectedCorners.Count),
                    new XAttribute("failureReason", selectedFailureReason ?? string.Empty),
                    BuildPointElement(doc, "LocationCurveStart", wallStart, null, "location-curve"),
                    BuildPointElement(doc, "LocationCurveEnd", wallEnd, null, "location-curve"),
                    new XElement(
                        "Corners",
                        selectedCorners.Select((corner, index) => BuildCornerElement(doc, index + 1, corner)))));

            XElement joinedWallsElement = new XElement("JoinedWalls", new XAttribute("count", joinedWalls.Count));
            foreach (Wall joinedWall in joinedWalls)
            {
                IReadOnlyList<QuickDimensionWallSpikeCornerProbePoint> joinedCorners =
                    QuickDimensionWallReferenceProbeService.CollectBoundaryCornerPointsForLog(
                        joinedWall,
                        wallStart,
                        wallDirection,
                        ShellLayerType.Exterior,
                        includeBothShells: true,
                        out string joinedFailureReason);

                joinedWallsElement.Add(
                    new XElement(
                        "JoinedWall",
                        new XAttribute("id", joinedWall.Id.Value),
                        new XAttribute("typeName", joinedWall.WallType?.Name ?? string.Empty),
                        new XAttribute("cornerCount", joinedCorners.Count),
                        new XAttribute("failureReason", joinedFailureReason ?? string.Empty),
                        new XElement(
                            "Corners",
                            joinedCorners.Select((corner, index) => BuildCornerElement(doc, index + 1, corner)))));
            }

            root.Add(joinedWallsElement);

            // Session 2.7 Section 11: read-only evidence only. This block never builds an ordered chain,
            // retains no Reference, and never creates a dimension.
            RevitView midRunView = doc.ActiveView;
            XElement midRunElement = midRunView == null
                ? new XElement(
                    "MidRunProbe",
                    new XAttribute("supported", false),
                    new XAttribute("message", "Document active view is null; mid-run probe skipped."))
                : BuildMidRunProbeElement(
                    doc,
                    QuickDimensionWallMidRunProbeService.Probe(doc, midRunView, selectedWall, sidePickPoint, probeResult));
            root.Add(midRunElement);

            string directory = Path.GetDirectoryName(doc.PathName);
            if (string.IsNullOrWhiteSpace(directory))
            {
                throw new InvalidOperationException("Could not resolve the folder of the active .rvt file.");
            }

            string safeTimestamp = DateTime.Now.ToString("yyyyMMdd_HHmmss", CultureInfo.InvariantCulture);
            string fileName = $"ArcTool_QD_WallSpike_{selectedWall.Id.Value}_{safeTimestamp}.xml";
            string path = Path.Combine(directory, fileName);

            XDocument document = new XDocument(new XDeclaration("1.0", "utf-8", "yes"), root);
            document.Save(path);
            return path;
        }

        private static XElement BuildMidRunProbeElement(Document doc, QuickDimensionWallMidRunProbeResult result)
        {
            XElement element = new XElement(
                "MidRunProbe",
                new XAttribute("supported", result.Supported),
                new XAttribute("side", result.Side),
                new XAttribute("shell", result.Shell?.ToString() ?? string.Empty),
                new XAttribute("axisLengthMm", FormatMm(UnitUtils.ConvertFromInternalUnits(result.AxisLength, UnitTypeId.Millimeters))),
                new XAttribute("candidateCount", result.Candidates.Count),
                new XAttribute("message", result.Message ?? string.Empty),
                BuildJoinProvenanceElement("ElementsAtJoinStart", result.ElementsAtJoinStartIds),
                BuildJoinProvenanceElement("ElementsAtJoinEnd", result.ElementsAtJoinEndIds),
                BuildJoinProvenanceElement("GeometryJoin", result.GeometryJoinIds));

            XElement candidatesElement = new XElement(
                "Candidates",
                result.Candidates.Select((candidate, index) => BuildMidRunCandidateElement(doc, index + 1, candidate)));
            element.Add(candidatesElement);
            return element;
        }

        private static XElement BuildJoinProvenanceElement(string name, IReadOnlyList<long> ids)
        {
            return new XElement(
                name,
                new XAttribute("count", ids?.Count ?? 0),
                new XAttribute("ids", ids == null ? string.Empty : string.Join(",", ids)));
        }

        private static XElement BuildMidRunCandidateElement(
            Document doc,
            int index,
            QuickDimensionWallMidRunCandidate candidate)
        {
            XElement element = new XElement(
                "Candidate",
                new XAttribute("index", index),
                new XAttribute("wallId", candidate.CandidateWallId),
                new XAttribute("typeName", candidate.CandidateTypeName ?? string.Empty),
                new XAttribute("relation", candidate.Relation),
                new XAttribute("inElementsAtJoinStart", candidate.InElementsAtJoinStart),
                new XAttribute("inElementsAtJoinEnd", candidate.InElementsAtJoinEnd),
                new XAttribute("inGeometryJoin", candidate.InGeometryJoin),
                new XAttribute("isPerpendicular", candidate.IsPerpendicular),
                new XAttribute("isParallel", candidate.IsParallel),
                new XAttribute("referenceHitCount", candidate.ReferenceHits.Count),
                new XAttribute("acceptedMidRunStationCount", candidate.AcceptedMidRunStationCount),
                new XAttribute("candidateWallExposesRefAtStation", candidate.CandidateWallExposesRefAtStation),
                new XAttribute("candidateReferenceNormalAlongAxis", candidate.CandidateReferenceNormalAlongAxis),
                new XAttribute("selectedWallExposesRefAtStation", candidate.SelectedWallExposesRefAtStation),
                new XAttribute("selectedReferenceNormalAlongAxis", candidate.SelectedReferenceNormalAlongAxis),
                new XAttribute("fallbackStationMm", FormatMm(UnitUtils.ConvertFromInternalUnits(candidate.FallbackStationOnSelectedAxis, UnitTypeId.Millimeters))),
                new XAttribute("fallbackDistanceToSideLineMm", FormatMm(UnitUtils.ConvertFromInternalUnits(candidate.FallbackDistanceToSideLine, UnitTypeId.Millimeters))),
                new XAttribute("source", candidate.SourceLabel ?? string.Empty));

            XElement referenceHitsElement = new XElement(
                "ReferenceHits",
                candidate.ReferenceHits.Select((hit, hitIndex) => BuildMidRunReferenceHitElement(doc, hitIndex + 1, hit)));
            element.Add(referenceHitsElement);
            return element;
        }

        private static XElement BuildMidRunReferenceHitElement(
            Document doc,
            int index,
            QuickDimensionWallMidRunReferenceHit hit)
        {
            XElement element = new XElement(
                "ReferenceHit",
                new XAttribute("index", index),
                new XAttribute("stationMm", FormatMm(UnitUtils.ConvertFromInternalUnits(hit.StationOnSelectedAxis, UnitTypeId.Millimeters))),
                new XAttribute("distanceToSideLineMm", FormatMm(UnitUtils.ConvertFromInternalUnits(hit.DistanceToSideLine, UnitTypeId.Millimeters))),
                new XAttribute("candidateReferenceNormalAlongAxis", hit.CandidateReferenceNormalAlongAxis),
                new XAttribute("selectedWallExposesRefAtStation", hit.SelectedWallExposesRefAtStation),
                new XAttribute("selectedReferenceNormalAlongAxis", hit.SelectedReferenceNormalAlongAxis));

            if (hit.Midpoint != null)
            {
                element.Add(BuildPointElement(
                    doc,
                    "Point",
                    hit.Midpoint,
                    hit.StationOnSelectedAxis,
                    "vertical-edge-on-side-line"));
            }

            return element;
        }

        private static XElement BuildAnchorElement(Document doc, string name, QuickDimensionWallSpikeAnchor anchor)
        {
            XElement element = new XElement(name);
            if (anchor == null)
            {
                element.SetAttributeValue("resolved", false);
                return element;
            }

            element.SetAttributeValue("resolved", true);
            element.SetAttributeValue("label", anchor.Label ?? string.Empty);
            element.SetAttributeValue("stationMm", FormatMm(UnitUtils.ConvertFromInternalUnits(anchor.ParameterOnWallAxis, UnitTypeId.Millimeters)));
            element.Add(BuildPointElement(doc, "Point", anchor.Midpoint, anchor.ParameterOnWallAxis, "resolved-anchor"));
            return element;
        }

        private static XElement BuildCornerElement(Document doc, int index, QuickDimensionWallSpikeCornerProbePoint corner)
        {
            XElement element = BuildPointElement(doc, "Corner", corner.Point, corner.ParameterOnSelectedWallAxis, corner.Source);
            element.SetAttributeValue("index", index);
            element.SetAttributeValue("sourceWallId", corner.SourceWallId);
            return element;
        }

        private static XElement BuildPointElement(Document doc, string name, XYZ point, double? station, string source)
        {
            // ToSharedMm returns survey coordinates in millimeters; the smoke-test labels read in
            // meters (N/E), so emit meters as the primary N/E attributes and keep mm for precision.
            ConvertedCoordinate coordinate = CoordinateConversionService.ToSharedMm(doc, point);

            XElement element = new XElement(
                name,
                new XAttribute("source", source ?? string.Empty),
                new XAttribute("n", FormatMeters(coordinate.NorthSouthMm)),
                new XAttribute("e", FormatMeters(coordinate.EastWestMm)),
                new XAttribute("elevation", FormatMeters(coordinate.ElevationMm)),
                new XAttribute("n_mm", FormatMm(coordinate.NorthSouthMm)),
                new XAttribute("e_mm", FormatMm(coordinate.EastWestMm)),
                new XAttribute("elevation_mm", FormatMm(coordinate.ElevationMm)));

            if (station.HasValue)
            {
                element.SetAttributeValue("stationOnSelectedWallAxisMm", FormatMm(UnitUtils.ConvertFromInternalUnits(station.Value, UnitTypeId.Millimeters)));
            }

            return element;
        }

        private static XYZ GetPlanarDirection(XYZ start, XYZ end)
        {
            double dx = end.X - start.X;
            double dy = end.Y - start.Y;
            double length = Math.Sqrt((dx * dx) + (dy * dy));
            if (length <= 1e-9)
            {
                throw new InvalidOperationException("The selected wall LocationCurve is too short for wall spike XML logging.");
            }

            return new XYZ(dx / length, dy / length, 0.0);
        }

        private static string FormatMm(double value)
        {
            return value.ToString("0.###", CultureInfo.InvariantCulture);
        }

        private static string FormatMeters(double millimeters)
        {
            return (millimeters / 1000.0).ToString("0.###", CultureInfo.InvariantCulture);
        }
    }
}
