#nullable enable
using System;
using System.Collections.Generic;
using System.Linq;
using ArcTool.Core.Archive.QuickDimension.Models;
using Autodesk.Revit.DB;
using RevitView = Autodesk.Revit.DB.View;

namespace ArcTool.Core.Archive.QuickDimension.Services
{
    /// <summary>
    /// Wall-only read-only aggregator for Quick Dimension wall-axis projection model.
    /// Collects accepted mid-run joining-wall reference hits along the selected wall axis,
    /// applies ADR-2026-07-19A gates, and preserves live Revit References for Phase 3.
    /// Performs no transactions and never creates dimensions.
    /// </summary>
    public static class QuickDimensionWallAxisAggregatorService
    {
        private const double MinimumWallAxisLength = 1e-6;
        private const double VerticalEdgeTolerance = 1e-3;
        private static readonly double SideLineTolerance = UnitUtils.ConvertToInternalUnits(5.0, UnitTypeId.Millimeters);
        private static readonly double StationEps = UnitUtils.ConvertToInternalUnits(5.0, UnitTypeId.Millimeters);
        private const double PerpendicularDotTolerance = 0.1;
        private const double ParallelDotTolerance = 0.9;
        private const double NormalAlongAxisDotTolerance = 0.9;

        /// <summary>
        /// Collects mid-run joining-wall candidates along the selected wall axis.
        /// Returns reference-preserving candidates with live Edge.Reference, station, hit point, and diagnostics.
        /// When <paramref name="trace"/> is supplied it is populated with per-candidate classification and
        /// per-hit accept/reject provenance so the read-only XML log can audit the logic. The classifier lives
        /// here; the trace only records the decision (no duplicate classification in the XML log service).
        /// </summary>
        public static List<QuickDimensionCandidate> CollectMidRunCandidates(
            Document doc,
            RevitView view,
            Wall selectedWall,
            QuickDimensionWallSpikeResult anchorResult,
            QuickDimensionLineContext lineContext,
            List<QuickDimensionDiagnostic> diagnostics,
            QuickDimensionWallAxisAggregationTrace? trace = null)
        {
            if (doc == null) throw new ArgumentNullException(nameof(doc));
            if (view == null) throw new ArgumentNullException(nameof(view));
            if (selectedWall == null) throw new ArgumentNullException(nameof(selectedWall));
            if (anchorResult == null) throw new ArgumentNullException(nameof(anchorResult));
            if (lineContext == null) throw new ArgumentNullException(nameof(lineContext));
            if (diagnostics == null) throw new ArgumentNullException(nameof(diagnostics));

            var result = new List<QuickDimensionCandidate>();

            if (trace != null)
            {
                trace.SelectedWallId = selectedWall.Id.Value;
                trace.SideSign = lineContext.SideSign;
                trace.SideLabel = lineContext.SideSign > 0 ? "Left" : lineContext.SideSign < 0 ? "Right" : "Unspecified";
            }

            if (!anchorResult.Succeeded || anchorResult.StartAnchor == null || anchorResult.FinishAnchor == null)
            {
                diagnostics.Add(new QuickDimensionDiagnostic(
                    QuickDimensionDiagnosticSeverity.Info,
                    QuickDimensionRejectedReason.None,
                    "Mid-run aggregator skipped: wall anchors are unavailable.",
                    selectedWall.Id,
                    QuickDimensionSourceType.Wall));
                SetTraceMessage(trace, false, "Mid-run aggregator skipped: wall anchors are unavailable.");
                return result;
            }

            if (selectedWall.Location is not LocationCurve locationCurve || locationCurve.Curve is not Line wallLine)
            {
                diagnostics.Add(new QuickDimensionDiagnostic(
                    QuickDimensionDiagnosticSeverity.Warning,
                    QuickDimensionRejectedReason.InvalidGeometry,
                    "Mid-run aggregator skipped: selected wall does not expose a straight LocationCurve line.",
                    selectedWall.Id,
                    QuickDimensionSourceType.Wall));
                SetTraceMessage(trace, false, "Mid-run aggregator skipped: selected wall does not expose a straight LocationCurve line.");
                return result;
            }

            XYZ wallStart = wallLine.GetEndPoint(0);
            XYZ wallEnd = wallLine.GetEndPoint(1);
            if (!TryGetPlanarDirection(wallStart, wallEnd, out XYZ wallDirection, out double axisLength))
            {
                SetTraceMessage(trace, false, "Mid-run aggregator skipped: selected wall axis is too short or invalid.");
                return result;
            }

            double resolvedStartStation = anchorResult.StartAnchor.ParameterOnWallAxis;
            double resolvedFinishStation = anchorResult.FinishAnchor.ParameterOnWallAxis;
            double resolvedAnchorMin = Math.Min(resolvedStartStation, resolvedFinishStation);
            double resolvedAnchorMax = Math.Max(resolvedStartStation, resolvedFinishStation);
            if (resolvedAnchorMax - resolvedAnchorMin <= StationEps * 2.0)
            {
                SetTraceMessage(trace, false, "Mid-run aggregator skipped: resolved wall anchor span is too short.");
                return result;
            }

            XYZ? nullableSideNormal = lineContext.SideNormal;
            if (nullableSideNormal == null)
            {
                SetTraceMessage(trace, false, "Mid-run aggregator skipped: side normal is unspecified.");
                return result;
            }

            XYZ sideNormal = nullableSideNormal;
            string sideLabel = lineContext.SideSign > 0 ? "Left" : "Right";
            ShellLayerType shellLayer = GetSelectedShellLayer(selectedWall, sideNormal);

            double halfWidth = 0.0;
            try { halfWidth = selectedWall.Width * 0.5; } catch { halfWidth = 0.0; }
            XYZ sideLinePoint = new XYZ(
                wallStart.X + (sideNormal.X * halfWidth),
                wallStart.Y + (sideNormal.Y * halfWidth),
                wallStart.Z);

            List<long> startJoinIds = CollectWallIdsFromJoin(locationCurve, 0);
            List<long> endJoinIds = CollectWallIdsFromJoin(locationCurve, 1);
            var startJoinSet = new HashSet<long>(startJoinIds);
            var endJoinSet = new HashSet<long>(endJoinIds);

            // Geometry-join provenance is diagnostic ONLY (never gates acceptance); collect it once for the trace.
            var geometryJoinSet = trace != null
                ? new HashSet<long>(CollectGeometryJoinIds(doc, selectedWall))
                : new HashSet<long>();

            // Selected-wall side-line evidence is diagnostic ONLY. It is captured once (not twice) and reused
            // to fill SelectedWallExposesRefAtStation per hit, matching the proven Section 11 probe fields.
            List<MidRunReferenceHit> selectedVerticalHits = trace != null
                ? CollectVerticalReferenceHits(doc, TryGetWallGeometry(selectedWall), wallStart, wallDirection, sideLinePoint)
                : new List<MidRunReferenceHit>();

            if (trace != null)
            {
                trace.Supported = true;
                trace.Message = "Mid-run aggregator ran.";
                trace.ShellLayer = shellLayer;
                trace.AxisLength = axisLength;
                trace.SideNormal = sideNormal;
                trace.SideLinePoint = sideLinePoint;
                trace.ResolvedAnchorMinStation = resolvedAnchorMin;
                trace.ResolvedAnchorMaxStation = resolvedAnchorMax;
                trace.StartAnchor = BuildAnchorTrace(anchorResult.StartAnchor);
                trace.FinishAnchor = BuildAnchorTrace(anchorResult.FinishAnchor);
                trace.ElementsAtJoinStartIds.AddRange(startJoinIds);
                trace.ElementsAtJoinEndIds.AddRange(endJoinIds);
                trace.GeometryJoinIds.AddRange(geometryJoinSet);
            }

            List<Wall> candidateWalls = new FilteredElementCollector(doc, view.Id)
                .OfClass(typeof(Wall))
                .OfCategory(BuiltInCategory.OST_Walls)
                .Cast<Wall>()
                .Where(w => w != null
                            && w.IsValidObject
                            && w.Id.Value != selectedWall.Id.Value
                            && w.WallType?.Kind != WallKind.Curtain
                            && w.Location is LocationCurve lc && lc.Curve is Line)
                .ToList();

            foreach (Wall candidate in candidateWalls)
            {
                long cid = candidate.Id.Value;
                bool inStart = startJoinSet.Contains(cid);
                bool inEnd = endJoinSet.Contains(cid);
                bool inGeom = geometryJoinSet.Contains(cid);

                // PRODUCTION GATE (unchanged): end-join members can never become mid-run crossings.
                if (inStart || inEnd)
                {
                    diagnostics.Add(new QuickDimensionDiagnostic(
                        QuickDimensionDiagnosticSeverity.Info,
                        QuickDimensionRejectedReason.None,
                        $"Mid-run candidate Wall {cid} excluded: end-join member.",
                        candidate.Id,
                        QuickDimensionSourceType.Wall));

                    // For the audit trail only, still record end-join artifacts (they may expose side-line refs).
                    if (trace != null)
                    {
                        RecordCandidateTrace(
                            trace, doc, candidate, wallStart, wallDirection, axisLength, sideLinePoint,
                            selectedVerticalHits, null, resolvedStartStation, resolvedFinishStation,
                            resolvedAnchorMin, resolvedAnchorMax, inStart, inEnd, inGeom,
                            "end-join member");
                    }
                    continue;
                }

                GeometryElement? candidateGeometry = TryGetWallGeometry(candidate);
                List<MidRunReferenceHit> hits = CollectVerticalReferenceHits(doc, candidateGeometry, wallStart, wallDirection, sideLinePoint);

                List<MidRunReferenceHit> acceptedHits = DedupeHitsByStation(hits
                    .Where(hit => PassesMidRunHitGates(
                        hit,
                        resolvedStartStation,
                        resolvedFinishStation,
                        resolvedAnchorMin,
                        resolvedAnchorMax,
                        axisLength,
                        inStart,
                        inEnd)));

                if (acceptedHits.Count < 2)
                {
                    if (hits.Count > 0)
                    {
                        diagnostics.Add(new QuickDimensionDiagnostic(
                            QuickDimensionDiagnosticSeverity.Info,
                            QuickDimensionRejectedReason.None,
                            $"Mid-run candidate Wall {cid} excluded: accepted station count {acceptedHits.Count} < 2.",
                            candidate.Id,
                            QuickDimensionSourceType.Wall));
                    }

                    if (trace != null)
                    {
                        RecordCandidateTrace(
                            trace, doc, candidate, wallStart, wallDirection, axisLength, sideLinePoint,
                            selectedVerticalHits, hits, resolvedStartStation, resolvedFinishStation,
                            resolvedAnchorMin, resolvedAnchorMax, inStart, inEnd, inGeom,
                            $"accepted distinct station count {acceptedHits.Count} < 2");
                    }
                    continue;
                }

                string displayName = $"Wall: {candidate.WallType?.Name ?? string.Empty}";
                string typeName = candidate.WallType?.Name ?? string.Empty;
                string shellLabel = shellLayer == ShellLayerType.Exterior ? "Exterior" : "Interior";

                foreach (MidRunReferenceHit hit in acceptedHits)
                {
                    result.Add(new QuickDimensionCandidate(
                        candidate.Id,
                        QuickDimensionSourceType.Wall,
                        $"{displayName} [{shellLabel}/{sideLabel} Mid-Run @ {UnitUtils.ConvertFromInternalUnits(hit.Station, UnitTypeId.Millimeters):0.##} mm]",
                        hit.Reference,
                        QuickDimensionReferenceStrategy.WallSideFace,
                        hit.Midpoint,
                        hit.Station,
                        selectedWall.Id,
                        typeName: typeName));
                }

                diagnostics.Add(new QuickDimensionDiagnostic(
                    QuickDimensionDiagnosticSeverity.Info,
                    QuickDimensionRejectedReason.None,
                    $"Accepted mid-run Wall {cid} with {acceptedHits.Count} reference hit(s).",
                    candidate.Id,
                    QuickDimensionSourceType.Wall));

                if (trace != null)
                {
                    RecordCandidateTrace(
                        trace, doc, candidate, wallStart, wallDirection, axisLength, sideLinePoint,
                        selectedVerticalHits, hits, resolvedStartStation, resolvedFinishStation,
                        resolvedAnchorMin, resolvedAnchorMax, inStart, inEnd, inGeom,
                        string.Empty);
                }
            }

            return result;
        }

        private static void SetTraceMessage(QuickDimensionWallAxisAggregationTrace? trace, bool supported, string message)
        {
            if (trace == null)
            {
                return;
            }

            trace.Supported = supported;
            trace.Message = message;
        }

        private static QuickDimensionWallAxisAnchorTrace BuildAnchorTrace(QuickDimensionWallSpikeAnchor anchor)
        {
            return new QuickDimensionWallAxisAnchorTrace
            {
                Label = anchor.Label,
                StationOnWallAxis = anchor.ParameterOnWallAxis,
                Point = anchor.Midpoint,
                EdgeReference = anchor.EdgeReference
            };
        }

        /// <summary>
        /// Records one candidate wall into the trace. Classification mirrors the proven Section 11 probe:
        /// end-join membership is exclusion evidence, MidRunCrossing requires >= 2 accepted distinct stations,
        /// and join/angle facts are provenance only (never gate a real reference-evidence crossing).
        /// </summary>
        private static void RecordCandidateTrace(
            QuickDimensionWallAxisAggregationTrace trace,
            Document doc,
            Wall candidate,
            XYZ wallStart,
            XYZ wallDirection,
            double axisLength,
            XYZ sideLinePoint,
            List<MidRunReferenceHit> selectedVerticalHits,
            List<MidRunReferenceHit>? preCollectedHits,
            double resolvedStartStation,
            double resolvedFinishStation,
            double resolvedAnchorMin,
            double resolvedAnchorMax,
            bool inStart,
            bool inEnd,
            bool inGeom,
            string rejectedReason)
        {
            var candidateTrace = new QuickDimensionWallAxisCandidateTrace
            {
                CandidateWallId = candidate.Id.Value,
                CandidateTypeName = candidate.WallType?.Name ?? string.Empty,
                InElementsAtJoinStart = inStart,
                InElementsAtJoinEnd = inEnd,
                InGeometryJoin = inGeom,
                RejectedReason = rejectedReason
            };

            Line? candidateLine = (candidate.Location as LocationCurve)?.Curve as Line;
            if (candidateLine != null &&
                TryGetPlanarDirection(candidateLine.GetEndPoint(0), candidateLine.GetEndPoint(1), out XYZ candidateDirection, out _))
            {
                double directionDot = Math.Abs(candidateDirection.DotProduct(wallDirection));
                candidateTrace.IsPerpendicular = directionDot < PerpendicularDotTolerance;
                candidateTrace.IsParallel = directionDot > ParallelDotTolerance;

                XYZ fallbackPoint = NearestEndpointToAxis(candidateLine, wallStart, wallDirection);
                candidateTrace.FallbackStationOnSelectedAxis = ProjectStation(fallbackPoint, wallStart, wallDirection);
                candidateTrace.FallbackDistanceToSideLine = DistanceToSideLine(fallbackPoint, sideLinePoint, wallDirection);
            }

            // Reuse the hits already scanned by the acceptance pass to avoid a second geometry scan;
            // only end-join artifacts (skipped before acceptance) need their own scan here.
            List<MidRunReferenceHit> hits = preCollectedHits
                ?? CollectVerticalReferenceHits(doc, TryGetWallGeometry(candidate), wallStart, wallDirection, sideLinePoint);

            var seenAcceptedStations = new List<double>();
            int index = 0;
            foreach (MidRunReferenceHit hit in hits.OrderBy(h => h.Station))
            {
                index++;
                bool baseAccepted = PassesMidRunHitGates(
                    hit,
                    resolvedStartStation,
                    resolvedFinishStation,
                    resolvedAnchorMin,
                    resolvedAnchorMax,
                    axisLength,
                    inStart,
                    inEnd);
                bool duplicate = baseAccepted && seenAcceptedStations.Any(s => Math.Abs(s - hit.Station) <= StationEps);

                bool accepted = baseAccepted && !duplicate;
                string hitReason = accepted
                    ? string.Empty
                    : duplicate
                        ? "duplicate station on this candidate"
                        : GetMidRunHitRejectedReason(
                            hit,
                            resolvedStartStation,
                            resolvedFinishStation,
                            resolvedAnchorMin,
                            resolvedAnchorMax,
                            axisLength,
                            inStart,
                            inEnd);

                if (accepted)
                {
                    seenAcceptedStations.Add(hit.Station);
                }

                MidRunReferenceHit? selectedHit = selectedVerticalHits
                    .Where(s => Math.Abs(s.Station - hit.Station) <= StationEps)
                    .OrderBy(s => s.Distance)
                    .FirstOrDefault();

                candidateTrace.ReferenceHits.Add(new QuickDimensionWallAxisReferenceHitTrace
                {
                    Index = index,
                    StationOnSelectedAxis = hit.Station,
                    DistanceToSideLine = hit.Distance,
                    CandidateReferenceNormalAlongAxis = hit.NormalAlongAxis,
                    SelectedWallExposesRefAtStation = selectedHit != null,
                    SelectedReferenceNormalAlongAxis = selectedHit?.NormalAlongAxis ?? false,
                    Accepted = accepted,
                    RejectedReason = hitReason,
                    EdgeReference = hit.Reference,
                    Point = hit.Midpoint
                });
            }

            candidateTrace.ReferenceHitCount = hits.Count;
            candidateTrace.AcceptedMidRunStationCount = seenAcceptedStations.Count;
            candidateTrace.Relation = ClassifyRelation(
                candidateTrace.IsParallel,
                candidateTrace.ReferenceHitCount,
                candidateTrace.AcceptedMidRunStationCount,
                inStart,
                inEnd,
                inGeom);

            trace.Candidates.Add(candidateTrace);
        }

        private static bool PassesMidRunHitGates(
            MidRunReferenceHit hit,
            double resolvedStartStation,
            double resolvedFinishStation,
            double resolvedAnchorMin,
            double resolvedAnchorMax,
            double axisLength,
            bool inStart,
            bool inEnd)
        {
            return !inStart
                && !inEnd
                && hit.NormalAlongAxis
                && hit.Station > resolvedAnchorMin + StationEps
                && hit.Station < resolvedAnchorMax - StationEps
                && Math.Abs(hit.Station - resolvedStartStation) > StationEps
                && Math.Abs(hit.Station - resolvedFinishStation) > StationEps
                && hit.Station > StationEps
                && hit.Station < axisLength - StationEps;
        }

        private static string GetMidRunHitRejectedReason(
            MidRunReferenceHit hit,
            double resolvedStartStation,
            double resolvedFinishStation,
            double resolvedAnchorMin,
            double resolvedAnchorMax,
            double axisLength,
            bool inStart,
            bool inEnd)
        {
            if (inStart || inEnd)
            {
                return "end-join member";
            }

            if (!hit.NormalAlongAxis)
            {
                return "normal not along axis";
            }

            if (hit.Station <= StationEps || hit.Station >= axisLength - StationEps)
            {
                return "outside (0, axisLength) span";
            }

            if (hit.Station <= resolvedAnchorMin + StationEps || hit.Station >= resolvedAnchorMax - StationEps)
            {
                return "outside resolved anchor span";
            }

            if (Math.Abs(hit.Station - resolvedStartStation) <= StationEps ||
                Math.Abs(hit.Station - resolvedFinishStation) <= StationEps)
            {
                return "coincident with start/finish anchor";
            }

            return "not accepted";
        }

        /// <summary>
        /// Mid-run relation classifier (single source of truth). MidRunCrossing means an accepted
        /// reference-evidence crossing (>= 2 accepted distinct stations), never Revit join topology.
        /// End-join membership is exclusion evidence per ADR-2026-07-18A.
        /// </summary>
        private static QuickDimensionWallMidRunRelation ClassifyRelation(
            bool isParallel,
            int referenceHitCount,
            int acceptedMidRunStationCount,
            bool inStart,
            bool inEnd,
            bool inGeom)
        {
            if (inStart || inEnd)
            {
                return QuickDimensionWallMidRunRelation.EndJoinOnly;
            }

            if (acceptedMidRunStationCount >= 2)
            {
                return QuickDimensionWallMidRunRelation.MidRunCrossing;
            }

            if (inGeom)
            {
                return QuickDimensionWallMidRunRelation.GeometryJoinOnly;
            }

            if (referenceHitCount > 0)
            {
                return QuickDimensionWallMidRunRelation.NonJoinedProximity;
            }

            return isParallel
                ? QuickDimensionWallMidRunRelation.ParallelNonJoined
                : QuickDimensionWallMidRunRelation.Ignored;
        }

        private static List<MidRunReferenceHit> CollectVerticalReferenceHits(
            Document doc,
            GeometryElement? geometry,
            XYZ wallStart,
            XYZ wallDirection,
            XYZ sideLinePoint)
        {
            var hits = new List<MidRunReferenceHit>();
            var seenKeys = new HashSet<string>();
            CollectVerticalReferenceHits(doc, geometry, wallStart, wallDirection, sideLinePoint, hits, seenKeys);
            return hits;
        }

        private static void CollectVerticalReferenceHits(
            Document doc,
            GeometryElement? geometry,
            XYZ wallStart,
            XYZ wallDirection,
            XYZ sideLinePoint,
            List<MidRunReferenceHit> hits,
            HashSet<string> seenKeys)
        {
            if (geometry == null)
            {
                return;
            }

            foreach (GeometryObject geometryObject in geometry)
            {
                if (geometryObject is Solid solid)
                {
                    ScanSolidForVerticalReferenceHits(doc, solid, wallStart, wallDirection, sideLinePoint, hits, seenKeys);
                }
                else if (geometryObject is GeometryInstance instance)
                {
                    CollectVerticalReferenceHits(doc, instance.GetInstanceGeometry(), wallStart, wallDirection, sideLinePoint, hits, seenKeys);
                }
            }
        }

        private static void ScanSolidForVerticalReferenceHits(
            Document doc,
            Solid solid,
            XYZ wallStart,
            XYZ wallDirection,
            XYZ sideLinePoint,
            List<MidRunReferenceHit> hits,
            HashSet<string> seenKeys)
        {
            if (solid?.Edges == null || solid.Edges.Size == 0)
            {
                return;
            }

            foreach (Edge edge in solid.Edges)
            {
                try
                {
                    if (edge.AsCurve() is not Line line || edge.Reference == null)
                    {
                        continue;
                    }

                    XYZ direction = line.Direction;
                    if (direction == null || Math.Abs(Math.Abs(direction.Z) - 1.0) > VerticalEdgeTolerance)
                    {
                        continue;
                    }

                    string stableKey = edge.Reference.ConvertToStableRepresentation(doc);
                    if (string.IsNullOrWhiteSpace(stableKey) || !seenKeys.Add(stableKey))
                    {
                        continue;
                    }

                    XYZ midpoint = (line.GetEndPoint(0) + line.GetEndPoint(1)) * 0.5;
                    double distance = DistanceToSideLine(midpoint, sideLinePoint, wallDirection);
                    double station = ProjectStation(midpoint, wallStart, wallDirection);
                    bool normalAlongAxis = EdgeFaceNormalAlongAxis(edge, wallDirection);

                    if (distance > SideLineTolerance)
                    {
                        continue;
                    }

                    hits.Add(new MidRunReferenceHit(
                        edge.Reference,
                        midpoint,
                        station,
                        distance,
                        normalAlongAxis));
                }
                catch
                {
                    // Skip malformed or non-persistable reference evidence.
                }
            }
        }

        private static bool EdgeFaceNormalAlongAxis(Edge edge, XYZ wallDirection)
        {
            Face? faceA = null;
            Face? faceB = null;
            try { faceA = edge.GetFace(0); } catch { }
            try { faceB = edge.GetFace(1); } catch { }

            return FaceNormalParallelToAxis(faceA, wallDirection) || FaceNormalParallelToAxis(faceB, wallDirection);
        }

        private static bool FaceNormalParallelToAxis(Face? face, XYZ wallDirection)
        {
            if (face is not PlanarFace planarFace || planarFace.FaceNormal == null)
            {
                return false;
            }

            XYZ planarNormal = new XYZ(planarFace.FaceNormal.X, planarFace.FaceNormal.Y, 0.0);
            if (planarNormal.GetLength() <= MinimumWallAxisLength)
            {
                return false;
            }

            return Math.Abs(planarNormal.Normalize().DotProduct(wallDirection)) > NormalAlongAxisDotTolerance;
        }

        private static XYZ NearestEndpointToAxis(Line candidateLine, XYZ axisPoint, XYZ axisDirection)
        {
            XYZ start = candidateLine.GetEndPoint(0);
            XYZ end = candidateLine.GetEndPoint(1);
            double distStart = DistanceToSideLine(start, axisPoint, axisDirection);
            double distEnd = DistanceToSideLine(end, axisPoint, axisDirection);
            return distStart <= distEnd ? start : end;
        }

        private static double DistanceToSideLine(XYZ point, XYZ linePoint, XYZ lineDirection)
        {
            XYZ offset = new XYZ(point.X - linePoint.X, point.Y - linePoint.Y, 0.0);
            return Math.Abs((offset.X * lineDirection.Y) - (offset.Y * lineDirection.X));
        }

        private static double ProjectStation(XYZ point, XYZ axisStart, XYZ axisDirection)
        {
            XYZ offset = new XYZ(point.X - axisStart.X, point.Y - axisStart.Y, 0.0);
            return offset.DotProduct(axisDirection);
        }

        private static List<long> CollectWallIdsFromJoin(LocationCurve locationCurve, int end)
        {
            var ids = new List<long>();
            try
            {
                ElementArray? joined = locationCurve.get_ElementsAtJoin(end);
                if (joined != null)
                {
                    foreach (Element element in joined)
                    {
                        if (element is Wall wall && wall.IsValidObject)
                        {
                            ids.Add(wall.Id.Value);
                        }
                    }
                }
            }
            catch
            {
                // Unsupported join state at this end.
            }

            return ids;
        }

        private static List<long> CollectGeometryJoinIds(Document doc, Wall selectedWall)
        {
            var ids = new List<long>();
            try
            {
                foreach (ElementId id in JoinGeometryUtils.GetJoinedElements(doc, selectedWall))
                {
                    if (doc.GetElement(id) is Wall wall && wall.IsValidObject)
                    {
                        ids.Add(wall.Id.Value);
                    }
                }
            }
            catch
            {
                // Geometry-join provenance is diagnostic only; unsupported states do not affect acceptance.
            }

            return ids;
        }

        private static ShellLayerType GetSelectedShellLayer(Wall wall, XYZ targetSideNormal)
        {
            XYZ? orientation = wall?.Orientation;
            if (orientation != null && targetSideNormal != null)
            {
                XYZ planarOrientation = new XYZ(orientation.X, orientation.Y, 0.0);
                if (planarOrientation.GetLength() > MinimumWallAxisLength &&
                    planarOrientation.Normalize().DotProduct(targetSideNormal) > 0.0)
                {
                    return ShellLayerType.Exterior;
                }
            }

            return ShellLayerType.Interior;
        }

        private static GeometryElement? TryGetWallGeometry(Wall wall)
        {
            try
            {
                Options options = new Options
                {
                    ComputeReferences = true,
                    IncludeNonVisibleObjects = true
                };
                return wall.get_Geometry(options);
            }
            catch
            {
                return null;
            }
        }

        private static bool TryGetPlanarDirection(XYZ start, XYZ end, out XYZ direction, out double length)
        {
            direction = null!;
            length = 0.0;
            if (start == null || end == null)
            {
                return false;
            }

            double dx = end.X - start.X;
            double dy = end.Y - start.Y;
            length = Math.Sqrt((dx * dx) + (dy * dy));
            if (length <= MinimumWallAxisLength)
            {
                return false;
            }

            direction = new XYZ(dx / length, dy / length, 0.0);
            return true;
        }

        private static List<MidRunReferenceHit> DedupeHitsByStation(IEnumerable<MidRunReferenceHit> hits)
        {
            var result = new List<MidRunReferenceHit>();
            var seenStations = new List<double>();

            foreach (MidRunReferenceHit hit in hits.OrderBy(h => h.Station))
            {
                bool isDuplicate = seenStations.Any(station => Math.Abs(station - hit.Station) <= StationEps);
                if (!isDuplicate)
                {
                    result.Add(hit);
                    seenStations.Add(hit.Station);
                }
            }

            return result;
        }

        private sealed class MidRunReferenceHit
        {
            public MidRunReferenceHit(
                Reference reference,
                XYZ midpoint,
                double station,
                double distance,
                bool normalAlongAxis)
            {
                Reference = reference ?? throw new ArgumentNullException(nameof(reference));
                Midpoint = midpoint ?? throw new ArgumentNullException(nameof(midpoint));
                Station = station;
                Distance = distance;
                NormalAlongAxis = normalAlongAxis;
            }

            public Reference Reference { get; }
            public XYZ Midpoint { get; }
            public double Station { get; }
            public double Distance { get; }
            public bool NormalAlongAxis { get; }
        }
    }
}
