using System;
using System.Collections.Generic;
using System.Linq;
using ArcTool.Core.Archive.QuickDimension.Models;
using Autodesk.Revit.DB;
using RevitView = Autodesk.Revit.DB.View;

namespace ArcTool.Core.Archive.QuickDimension.Services
{
    /// <summary>
    /// Session 2.7 Section 11 LOG-ONLY, READ-ONLY mid-run probe.
    /// It observes how visible candidate walls relate to one selected straight wall axis so the XML smoke
    /// log can capture evidence for mid-run T-joints and non-joined proximity. It performs NO station
    /// aggregation, selects NO canonical station, builds NO ReferenceArray, and creates NO dimension.
    ///
    /// Provenance discipline: end-join sources are captured SEPARATELY (ElementsAtJoin start, ElementsAtJoin
    /// end, JoinGeometryUtils) so the log shows which mechanism, if any, exposes each candidate.
    ///
    /// Reference discipline: mid-run reference evidence is gathered by scanning candidate/selected wall
    /// geometry for vertical Edge.Reference midpoints that land on the selected side line inside the axis
    /// span. It is intentionally NOT derived from a candidate LocationCurve endpoint.
    /// </summary>
    public static class QuickDimensionWallMidRunProbeService
    {
        private const double MinimumWallAxisLength = 1e-6;
        private const double VerticalEdgeTolerance = 1e-3;

        // Reused spike conventions (see QuickDimensionWallReferenceProbeService).
        private static readonly double SideLineTolerance = UnitUtils.ConvertToInternalUnits(5.0, UnitTypeId.Millimeters);
        private static readonly double StationEps = UnitUtils.ConvertToInternalUnits(5.0, UnitTypeId.Millimeters);

        // Observational angle bands only. These are NOT part of the proven spike tolerance set and MUST be
        // validated by the Section 11 smoke run before any production use.
        private const double PerpendicularDotTolerance = 0.1;
        private const double ParallelDotTolerance = 0.9;
        private const double NormalAlongAxisDotTolerance = 0.9;

        public static QuickDimensionWallMidRunProbeResult Probe(
            Document doc,
            RevitView view,
            Wall selectedWall,
            XYZ sidePickPoint,
            QuickDimensionWallSpikeResult anchorResult)
        {
            if (doc == null) throw new ArgumentNullException(nameof(doc));
            if (view == null) throw new ArgumentNullException(nameof(view));
            if (selectedWall == null) throw new ArgumentNullException(nameof(selectedWall));
            if (sidePickPoint == null) throw new ArgumentNullException(nameof(sidePickPoint));
            if (anchorResult == null) throw new ArgumentNullException(nameof(anchorResult));

            long selectedId = selectedWall.Id.Value;

            if (!selectedWall.IsValidObject)
            {
                return Unsupported(selectedId, "Selected wall is not a valid Revit object.");
            }

            if (selectedWall.WallType?.Kind == WallKind.Curtain)
            {
                return Unsupported(selectedId, "Curtain walls are outside the mid-run probe scope.");
            }

            if (selectedWall.Location is not LocationCurve locationCurve || locationCurve.Curve is not Line wallLine)
            {
                return Unsupported(selectedId, "Selected wall does not expose a straight LocationCurve line.");
            }

            XYZ wallStart = wallLine.GetEndPoint(0);
            XYZ wallEnd = wallLine.GetEndPoint(1);
            if (!TryGetPlanarDirection(wallStart, wallEnd, out XYZ wallDirection, out double axisLength))
            {
                return Unsupported(selectedId, "Selected wall axis is too short or invalid.");
            }

            if (!anchorResult.Succeeded || anchorResult.StartAnchor == null || anchorResult.FinishAnchor == null)
            {
                return Unsupported(selectedId, "Resolved wall anchors are unavailable; mid-run acceptance skipped.");
            }

            double resolvedStartStation = anchorResult.StartAnchor.ParameterOnWallAxis;
            double resolvedFinishStation = anchorResult.FinishAnchor.ParameterOnWallAxis;
            double resolvedAnchorMin = Math.Min(resolvedStartStation, resolvedFinishStation);
            double resolvedAnchorMax = Math.Max(resolvedStartStation, resolvedFinishStation);
            if (resolvedAnchorMax - resolvedAnchorMin <= StationEps * 2.0)
            {
                return Unsupported(selectedId, "Resolved wall anchor span is too short for mid-run acceptance.");
            }

            QuickDimensionWallSpikeSide side = GetSide(wallStart, wallDirection, sidePickPoint);
            if (side == QuickDimensionWallSpikeSide.Unspecified)
            {
                return Unsupported(selectedId, "Side pick point lies on the wall axis; pick clearly to one side.");
            }

            XYZ sideNormal = GetSideNormal(wallDirection, side);
            ShellLayerType shell = GetSelectedShellLayer(selectedWall, sideNormal);

            double halfWidth = 0.0;
            try { halfWidth = selectedWall.Width * 0.5; } catch { halfWidth = 0.0; }
            XYZ sideLinePoint = new XYZ(
                wallStart.X + (sideNormal.X * halfWidth),
                wallStart.Y + (sideNormal.Y * halfWidth),
                wallStart.Z);

            // Separate end-join provenance sets.
            List<long> startJoinIds = CollectWallIdsFromJoin(locationCurve, 0);
            List<long> endJoinIds = CollectWallIdsFromJoin(locationCurve, 1);
            List<long> geometryJoinIds = CollectGeometryJoinIds(doc, selectedWall);

            var startJoinSet = new HashSet<long>(startJoinIds);
            var endJoinSet = new HashSet<long>(endJoinIds);
            var geometryJoinSet = new HashSet<long>(geometryJoinIds);

            GeometryElement selectedGeometry = TryGetWallGeometry(selectedWall);
            var selectedVerticalHits = new List<VerticalReferenceHit>();
            CollectVerticalReferenceHits(doc, selectedGeometry, wallStart, wallDirection, sideLinePoint, selectedVerticalHits);

            List<Wall> candidateWalls = new FilteredElementCollector(doc, view.Id)
                .OfClass(typeof(Wall))
                .OfCategory(BuiltInCategory.OST_Walls)
                .Cast<Wall>()
                .Where(w => w != null
                            && w.IsValidObject
                            && w.Id.Value != selectedId
                            && w.WallType?.Kind != WallKind.Curtain
                            && w.Location is LocationCurve lc && lc.Curve is Line)
                .ToList();

            var records = new List<QuickDimensionWallMidRunCandidate>();

            foreach (Wall candidate in candidateWalls)
            {
                QuickDimensionWallMidRunCandidate record = ProbeCandidate(
                    doc,
                    candidate,
                    wallStart,
                    wallDirection,
                    axisLength,
                    resolvedStartStation,
                    resolvedFinishStation,
                    resolvedAnchorMin,
                    resolvedAnchorMax,
                    sideLinePoint,
                    selectedVerticalHits,
                    startJoinSet,
                    endJoinSet,
                    geometryJoinSet);

                if (record != null)
                {
                    records.Add(record);
                }
            }

            int midRun = records.Count(r => r.Relation == QuickDimensionWallMidRunRelation.MidRunCrossing);
            int proximity = records.Count(r => r.Relation == QuickDimensionWallMidRunRelation.NonJoinedProximity);
            string message =
                $"Mid-run probe (log-only). Shell: {shell}; side: {side}. " +
                $"Candidates: {records.Count}; mid-run crossings: {midRun}; non-joined proximity: {proximity}. " +
                $"Join sets -> ElementsAtJoin[start]:{startJoinIds.Count}, ElementsAtJoin[end]:{endJoinIds.Count}, GeometryJoin:{geometryJoinIds.Count}.";

            return new QuickDimensionWallMidRunProbeResult(
                true,
                selectedId,
                side,
                shell,
                axisLength,
                startJoinIds,
                endJoinIds,
                geometryJoinIds,
                records,
                message);
        }

        private static QuickDimensionWallMidRunCandidate ProbeCandidate(
            Document doc,
            Wall candidate,
            XYZ wallStart,
            XYZ wallDirection,
            double axisLength,
            double resolvedStartStation,
            double resolvedFinishStation,
            double resolvedAnchorMin,
            double resolvedAnchorMax,
            XYZ sideLinePoint,
            IReadOnlyList<VerticalReferenceHit> selectedVerticalHits,
            HashSet<long> startJoinSet,
            HashSet<long> endJoinSet,
            HashSet<long> geometryJoinSet)
        {
            long cid = candidate.Id.Value;

            Line candidateLine = ((LocationCurve)candidate.Location).Curve as Line;
            if (candidateLine == null ||
                !TryGetPlanarDirection(candidateLine.GetEndPoint(0), candidateLine.GetEndPoint(1), out XYZ candidateDirection, out _))
            {
                return null;
            }

            double directionDot = Math.Abs(candidateDirection.DotProduct(wallDirection));
            bool isPerpendicular = directionDot < PerpendicularDotTolerance;
            bool isParallel = directionDot > ParallelDotTolerance;

            bool inStart = startJoinSet.Contains(cid);
            bool inEnd = endJoinSet.Contains(cid);
            bool inGeom = geometryJoinSet.Contains(cid);

            // Capture EVERY distinct vertical Edge.Reference fact on the selected side line inside the span.
            // A T-joint can expose two jambs; selecting one representative would hide a real station.
            GeometryElement candidateGeometry = TryGetWallGeometry(candidate);
            var candidateVerticalHits = new List<VerticalReferenceHit>();
            CollectVerticalReferenceHits(doc, candidateGeometry, wallStart, wallDirection, sideLinePoint, candidateVerticalHits);

            IReadOnlyList<QuickDimensionWallMidRunReferenceHit> referenceHits = candidateVerticalHits
                .Where(hit => hit.Distance <= SideLineTolerance
                              && hit.Station > StationEps
                              && hit.Station < axisLength - StationEps)
                .Select(hit => BuildReferenceHit(hit, selectedVerticalHits))
                .ToList()
                .AsReadOnly();

            int acceptedMidRunStationCount = CountAcceptedMidRunStations(
                referenceHits,
                resolvedStartStation,
                resolvedFinishStation,
                resolvedAnchorMin,
                resolvedAnchorMax,
                axisLength,
                inStart,
                inEnd);

            XYZ fallbackPoint = NearestEndpointToAxis(candidateLine, wallStart, wallDirection);
            double fallbackStation = ProjectStation(fallbackPoint, wallStart, wallDirection);
            double fallbackDistance = DistanceToSideLine(fallbackPoint, sideLinePoint, wallDirection);

            // Join membership and wall angle remain diagnostic provenance only. They must not gate a real
            // reference-evidence crossing: the smoke fixture proved an oblique mid-run T is invisible to both join APIs.
            QuickDimensionWallMidRunRelation relation = Classify(
                isParallel,
                referenceHits,
                acceptedMidRunStationCount,
                inStart,
                inEnd,
                inGeom);

            return new QuickDimensionWallMidRunCandidate(
                cid,
                candidate.WallType?.Name ?? string.Empty,
                relation,
                inStart,
                inEnd,
                inGeom,
                isPerpendicular,
                isParallel,
                fallbackStation,
                fallbackDistance,
                referenceHits.Count > 0 ? "vertical-edge-on-side-line" : "location-curve-endpoint",
                referenceHits,
                acceptedMidRunStationCount);
        }

        private static QuickDimensionWallMidRunRelation Classify(
            bool isParallel,
            IReadOnlyList<QuickDimensionWallMidRunReferenceHit> referenceHits,
            int acceptedMidRunStationCount,
            bool inStart,
            bool inEnd,
            bool inGeom)
        {
            if (inStart || inEnd)
            {
                return QuickDimensionWallMidRunRelation.EndJoinOnly;
            }

            // MidRunCrossing means accepted reference-evidence crossing only. It is NOT proof of Revit join topology.
            // The Section 11 fixture proved both join APIs and perpendicularity can be false for a real T-joint.
            if (acceptedMidRunStationCount >= 2)
            {
                return QuickDimensionWallMidRunRelation.MidRunCrossing;
            }

            if (inGeom)
            {
                return QuickDimensionWallMidRunRelation.GeometryJoinOnly;
            }

            if (referenceHits.Count > 0)
            {
                return QuickDimensionWallMidRunRelation.NonJoinedProximity;
            }

            return isParallel
                ? QuickDimensionWallMidRunRelation.ParallelNonJoined
                : QuickDimensionWallMidRunRelation.Ignored;
        }

        private static int CountAcceptedMidRunStations(
            IReadOnlyList<QuickDimensionWallMidRunReferenceHit> referenceHits,
            double resolvedStartStation,
            double resolvedFinishStation,
            double resolvedAnchorMin,
            double resolvedAnchorMax,
            double axisLength,
            bool inStart,
            bool inEnd)
        {
            if (referenceHits == null || referenceHits.Count == 0 || inStart || inEnd)
            {
                return 0;
            }

            List<double> acceptedStations = referenceHits
                .Where(hit => hit.CandidateReferenceNormalAlongAxis
                              && hit.StationOnSelectedAxis > resolvedAnchorMin + StationEps
                              && hit.StationOnSelectedAxis < resolvedAnchorMax - StationEps
                              && Math.Abs(hit.StationOnSelectedAxis - resolvedStartStation) > StationEps
                              && Math.Abs(hit.StationOnSelectedAxis - resolvedFinishStation) > StationEps
                              && hit.StationOnSelectedAxis > StationEps
                              && hit.StationOnSelectedAxis < axisLength - StationEps)
                .Select(hit => hit.StationOnSelectedAxis)
                .OrderBy(station => station)
                .ToList();

            int count = 0;
            double? previousStation = null;
            foreach (double station in acceptedStations)
            {
                if (!previousStation.HasValue || Math.Abs(station - previousStation.Value) > StationEps)
                {
                    count++;
                    previousStation = station;
                }
            }

            return count;
        }

        private static QuickDimensionWallMidRunReferenceHit BuildReferenceHit(
            VerticalReferenceHit candidateHit,
            IReadOnlyList<VerticalReferenceHit> selectedVerticalHits)
        {
            VerticalReferenceHit selectedHit = selectedVerticalHits
                .Where(hit => hit.Distance <= SideLineTolerance &&
                              Math.Abs(hit.Station - candidateHit.Station) <= StationEps)
                .OrderBy(hit => hit.Distance)
                .FirstOrDefault();

            return new QuickDimensionWallMidRunReferenceHit(
                candidateHit.Midpoint,
                candidateHit.Station,
                candidateHit.Distance,
                candidateHit.NormalAlongAxis,
                selectedHit != null,
                selectedHit?.NormalAlongAxis ?? false);
        }

        private static void CollectVerticalReferenceHits(
            Document doc,
            GeometryElement geometry,
            XYZ wallStart,
            XYZ wallDirection,
            XYZ sideLinePoint,
            List<VerticalReferenceHit> hits)
        {
            if (geometry == null)
            {
                return;
            }

            foreach (GeometryObject geometryObject in geometry)
            {
                if (geometryObject is Solid solid)
                {
                    ScanSolidForVerticalReferenceHits(doc, solid, wallStart, wallDirection, sideLinePoint, hits);
                }
                else if (geometryObject is GeometryInstance instance)
                {
                    GeometryElement instanceGeometry = instance.GetInstanceGeometry();
                    if (instanceGeometry != null)
                    {
                        CollectVerticalReferenceHits(doc, instanceGeometry, wallStart, wallDirection, sideLinePoint, hits);
                    }
                }
            }
        }

        private static void ScanSolidForVerticalReferenceHits(
            Document doc,
            Solid solid,
            XYZ wallStart,
            XYZ wallDirection,
            XYZ sideLinePoint,
            List<VerticalReferenceHit> hits)
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
                    if (string.IsNullOrWhiteSpace(stableKey) || hits.Any(hit => hit.StableKey == stableKey))
                    {
                        continue;
                    }

                    XYZ midpoint = (line.GetEndPoint(0) + line.GetEndPoint(1)) * 0.5;
                    double distance = DistanceToSideLine(midpoint, sideLinePoint, wallDirection);
                    double station = ProjectStation(midpoint, wallStart, wallDirection);
                    bool normalAlongAxis = EdgeFaceNormalAlongAxis(edge, wallDirection);

                    hits.Add(new VerticalReferenceHit(
                        stableKey,
                        new XYZ(midpoint.X, midpoint.Y, midpoint.Z),
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
            Face faceA = null;
            Face faceB = null;
            try { faceA = edge.GetFace(0); } catch { }
            try { faceB = edge.GetFace(1); } catch { }

            return FaceNormalParallelToAxis(faceA, wallDirection) || FaceNormalParallelToAxis(faceB, wallDirection);
        }

        private static bool FaceNormalParallelToAxis(Face face, XYZ wallDirection)
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
                ElementArray joined = locationCurve.get_ElementsAtJoin(end);
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
                // Wall end-joins often are not geometry-joins.
            }

            return ids;
        }

        private static GeometryElement TryGetWallGeometry(Wall wall)
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

        private static ShellLayerType GetSelectedShellLayer(Wall wall, XYZ targetSideNormal)
        {
            XYZ orientation = wall?.Orientation;
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

        private static QuickDimensionWallSpikeSide GetSide(XYZ wallStart, XYZ wallDirection, XYZ sidePickPoint)
        {
            XYZ offset = sidePickPoint - wallStart;
            double cross = (wallDirection.X * offset.Y) - (wallDirection.Y * offset.X);
            if (cross > StationEps)
            {
                return QuickDimensionWallSpikeSide.Left;
            }

            if (cross < -StationEps)
            {
                return QuickDimensionWallSpikeSide.Right;
            }

            return QuickDimensionWallSpikeSide.Unspecified;
        }

        private static XYZ GetSideNormal(XYZ wallDirection, QuickDimensionWallSpikeSide side)
        {
            XYZ leftNormal = new XYZ(-wallDirection.Y, wallDirection.X, 0.0);
            return side == QuickDimensionWallSpikeSide.Left ? leftNormal : leftNormal.Negate();
        }

        private static bool TryGetPlanarDirection(XYZ start, XYZ end, out XYZ direction, out double length)
        {
            direction = null;
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

        private sealed class VerticalReferenceHit
        {
            public VerticalReferenceHit(
                string stableKey,
                XYZ midpoint,
                double station,
                double distance,
                bool normalAlongAxis)
            {
                StableKey = stableKey ?? string.Empty;
                Midpoint = midpoint;
                Station = station;
                Distance = distance;
                NormalAlongAxis = normalAlongAxis;
            }

            public string StableKey { get; }
            public XYZ Midpoint { get; }
            public double Station { get; }
            public double Distance { get; }
            public bool NormalAlongAxis { get; }
        }

        private static QuickDimensionWallMidRunProbeResult Unsupported(long selectedWallId, string message)
        {
            return new QuickDimensionWallMidRunProbeResult(
                false,
                selectedWallId,
                QuickDimensionWallSpikeSide.Unspecified,
                null,
                0.0,
                new List<long>(),
                new List<long>(),
                new List<long>(),
                new List<QuickDimensionWallMidRunCandidate>(),
                message);
        }
    }
}
