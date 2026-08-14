using System;
using System.Collections.Generic;
using System.Linq;
using ArcTool.Core.Archive.QuickDimension.Models;
using Autodesk.Revit.DB;

namespace ArcTool.Core.Archive.QuickDimension.Services
{
    /// <summary>
    /// Wall-axis spike service for validating only the selected wall's picked-side end anchors.
    /// Performs no transactions and never creates dimensions.
    ///
    /// Anchor model (derived from Genius Loci "Wall Edges References" raw edge data):
    /// - T-joints are handled by selected-wall side-face boundary candidates.
    /// - L-joints need both selected-wall side-face horizontal loop vertices and joined-wall
    ///   outward candidates because some visible outer corners belong to the wall being joined.
    /// </summary>
    public static class QuickDimensionWallReferenceProbeService
    {
        private const double MinimumWallAxisLength = 1e-6;
        private const double VerticalEdgeTolerance = 1e-3;
        private const double DuplicateStationTolerance = 1e-4;
        private static readonly double SideFacePlaneTolerance = UnitUtils.ConvertToInternalUnits(1.0, UnitTypeId.Millimeters);
        private static readonly double SideLineTolerance = UnitUtils.ConvertToInternalUnits(5.0, UnitTypeId.Millimeters);
        private static readonly double JoinExtensionMargin = UnitUtils.ConvertToInternalUnits(500.0, UnitTypeId.Millimeters);
        // When a joined-wall candidate lands on a station the selected wall already owns, that point is
        // Revit join-cleanup geometry on the selected wall itself, not a genuinely missing outer corner.
        // Extending to it overshoots the true wall end (the interior-shell failure proven by smoke logs).
        private static readonly double JoinCoincidenceTolerance = UnitUtils.ConvertToInternalUnits(1.0, UnitTypeId.Millimeters);
        private static readonly double FullHeightReferenceTolerance = UnitUtils.ConvertToInternalUnits(10.0, UnitTypeId.Millimeters);

        public static QuickDimensionWallSpikeResult RunWallReferenceProbe(Wall selectedWall, XYZ sidePickPoint)
        {
            if (selectedWall == null) throw new ArgumentNullException(nameof(selectedWall));
            if (sidePickPoint == null) throw new ArgumentNullException(nameof(sidePickPoint));

            string wallTypeName = selectedWall.WallType?.Name ?? string.Empty;

            if (!selectedWall.IsValidObject)
            {
                return Fail(selectedWall.Id, wallTypeName, 0.0, QuickDimensionWallSpikeSide.Unspecified, "The selected wall is not a valid Revit object.");
            }

            if (selectedWall.WallType?.Kind == WallKind.Curtain)
            {
                return Fail(selectedWall.Id, wallTypeName, 0.0, QuickDimensionWallSpikeSide.Unspecified, "Curtain walls are outside this wall spike scope.");
            }

            if (selectedWall.Location is not LocationCurve locationCurve || locationCurve.Curve is not Line wallLine)
            {
                return Fail(selectedWall.Id, wallTypeName, 0.0, QuickDimensionWallSpikeSide.Unspecified, "The selected wall does not expose a straight LocationCurve line.");
            }

            XYZ wallStart = wallLine.GetEndPoint(0);
            XYZ wallEnd = wallLine.GetEndPoint(1);
            if (!TryGetPlanarDirection(wallStart, wallEnd, out XYZ wallDirection, out double wallAxisLength))
            {
                return Fail(selectedWall.Id, wallTypeName, 0.0, QuickDimensionWallSpikeSide.Unspecified, "The selected wall axis is too short or invalid.");
            }

            QuickDimensionWallSpikeSide side = GetSide(wallStart, wallDirection, sidePickPoint);
            if (side == QuickDimensionWallSpikeSide.Unspecified)
            {
                return Fail(selectedWall.Id, wallTypeName, wallAxisLength, side, "The side pick point lies on the wall axis. Pick clearly on one side of the wall.");
            }

            XYZ targetSideNormal = GetSideNormal(wallDirection, side);
            ShellLayerType shellLayer = GetSelectedShellLayer(selectedWall, targetSideNormal);

            if (!TryCollectSideFaceBoundary(
                    selectedWall,
                    wallStart,
                    wallDirection,
                    shellLayer,
                    includeBothShells: false,
                    out List<WallSpikeBoundaryCandidate> selectedCandidates,
                    out List<WallSpikeHorizontalSegmentCandidate> selectedHorizontalSegments,
                    out string failureReason))
            {
                return Fail(selectedWall.Id, wallTypeName, wallAxisLength, side, shellLayer, 0, failureReason);
            }

            WallSpikeHorizontalSegmentCandidate mainRun = selectedHorizontalSegments
                .OrderByDescending(segment => segment.Length)
                .FirstOrDefault();

            if (mainRun == null)
            {
                return Fail(selectedWall.Id, wallTypeName, wallAxisLength, side, shellLayer, selectedCandidates.Count, "Picked side face did not expose a horizontal boundary edge parallel to the wall axis.");
            }

            WallSpikeBoundaryCandidate baseStart = SelectBestCandidateAtPoint(selectedCandidates, mainRun.StartPoint, mainRun.StartStation);
            WallSpikeBoundaryCandidate baseFinish = SelectBestCandidateAtPoint(selectedCandidates, mainRun.EndPoint, mainRun.EndStation);
            if (baseFinish.ParameterOnWallAxis < baseStart.ParameterOnWallAxis)
            {
                (baseStart, baseFinish) = (baseFinish, baseStart);
            }

            List<Wall> joinedWalls = CollectJoinedWalls(selectedWall);
            List<WallSpikeBoundaryCandidate> joinedCandidates = CollectJoinedWallBoundaryCandidates(
                joinedWalls,
                wallStart,
                wallDirection,
                mainRun.StartPoint,
                targetSideNormal);

            WallSpikeBoundaryCandidate startAnchorCandidate = ResolveEndCandidate(
                baseStart,
                selectedCandidates,
                joinedCandidates,
                shellLayer,
                isStartEnd: true);

            WallSpikeBoundaryCandidate finishAnchorCandidate = ResolveEndCandidate(
                baseFinish,
                selectedCandidates,
                joinedCandidates,
                shellLayer,
                isStartEnd: false);

            if (Math.Abs(finishAnchorCandidate.ParameterOnWallAxis - startAnchorCandidate.ParameterOnWallAxis) <= DuplicateStationTolerance)
            {
                return Fail(selectedWall.Id, wallTypeName, wallAxisLength, side, shellLayer, selectedCandidates.Count, "Resolved anchors did not produce two distinct wall-end stations.");
            }

            var startAnchor = new QuickDimensionWallSpikeAnchor(
                "Start Anchor",
                startAnchorCandidate.Reference,
                startAnchorCandidate.Point,
                startAnchorCandidate.ParameterOnWallAxis);

            var finishAnchor = new QuickDimensionWallSpikeAnchor(
                "Finish Anchor",
                finishAnchorCandidate.Reference,
                finishAnchorCandidate.Point,
                finishAnchorCandidate.ParameterOnWallAxis);

            int mappedReferenceCount = (startAnchorCandidate.Reference != null ? 1 : 0) + (finishAnchorCandidate.Reference != null ? 1 : 0);
            string message =
                $"Side-face boundary model (vertical + horizontal + joined-wall outward candidates). Shell layer: {shellLayer}. " +
                $"Selected candidates: {selectedCandidates.Count}; joined walls: {joinedWalls.Count}; joined candidates on side line: {joinedCandidates.Count}. " +
                $"Start source: {startAnchorCandidate.Source}; Finish source: {finishAnchorCandidate.Source}. " +
                $"Edge.Reference mapped: {mappedReferenceCount}/2.";

            return new QuickDimensionWallSpikeResult(
                true,
                selectedWall.Id,
                wallTypeName,
                wallAxisLength,
                side,
                shellLayer,
                selectedCandidates.Count,
                startAnchor,
                finishAnchor,
                message);
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

        public static IReadOnlyList<QuickDimensionWallSpikeCornerProbePoint> CollectBoundaryCornerPointsForLog(
            Wall wall,
            XYZ selectedWallStart,
            XYZ selectedWallDirection,
            ShellLayerType shellLayer,
            bool includeBothShells,
            out string failureReason)
        {
            if (wall == null) throw new ArgumentNullException(nameof(wall));
            if (selectedWallStart == null) throw new ArgumentNullException(nameof(selectedWallStart));
            if (selectedWallDirection == null) throw new ArgumentNullException(nameof(selectedWallDirection));

            if (!TryCollectSideFaceBoundary(
                    wall,
                    selectedWallStart,
                    selectedWallDirection,
                    shellLayer,
                    includeBothShells,
                    out List<WallSpikeBoundaryCandidate> candidates,
                    out _,
                    out failureReason))
            {
                return Array.Empty<QuickDimensionWallSpikeCornerProbePoint>();
            }

            return candidates
                .Select(candidate => new QuickDimensionWallSpikeCornerProbePoint(
                    candidate.Point,
                    candidate.ParameterOnWallAxis,
                    candidate.SourceWallId,
                    candidate.Source))
                .ToList();
        }

        private static bool TryCollectSideFaceBoundary(
            Wall wall,
            XYZ wallStart,
            XYZ wallDirection,
            ShellLayerType shellLayer,
            bool includeBothShells,
            out List<WallSpikeBoundaryCandidate> candidates,
            out List<WallSpikeHorizontalSegmentCandidate> horizontalSegments,
            out string failureReason)
        {
            candidates = new List<WallSpikeBoundaryCandidate>();
            horizontalSegments = new List<WallSpikeHorizontalSegmentCandidate>();
            failureReason = string.Empty;

            var sidePlanes = new List<PlanarFace>();
            var shells = includeBothShells
                ? new[] { ShellLayerType.Exterior, ShellLayerType.Interior }
                : new[] { shellLayer };

            foreach (ShellLayerType layer in shells)
            {
                IList<Reference> sideFaceRefs;
                try
                {
                    sideFaceRefs = HostObjectUtils.GetSideFaces(wall, layer);
                }
                catch (Exception ex)
                {
                    if (!includeBothShells)
                    {
                        failureReason = $"HostObjectUtils.GetSideFaces({layer}) failed: {ex.Message}";
                        return false;
                    }

                    continue;
                }

                if (sideFaceRefs == null)
                {
                    continue;
                }

                foreach (Reference sideFaceRef in sideFaceRefs)
                {
                    if (wall.GetGeometryObjectFromReference(sideFaceRef) is PlanarFace planarFace)
                    {
                        sidePlanes.Add(planarFace);
                    }
                }
            }

            if (sidePlanes.Count == 0)
            {
                failureReason = includeBothShells
                    ? "Joined wall side faces could not be resolved as planar faces."
                    : $"Side faces for {shellLayer} shell layer were not planar; unsupported wall geometry.";
                return false;
            }

            if (!TryGetWallGeometry(wall, out GeometryElement geometryElement, out failureReason))
            {
                return false;
            }

            CollectSideFaceBoundaryCandidates(geometryElement, wallStart, wallDirection, sidePlanes, candidates, horizontalSegments, wall.Id.Value);
            return candidates.Count > 0;
        }

        private static void CollectSideFaceBoundaryCandidates(
            GeometryElement geometryElement,
            XYZ wallStart,
            XYZ wallDirection,
            IReadOnlyList<PlanarFace> sidePlanes,
            List<WallSpikeBoundaryCandidate> candidates,
            List<WallSpikeHorizontalSegmentCandidate> horizontalSegments,
            long sourceWallId)
        {
            foreach (GeometryObject geometryObject in geometryElement)
            {
                if (geometryObject is Solid solid)
                {
                    CollectSideFaceBoundaryCandidates(solid, wallStart, wallDirection, sidePlanes, candidates, horizontalSegments, sourceWallId);
                }
                else if (geometryObject is GeometryInstance geometryInstance)
                {
                    GeometryElement instanceGeometry = geometryInstance.GetInstanceGeometry();
                    if (instanceGeometry != null)
                    {
                        CollectSideFaceBoundaryCandidates(instanceGeometry, wallStart, wallDirection, sidePlanes, candidates, horizontalSegments, sourceWallId);
                    }
                }
            }
        }

        private static void CollectSideFaceBoundaryCandidates(
            Solid solid,
            XYZ wallStart,
            XYZ wallDirection,
            IReadOnlyList<PlanarFace> sidePlanes,
            List<WallSpikeBoundaryCandidate> candidates,
            List<WallSpikeHorizontalSegmentCandidate> horizontalSegments,
            long sourceWallId)
        {
            if (solid == null || solid.Edges == null || solid.Edges.Size == 0)
            {
                return;
            }

            foreach (Edge edge in solid.Edges)
            {
                try
                {
                    if (edge.AsCurve() is not Line line)
                    {
                        continue;
                    }

                    XYZ direction = line.Direction;
                    if (direction == null)
                    {
                        continue;
                    }

                    if (!EdgeTouchesSideFace(edge, sidePlanes))
                    {
                        continue;
                    }

                    Reference reference = edge.Reference;
                    if (Math.Abs(Math.Abs(direction.Z) - 1.0) <= VerticalEdgeTolerance)
                    {
                        XYZ midpoint = (line.GetEndPoint(0) + line.GetEndPoint(1)) * 0.5;
                        AddBoundaryCandidate(candidates, reference, midpoint, wallStart, wallDirection, sourceWallId, "vertical-edge");
                        continue;
                    }

                    XYZ planarDirection = new XYZ(direction.X, direction.Y, 0.0);
                    if (planarDirection.GetLength() <= MinimumWallAxisLength)
                    {
                        continue;
                    }

                    if (Math.Abs(Math.Abs(planarDirection.Normalize().DotProduct(wallDirection)) - 1.0) > VerticalEdgeTolerance)
                    {
                        continue;
                    }

                    XYZ startPoint = line.GetEndPoint(0);
                    XYZ endPoint = line.GetEndPoint(1);
                    WallSpikeBoundaryCandidate start = AddBoundaryCandidate(candidates, null, startPoint, wallStart, wallDirection, sourceWallId, "horizontal-endpoint");
                    WallSpikeBoundaryCandidate end = AddBoundaryCandidate(candidates, null, endPoint, wallStart, wallDirection, sourceWallId, "horizontal-endpoint");
                    horizontalSegments.Add(new WallSpikeHorizontalSegmentCandidate(start.Point, end.Point, start.ParameterOnWallAxis, end.ParameterOnWallAxis));
                }
                catch
                {
                    // Skip malformed edge.
                }
            }
        }

        private static WallSpikeBoundaryCandidate AddBoundaryCandidate(
            List<WallSpikeBoundaryCandidate> candidates,
            Reference reference,
            XYZ point,
            XYZ wallStart,
            XYZ wallDirection,
            long sourceWallId,
            string source)
        {
            XYZ planPoint = new XYZ(point.X, point.Y, point.Z);
            double station = (planPoint - wallStart).DotProduct(wallDirection);
            var candidate = new WallSpikeBoundaryCandidate(reference, planPoint, station, sourceWallId, source);
            candidates.Add(candidate);
            return candidate;
        }

        private static WallSpikeBoundaryCandidate SelectBestCandidateAtPoint(
            IReadOnlyList<WallSpikeBoundaryCandidate> candidates,
            XYZ point,
            double station)
        {
            WallSpikeBoundaryCandidate matched = candidates
                .Where(candidate => Math.Abs(candidate.ParameterOnWallAxis - station) <= DuplicateStationTolerance)
                .OrderByDescending(candidate => candidate.Reference != null)
                .FirstOrDefault();

            return matched ?? new WallSpikeBoundaryCandidate(null, point, station, 0, "main-run-endpoint");
        }

        private static WallSpikeBoundaryCandidate ResolveEndCandidate(
            WallSpikeBoundaryCandidate baseCandidate,
            IReadOnlyList<WallSpikeBoundaryCandidate> selectedCandidates,
            IReadOnlyList<WallSpikeBoundaryCandidate> joinedCandidates,
            ShellLayerType shellLayer,
            bool isStartEnd)
        {
            if (shellLayer == ShellLayerType.Interior)
            {
                WallSpikeBoundaryCandidate inward = SelectDirectionalFullHeightReference(
                    baseCandidate,
                    selectedCandidates,
                    joinedCandidates,
                    isStartEnd,
                    searchInward: true);
                return inward ?? baseCandidate;
            }

            WallSpikeBoundaryCandidate outward = SelectDirectionalFullHeightReference(
                baseCandidate,
                Array.Empty<WallSpikeBoundaryCandidate>(),
                joinedCandidates,
                isStartEnd,
                searchInward: false);

            return outward ?? baseCandidate;
        }

        private static WallSpikeBoundaryCandidate SelectDirectionalFullHeightReference(
            WallSpikeBoundaryCandidate baseCandidate,
            IReadOnlyList<WallSpikeBoundaryCandidate> selectedCandidates,
            IReadOnlyList<WallSpikeBoundaryCandidate> joinedCandidates,
            bool isStartEnd,
            bool searchInward)
        {
            List<WallSpikeBoundaryCandidate> candidates = new List<WallSpikeBoundaryCandidate>();
            if (selectedCandidates != null)
            {
                candidates.AddRange(selectedCandidates);
            }

            if (joinedCandidates != null)
            {
                candidates.AddRange(joinedCandidates);
            }

            if (candidates.Count == 0)
            {
                return null;
            }

            List<WallSpikeBoundaryCandidate> referenceCandidates = candidates
                .Where(candidate => candidate.Reference != null && candidate.Point != null)
                .ToList();
            if (referenceCandidates.Count == 0)
            {
                return null;
            }

            double maxMidpointZ = referenceCandidates.Max(candidate => candidate.Point.Z);
            double fullHeightThreshold = maxMidpointZ - FullHeightReferenceTolerance;

            IEnumerable<WallSpikeBoundaryCandidate> directional = referenceCandidates.Where(candidate =>
                candidate.Point.Z >= fullHeightThreshold &&
                IsCandidateInRequestedDirection(baseCandidate, candidate, isStartEnd, searchInward) &&
                Math.Abs(candidate.ParameterOnWallAxis - baseCandidate.ParameterOnWallAxis) <= JoinExtensionMargin);

            return searchInward
                ? directional.OrderBy(candidate => Math.Abs(candidate.ParameterOnWallAxis - baseCandidate.ParameterOnWallAxis)).FirstOrDefault()
                : (isStartEnd
                    ? directional.OrderBy(candidate => candidate.ParameterOnWallAxis).FirstOrDefault()
                    : directional.OrderByDescending(candidate => candidate.ParameterOnWallAxis).FirstOrDefault());
        }

        private static bool IsCandidateInRequestedDirection(
            WallSpikeBoundaryCandidate baseCandidate,
            WallSpikeBoundaryCandidate candidate,
            bool isStartEnd,
            bool searchInward)
        {
            if (isStartEnd)
            {
                return searchInward
                    ? candidate.ParameterOnWallAxis > baseCandidate.ParameterOnWallAxis + DuplicateStationTolerance
                    : candidate.ParameterOnWallAxis < baseCandidate.ParameterOnWallAxis - DuplicateStationTolerance;
            }

            return searchInward
                ? candidate.ParameterOnWallAxis < baseCandidate.ParameterOnWallAxis - DuplicateStationTolerance
                : candidate.ParameterOnWallAxis > baseCandidate.ParameterOnWallAxis + DuplicateStationTolerance;
        }

        /// <summary>
        /// Collects the walls joined to <paramref name="selectedWall"/> using the same
        /// ElementsAtJoin (both ends) + JoinGeometryUtils path the anchor resolver uses.
        /// Exposed so the XML smoke-log service reports the identical joined set the probe relies on.
        /// </summary>
        public static List<Wall> CollectJoinedWalls(Wall selectedWall)
        {
            var result = new List<Wall>();
            var seenIds = new HashSet<long> { selectedWall.Id.Value };

            void AddWall(Element element)
            {
                if (element is Wall wall && wall.IsValidObject && wall.WallType?.Kind != WallKind.Curtain && seenIds.Add(wall.Id.Value))
                {
                    result.Add(wall);
                }
            }

            if (selectedWall.Location is LocationCurve locationCurve)
            {
                for (int end = 0; end <= 1; end++)
                {
                    try
                    {
                        ElementArray joined = locationCurve.get_ElementsAtJoin(end);
                        if (joined != null)
                        {
                            foreach (Element element in joined)
                            {
                                AddWall(element);
                            }
                        }
                    }
                    catch
                    {
                        // Ignore unsupported join state.
                    }
                }
            }

            Document doc = selectedWall.Document;
            if (doc != null)
            {
                try
                {
                    foreach (ElementId id in JoinGeometryUtils.GetJoinedElements(doc, selectedWall))
                    {
                        AddWall(doc.GetElement(id));
                    }
                }
                catch
                {
                    // Wall end-joins often are not geometry-joins; ElementsAtJoin is the primary path.
                }
            }

            return result;
        }

        private static List<WallSpikeBoundaryCandidate> CollectJoinedWallBoundaryCandidates(
            IReadOnlyList<Wall> joinedWalls,
            XYZ wallStart,
            XYZ wallDirection,
            XYZ sideLinePoint,
            XYZ targetSideNormal)
        {
            var result = new List<WallSpikeBoundaryCandidate>();
            if (joinedWalls == null || joinedWalls.Count == 0)
            {
                return result;
            }

            foreach (Wall joinedWall in joinedWalls)
            {
                if (!TryCollectSideFaceBoundary(
                        joinedWall,
                        wallStart,
                        wallDirection,
                        ShellLayerType.Exterior,
                        includeBothShells: true,
                        out List<WallSpikeBoundaryCandidate> candidates,
                        out _,
                        out _))
                {
                    continue;
                }

                foreach (WallSpikeBoundaryCandidate candidate in candidates)
                {
                    if (DistanceToSideLine(candidate.Point, sideLinePoint, wallDirection) <= SideLineTolerance)
                    {
                        result.Add(candidate);
                    }
                }
            }

            return result
                .GroupBy(candidate => new
                {
                    candidate.SourceWallId,
                    StationBucket = Math.Round(candidate.ParameterOnWallAxis / DuplicateStationTolerance)
                })
                .Select(group => group.OrderByDescending(candidate => candidate.Reference != null).First())
                .OrderBy(candidate => candidate.ParameterOnWallAxis)
                .ToList();
        }

        private static double DistanceToSideLine(XYZ point, XYZ linePoint, XYZ lineDirection)
        {
            XYZ offset = new XYZ(point.X - linePoint.X, point.Y - linePoint.Y, 0.0);
            return Math.Abs((offset.X * lineDirection.Y) - (offset.Y * lineDirection.X));
        }

        private static bool EdgeTouchesSideFace(Edge edge, IReadOnlyList<PlanarFace> sidePlanes)
        {
            Face faceA = null;
            Face faceB = null;
            try { faceA = edge.GetFace(0); } catch { }
            try { faceB = edge.GetFace(1); } catch { }

            foreach (PlanarFace sideFace in sidePlanes)
            {
                if (ReferenceEquals(faceA, sideFace) || ReferenceEquals(faceB, sideFace))
                {
                    return true;
                }
            }

            XYZ midpoint = edge.AsCurve() is Line line
                ? (line.GetEndPoint(0) + line.GetEndPoint(1)) * 0.5
                : null;
            if (midpoint == null)
            {
                return false;
            }

            foreach (PlanarFace sideFace in sidePlanes)
            {
                XYZ origin = sideFace.Origin;
                XYZ normal = sideFace.FaceNormal;
                if (origin == null || normal == null)
                {
                    continue;
                }

                double signedDistance = (midpoint - origin).DotProduct(normal);
                if (Math.Abs(signedDistance) > SideFacePlaneTolerance)
                {
                    continue;
                }

                if (FaceNormalMatches(faceA, normal) || FaceNormalMatches(faceB, normal))
                {
                    return true;
                }
            }

            return false;
        }

        private static bool FaceNormalMatches(Face face, XYZ sideNormal)
        {
            if (face is not PlanarFace planarFace || planarFace.FaceNormal == null || sideNormal == null)
            {
                return false;
            }

            return planarFace.FaceNormal.Normalize().DotProduct(sideNormal.Normalize()) > 0.9;
        }

        private static bool TryGetWallGeometry(Wall wall, out GeometryElement geometryElement, out string failureReason)
        {
            geometryElement = null;
            failureReason = string.Empty;

            Options options = new Options
            {
                ComputeReferences = true,
                IncludeNonVisibleObjects = true
            };

            geometryElement = wall.get_Geometry(options);
            if (geometryElement == null)
            {
                failureReason = "wall.get_Geometry(ComputeReferences=true) returned no geometry.";
                return false;
            }

            return true;
        }

        private static QuickDimensionWallSpikeSide GetSide(XYZ wallStart, XYZ wallDirection, XYZ sidePickPoint)
        {
            XYZ offset = sidePickPoint - wallStart;
            double cross = (wallDirection.X * offset.Y) - (wallDirection.Y * offset.X);
            if (cross > DuplicateStationTolerance)
            {
                return QuickDimensionWallSpikeSide.Left;
            }

            if (cross < -DuplicateStationTolerance)
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

        private static QuickDimensionWallSpikeResult Fail(
            ElementId wallId,
            string wallTypeName,
            double wallAxisLength,
            QuickDimensionWallSpikeSide side,
            string message)
        {
            return Fail(wallId, wallTypeName, wallAxisLength, side, null, 0, message);
        }

        private static QuickDimensionWallSpikeResult Fail(
            ElementId wallId,
            string wallTypeName,
            double wallAxisLength,
            QuickDimensionWallSpikeSide side,
            ShellLayerType? selectedShellLayer,
            int totalVerticalEdgesOnSide,
            string message)
        {
            return new QuickDimensionWallSpikeResult(
                false,
                wallId,
                wallTypeName,
                wallAxisLength,
                side,
                selectedShellLayer,
                totalVerticalEdgesOnSide,
                null,
                null,
                message);
        }

        private sealed class WallSpikeBoundaryCandidate
        {
            public WallSpikeBoundaryCandidate(Reference reference, XYZ point, double parameterOnWallAxis, long sourceWallId, string source)
            {
                Reference = reference;
                Point = point;
                ParameterOnWallAxis = parameterOnWallAxis;
                SourceWallId = sourceWallId;
                Source = source ?? string.Empty;
            }

            public Reference Reference { get; }
            public XYZ Point { get; }
            public double ParameterOnWallAxis { get; }
            public long SourceWallId { get; }
            public string Source { get; }
        }

        private sealed class WallSpikeHorizontalSegmentCandidate
        {
            public WallSpikeHorizontalSegmentCandidate(XYZ firstPoint, XYZ secondPoint, double firstStation, double secondStation)
            {
                if (firstStation <= secondStation)
                {
                    StartPoint = firstPoint;
                    EndPoint = secondPoint;
                    StartStation = firstStation;
                    EndStation = secondStation;
                }
                else
                {
                    StartPoint = secondPoint;
                    EndPoint = firstPoint;
                    StartStation = secondStation;
                    EndStation = firstStation;
                }
            }

            public XYZ StartPoint { get; }
            public XYZ EndPoint { get; }
            public double StartStation { get; }
            public double EndStation { get; }
            public double Length => Math.Abs(EndStation - StartStation);
        }
    }
}
