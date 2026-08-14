using System;
using System.Collections.Generic;
using System.Linq;
using ArcTool.Core.Archive.QuickDimension.Models;
using Autodesk.Revit.DB;
using RevitView = Autodesk.Revit.DB.View;

namespace ArcTool.Core.Archive.QuickDimension.Services
{
    /// <summary>
    /// Session 1.5 spike service: tests full mixed Grid + Wall + Door + Window reference arrays for NewDimension.
    /// Merges proven strategies from:
    /// - Session 1.1: Grid via new Reference(grid)
    /// - Session 1.2: Wall via HostObjectUtils.GetSideFaces with closest-face selection
    /// - Session 1.4: Door/Window via FamilyInstance.GetReferences(Left/Right) with HostWallOpeningGeometry fallback
    /// </summary>
    public static class QuickDimensionFullMixedReferenceProbeService
    {
        private const double MinimumDimensionLineLength = 1e-6;
        private const double ParallelDotTolerance = 0.98;
        private const double ProjectionTolerance = 1e-4;

        /// <summary>
        /// Runs the full mixed reference probe, testing Grid + Wall + Door + Window references in the same ReferenceArray.
        /// </summary>
        public static QuickDimensionFullMixedProbeSummary RunFullMixedReferenceProbe(
            Document doc,
            RevitView view,
            XYZ firstPoint,
            XYZ secondPoint)
        {
            if (doc == null) throw new ArgumentNullException(nameof(doc));
            if (view == null) throw new ArgumentNullException(nameof(view));
            if (firstPoint == null) throw new ArgumentNullException(nameof(firstPoint));
            if (secondPoint == null) throw new ArgumentNullException(nameof(secondPoint));

            Line dimensionLine = CreateDimensionLine(firstPoint, secondPoint);
            XYZ dimensionDirection = dimensionLine.Direction;
            double dimensionLength = dimensionLine.Length;

            // Collect all source types
            var gridResult = CollectGridCandidates(doc, view, firstPoint, dimensionDirection, dimensionLength);
            var wallResult = CollectWallCandidates(doc, view, firstPoint, dimensionDirection, dimensionLength);
            var openingResult = CollectDoorWindowCandidates(doc, view, firstPoint, dimensionDirection, dimensionLength);

            // Merge all candidates into unified list
            List<QuickDimensionFullMixedCandidate> allCandidates = new List<QuickDimensionFullMixedCandidate>();
            allCandidates.AddRange(gridResult.Candidates);
            allCandidates.AddRange(wallResult.Candidates);
            allCandidates.AddRange(openingResult.Candidates);

            // Sort by position along dimension line
            List<QuickDimensionFullMixedCandidate> sortedCandidates = allCandidates
                .OrderBy(c => c.ParameterOnDimensionLine)
                .ToList();

            // Probe all test scenarios
            var fullMixedResult = ProbeTest(doc, view, dimensionLine, sortedCandidates, "FullMixed");

            var gridsOnlyResult = ProbeTest(doc, view, dimensionLine,
                sortedCandidates.Where(c => c.SourceType == QuickDimensionFullMixedSourceType.Grid).ToList(),
                "GridsOnly");

            var wallsOnlyResult = ProbeTest(doc, view, dimensionLine,
                sortedCandidates.Where(c => c.SourceType == QuickDimensionFullMixedSourceType.Wall).ToList(),
                "WallsOnly");

            var openingsOnlyResult = ProbeTest(doc, view, dimensionLine,
                sortedCandidates.Where(c => c.IsOpening).ToList(),
                "OpeningsOnly");

            var gridWallResult = ProbeTest(doc, view, dimensionLine,
                sortedCandidates.Where(c => c.SourceType == QuickDimensionFullMixedSourceType.Grid ||
                                            c.SourceType == QuickDimensionFullMixedSourceType.Wall).ToList(),
                "GridWall");

            var wallOpeningResult = ProbeTest(doc, view, dimensionLine,
                sortedCandidates.Where(c => c.SourceType == QuickDimensionFullMixedSourceType.Wall || c.IsOpening).ToList(),
                "WallOpening");

            return new QuickDimensionFullMixedProbeSummary(
                gridResult.CollectedCount,
                wallResult.CollectedCount,
                openingResult.CollectedDoorCount,
                openingResult.CollectedWindowCount,
                gridResult.AcceptedCount,
                wallResult.AcceptedCount,
                openingResult.AcceptedDoorCount,
                openingResult.AcceptedWindowCount,
                gridResult.SkippedArcCount,
                gridResult.SkippedParallelCount,
                wallResult.SkippedCurtainCount,
                wallResult.SkippedParallelCount,
                wallResult.SkippedNoFaceReferenceCount,
                openingResult.SkippedNonHostedCount,
                openingResult.SkippedParallelCount,
                openingResult.SkippedOutsideSpanCount,
                openingResult.SkippedNoReferenceCount,
                fullMixedResult,
                gridsOnlyResult,
                wallsOnlyResult,
                openingsOnlyResult,
                gridWallResult,
                wallOpeningResult);
        }

        private static Line CreateDimensionLine(XYZ firstPoint, XYZ secondPoint)
        {
            if (firstPoint.DistanceTo(secondPoint) < MinimumDimensionLineLength)
            {
                throw new InvalidOperationException("The two picked points are too close to define a dimension line.");
            }

            return Line.CreateBound(firstPoint, secondPoint);
        }

        #region Grid Collection (Session 1.1 strategy)

        private sealed class GridCollectionResult
        {
            public int CollectedCount { get; set; }
            public int AcceptedCount { get; set; }
            public int SkippedArcCount { get; set; }
            public int SkippedParallelCount { get; set; }
            public List<QuickDimensionFullMixedCandidate> Candidates { get; } = new List<QuickDimensionFullMixedCandidate>();
        }

        private static GridCollectionResult CollectGridCandidates(
            Document doc,
            RevitView view,
            XYZ firstPoint,
            XYZ dimensionDirection,
            double dimensionLength)
        {
            var result = new GridCollectionResult();

            List<Grid> grids = new FilteredElementCollector(doc, view.Id)
                .OfClass(typeof(Grid))
                .Cast<Grid>()
                .ToList();

            result.CollectedCount = grids.Count;

            foreach (Grid grid in grids)
            {
                if (grid?.IsValidObject != true)
                {
                    continue;
                }

                Curve gridCurve = grid.Curve;
                if (gridCurve is not Line gridLine)
                {
                    result.SkippedArcCount++;
                    continue;
                }

                XYZ gridDirection = gridLine.Direction;
                if (Math.Abs(gridDirection.DotProduct(dimensionDirection)) > ParallelDotTolerance)
                {
                    result.SkippedParallelCount++;
                    continue;
                }

                // Project grid midpoint onto dimension line
                XYZ midpoint = (gridLine.GetEndPoint(0) + gridLine.GetEndPoint(1)) * 0.5;
                double parameter = (midpoint - firstPoint).DotProduct(dimensionDirection);
                if (parameter < -ProjectionTolerance || parameter > dimensionLength + ProjectionTolerance)
                {
                    continue;
                }

                // Session 1.1 proven strategy: new Reference(grid)
                Reference gridReference = new Reference(grid);

                result.Candidates.Add(new QuickDimensionFullMixedCandidate(
                    grid.Id,
                    QuickDimensionFullMixedSourceType.Grid,
                    $"Grid: {grid.Name}",
                    parameter,
                    gridReference));

                result.AcceptedCount++;
            }

            // Deduplicate by ElementId
            var seen = new HashSet<long>();
            result.Candidates.RemoveAll(c => !seen.Add(c.ElementId.Value));

            return result;
        }

        #endregion

        #region Wall Collection (Session 1.2 strategy)

        private sealed class WallCollectionResult
        {
            public int CollectedCount { get; set; }
            public int AcceptedCount { get; set; }
            public int SkippedCurtainCount { get; set; }
            public int SkippedParallelCount { get; set; }
            public int SkippedNoFaceReferenceCount { get; set; }
            public List<QuickDimensionFullMixedCandidate> Candidates { get; } = new List<QuickDimensionFullMixedCandidate>();
        }

        private static WallCollectionResult CollectWallCandidates(
            Document doc,
            RevitView view,
            XYZ firstPoint,
            XYZ dimensionDirection,
            double dimensionLength)
        {
            var result = new WallCollectionResult();

            List<Wall> walls = new FilteredElementCollector(doc, view.Id)
                .OfClass(typeof(Wall))
                .Cast<Wall>()
                .ToList();

            result.CollectedCount = walls.Count;

            foreach (Wall wall in walls)
            {
                if (wall?.IsValidObject != true)
                {
                    continue;
                }

                // Skip curtain walls
                if (wall.WallType?.Kind == WallKind.Curtain)
                {
                    result.SkippedCurtainCount++;
                    continue;
                }

                // Get wall orientation
                XYZ wallNormal = wall.Orientation;
                if (wallNormal == null)
                {
                    continue;
                }

                // Skip walls parallel to dimension line
                double dotProduct = Math.Abs(wallNormal.DotProduct(dimensionDirection));
                if (dotProduct < (1.0 - ParallelDotTolerance))
                {
                    result.SkippedParallelCount++;
                    continue;
                }

                // Get wall location curve
                LocationCurve locationCurve = wall.Location as LocationCurve;
                if (locationCurve?.Curve is not Line wallLine)
                {
                    // Skip arc walls
                    continue;
                }

                // Project wall midpoint onto dimension line
                XYZ wallMidpoint = (wallLine.GetEndPoint(0) + wallLine.GetEndPoint(1)) * 0.5;
                double parameter = (wallMidpoint - firstPoint).DotProduct(dimensionDirection);
                if (parameter < -ProjectionTolerance || parameter > dimensionLength + ProjectionTolerance)
                {
                    continue;
                }

                // Session 1.2 proven strategy: HostObjectUtils.GetSideFaces with closest-face selection
                Reference wallReference = GetClosestSideFaceReference(wall, firstPoint, dimensionDirection);
                if (wallReference == null)
                {
                    result.SkippedNoFaceReferenceCount++;
                    continue;
                }

                string wallTypeName = wall.WallType?.Name ?? "Unknown";

                result.Candidates.Add(new QuickDimensionFullMixedCandidate(
                    wall.Id,
                    QuickDimensionFullMixedSourceType.Wall,
                    $"Wall: {wallTypeName}",
                    parameter,
                    wallReference));

                result.AcceptedCount++;
            }

            // Deduplicate by ElementId
            var seen = new HashSet<long>();
            result.Candidates.RemoveAll(c => !seen.Add(c.ElementId.Value));

            return result;
        }

        private static Reference GetClosestSideFaceReference(Wall wall, XYZ dimensionLinePoint, XYZ dimensionDirection)
        {
            try
            {
                Document doc = wall.Document;

                IList<Reference> exteriorRefs = HostObjectUtils.GetSideFaces(wall, ShellLayerType.Exterior);
                IList<Reference> interiorRefs = HostObjectUtils.GetSideFaces(wall, ShellLayerType.Interior);

                Reference exteriorRef = exteriorRefs?.Count > 0 ? exteriorRefs[0] : null;
                Reference interiorRef = interiorRefs?.Count > 0 ? interiorRefs[0] : null;

                if (exteriorRef == null && interiorRef == null)
                {
                    return null;
                }

                if (exteriorRef == null) return interiorRef;
                if (interiorRef == null) return exteriorRef;

                // Both faces exist - pick the one closer to the dimension line
                double exteriorDist = GetFaceDistanceToDimensionLine(doc, exteriorRef, dimensionLinePoint, dimensionDirection);
                double interiorDist = GetFaceDistanceToDimensionLine(doc, interiorRef, dimensionLinePoint, dimensionDirection);

                return exteriorDist <= interiorDist ? exteriorRef : interiorRef;
            }
            catch
            {
                // HostObjectUtils may throw for certain wall types
            }

            return null;
        }

        private static double GetFaceDistanceToDimensionLine(
            Document doc,
            Reference faceRef,
            XYZ dimensionLinePoint,
            XYZ dimensionDirection)
        {
            try
            {
                GeometryObject geomObj = doc.GetElement(faceRef).GetGeometryObjectFromReference(faceRef);
                if (geomObj is PlanarFace planarFace)
                {
                    BoundingBoxUV bbox = planarFace.GetBoundingBox();
                    UV midUV = (bbox.Min + bbox.Max) * 0.5;
                    XYZ faceCentroid = planarFace.Evaluate(midUV);

                    XYZ toFace = faceCentroid - dimensionLinePoint;
                    XYZ perpDirection = new XYZ(-dimensionDirection.Y, dimensionDirection.X, 0).Normalize();

                    return Math.Abs(toFace.DotProduct(perpDirection));
                }
            }
            catch
            {
                // Fall back to large distance if calculation fails
            }

            return double.MaxValue;
        }

        #endregion

        #region Door/Window Collection (Session 1.4 strategy)

        private sealed class DoorWindowCollectionResult
        {
            public int CollectedDoorCount { get; set; }
            public int CollectedWindowCount { get; set; }
            public int AcceptedDoorCount { get; set; }
            public int AcceptedWindowCount { get; set; }
            public int SkippedNonHostedCount { get; set; }
            public int SkippedParallelCount { get; set; }
            public int SkippedOutsideSpanCount { get; set; }
            public int SkippedNoReferenceCount { get; set; }
            public List<QuickDimensionFullMixedCandidate> Candidates { get; } = new List<QuickDimensionFullMixedCandidate>();
        }

        private static DoorWindowCollectionResult CollectDoorWindowCandidates(
            Document doc,
            RevitView view,
            XYZ firstPoint,
            XYZ dimensionDirection,
            double dimensionLength)
        {
            var result = new DoorWindowCollectionResult();

            // Collect Doors
            List<FamilyInstance> doors = new FilteredElementCollector(doc, view.Id)
                .OfCategory(BuiltInCategory.OST_Doors)
                .OfClass(typeof(FamilyInstance))
                .Cast<FamilyInstance>()
                .ToList();

            result.CollectedDoorCount = doors.Count;

            foreach (FamilyInstance door in doors)
            {
                var candidates = TryCreateOpeningCandidates(
                    doc, door, QuickDimensionFullMixedSourceType.Door,
                    firstPoint, dimensionDirection, dimensionLength, result);

                if (candidates != null && candidates.Count > 0)
                {
                    result.Candidates.AddRange(candidates);
                    result.AcceptedDoorCount++;
                }
            }

            // Collect Windows
            List<FamilyInstance> windows = new FilteredElementCollector(doc, view.Id)
                .OfCategory(BuiltInCategory.OST_Windows)
                .OfClass(typeof(FamilyInstance))
                .Cast<FamilyInstance>()
                .ToList();

            result.CollectedWindowCount = windows.Count;

            foreach (FamilyInstance window in windows)
            {
                var candidates = TryCreateOpeningCandidates(
                    doc, window, QuickDimensionFullMixedSourceType.Window,
                    firstPoint, dimensionDirection, dimensionLength, result);

                if (candidates != null && candidates.Count > 0)
                {
                    result.Candidates.AddRange(candidates);
                    result.AcceptedWindowCount++;
                }
            }

            return result;
        }

        private static List<QuickDimensionFullMixedCandidate> TryCreateOpeningCandidates(
            Document doc,
            FamilyInstance instance,
            QuickDimensionFullMixedSourceType sourceType,
            XYZ firstPoint,
            XYZ dimensionDirection,
            double dimensionLength,
            DoorWindowCollectionResult result)
        {
            if (instance?.IsValidObject != true)
            {
                return null;
            }

            // Must be hosted in a wall
            Element host = instance.Host;
            if (host is not Wall hostWall)
            {
                result.SkippedNonHostedCount++;
                return null;
            }

            // Get host wall orientation to check if perpendicular to dimension line
            XYZ wallNormal = hostWall.Orientation;
            if (wallNormal == null)
            {
                result.SkippedNonHostedCount++;
                return null;
            }

            // Skip if wall is parallel to dimension line (opening would be perpendicular)
            double dotProduct = Math.Abs(wallNormal.DotProduct(dimensionDirection));
            if (dotProduct < (1.0 - ParallelDotTolerance))
            {
                result.SkippedParallelCount++;
                return null;
            }

            // Get instance location
            LocationPoint locationPoint = instance.Location as LocationPoint;
            if (locationPoint == null)
            {
                result.SkippedNonHostedCount++;
                return null;
            }

            XYZ instanceLocation = locationPoint.Point;

            // Project onto dimension line
            double parameter = (instanceLocation - firstPoint).DotProduct(dimensionDirection);
            if (parameter < -ProjectionTolerance || parameter > dimensionLength + ProjectionTolerance)
            {
                result.SkippedOutsideSpanCount++;
                return null;
            }

            // Session 1.4 proven strategy: FamilyInstance.GetReferences(Left/Right) as primary
            Reference leftRef = null;
            Reference rightRef = null;

            try
            {
                IList<Reference> leftRefs = instance.GetReferences(FamilyInstanceReferenceType.Left);
                if (leftRefs?.Count > 0)
                {
                    leftRef = leftRefs[0];
                }
            }
            catch { }

            try
            {
                IList<Reference> rightRefs = instance.GetReferences(FamilyInstanceReferenceType.Right);
                if (rightRefs?.Count > 0)
                {
                    rightRef = rightRefs[0];
                }
            }
            catch { }

            // Session 1.4 fallback: HostWallOpeningGeometry if FamilyInstance refs not available
            if (leftRef == null || rightRef == null)
            {
                var fallbackRefs = ExtractHostWallOpeningReferences(doc, instance, hostWall, dimensionDirection);
                if (leftRef == null && fallbackRefs.leftRef != null)
                {
                    leftRef = fallbackRefs.leftRef;
                }
                if (rightRef == null && fallbackRefs.rightRef != null)
                {
                    rightRef = fallbackRefs.rightRef;
                }
            }

            if (leftRef == null && rightRef == null)
            {
                result.SkippedNoReferenceCount++;
                return null;
            }

            string familyName = instance.Symbol?.Family?.Name ?? "Unknown";
            string typeName = instance.Symbol?.Name ?? "Unknown";
            string sourceLabel = sourceType == QuickDimensionFullMixedSourceType.Door ? "Door" : "Window";
            string displayName = $"{sourceLabel}: {familyName} - {typeName}";

            var candidates = new List<QuickDimensionFullMixedCandidate>();

            // Calculate left/right positions along dimension line
            // Use instance location as base, offset by half width estimate
            double halfWidth = EstimateOpeningHalfWidth(instance);

            if (leftRef != null)
            {
                double leftParameter = parameter - halfWidth;
                candidates.Add(new QuickDimensionFullMixedCandidate(
                    instance.Id,
                    sourceType,
                    $"{displayName} [Left]",
                    leftParameter,
                    leftRef,
                    hostWall.Id));
            }

            if (rightRef != null)
            {
                double rightParameter = parameter + halfWidth;
                candidates.Add(new QuickDimensionFullMixedCandidate(
                    instance.Id,
                    sourceType,
                    $"{displayName} [Right]",
                    rightParameter,
                    rightRef,
                    hostWall.Id));
            }

            return candidates;
        }

        private static double EstimateOpeningHalfWidth(FamilyInstance instance)
        {
            try
            {
                BoundingBoxXYZ bbox = instance.get_BoundingBox(null);
                if (bbox != null)
                {
                    // Use the larger of X or Y extent as width (depends on orientation)
                    double xExtent = bbox.Max.X - bbox.Min.X;
                    double yExtent = bbox.Max.Y - bbox.Min.Y;
                    return Math.Max(xExtent, yExtent) * 0.5;
                }
            }
            catch { }

            // Default fallback: 1.5 feet (typical door half-width)
            return 1.5;
        }

        private static (Reference leftRef, Reference rightRef) ExtractHostWallOpeningReferences(
            Document doc,
            FamilyInstance instance,
            Wall hostWall,
            XYZ dimensionDirection)
        {
            Reference leftRef = null;
            Reference rightRef = null;

            try
            {
                BoundingBoxXYZ instanceBBox = instance.get_BoundingBox(null);
                if (instanceBBox == null) return (null, null);

                Options options = new Options
                {
                    ComputeReferences = true,
                    IncludeNonVisibleObjects = false
                };

                GeometryElement wallGeom = hostWall.get_Geometry(options);
                if (wallGeom == null) return (null, null);

                List<(Edge edge, double position)> openingEdges = new List<(Edge, double)>();

                foreach (GeometryObject geomObj in wallGeom)
                {
                    if (geomObj is Solid solid)
                    {
                        CollectOpeningEdges(solid, instanceBBox, dimensionDirection, openingEdges);
                    }
                }

                if (openingEdges.Count < 2) return (null, null);

                openingEdges.Sort((a, b) => a.position.CompareTo(b.position));

                Edge leftEdge = openingEdges.First().edge;
                Edge rightEdge = openingEdges.Last().edge;

                leftRef = leftEdge.Reference;
                rightRef = rightEdge.Reference;
            }
            catch { }

            return (leftRef, rightRef);
        }

        private static void CollectOpeningEdges(
            Solid solid,
            BoundingBoxXYZ instanceBBox,
            XYZ dimensionDirection,
            List<(Edge edge, double position)> result)
        {
            double tolerance = 0.5; // feet
            XYZ bboxMin = instanceBBox.Min - new XYZ(tolerance, tolerance, tolerance);
            XYZ bboxMax = instanceBBox.Max + new XYZ(tolerance, tolerance, tolerance);

            foreach (Edge edge in solid.Edges)
            {
                try
                {
                    Curve edgeCurve = edge.AsCurve();
                    if (edgeCurve is not Line edgeLine) continue;

                    XYZ edgeDirection = edgeLine.Direction;
                    if (Math.Abs(edgeDirection.Z) < 0.9) continue;

                    XYZ edgeMidpoint = (edgeLine.GetEndPoint(0) + edgeLine.GetEndPoint(1)) * 0.5;

                    if (edgeMidpoint.X < bboxMin.X || edgeMidpoint.X > bboxMax.X) continue;
                    if (edgeMidpoint.Y < bboxMin.Y || edgeMidpoint.Y > bboxMax.Y) continue;

                    double position = edgeMidpoint.DotProduct(dimensionDirection);
                    result.Add((edge, position));
                }
                catch { }
            }
        }

        #endregion

        #region Test Probing

        private static QuickDimensionFullMixedTestResult ProbeTest(
            Document doc,
            RevitView view,
            Line dimensionLine,
            IReadOnlyList<QuickDimensionFullMixedCandidate> candidates,
            string testName)
        {
            int gridCount = candidates.Count(c => c.SourceType == QuickDimensionFullMixedSourceType.Grid);
            int wallCount = candidates.Count(c => c.SourceType == QuickDimensionFullMixedSourceType.Wall);
            int doorCount = candidates.Count(c => c.SourceType == QuickDimensionFullMixedSourceType.Door);
            int windowCount = candidates.Count(c => c.SourceType == QuickDimensionFullMixedSourceType.Window);

            ReferenceArray references = new ReferenceArray();
            foreach (var candidate in candidates)
            {
                if (candidate.Reference != null)
                {
                    references.Append(candidate.Reference);
                }
            }

            int totalReferences = references.Size;

            if (references.Size < 2)
            {
                return new QuickDimensionFullMixedTestResult(
                    false,
                    gridCount,
                    wallCount,
                    doorCount,
                    windowCount,
                    totalReferences,
                    $"[{testName}] Need at least 2 valid references. Got {references.Size}.");
            }

            using Transaction tx = new Transaction(doc, $"ArcTool: Probe {testName}");
            tx.Start();

            try
            {
                Dimension dimension = doc.Create.NewDimension(view, dimensionLine, references);
                if (dimension == null)
                {
                    tx.RollBack();
                    return new QuickDimensionFullMixedTestResult(
                        false,
                        gridCount,
                        wallCount,
                        doorCount,
                        windowCount,
                        totalReferences,
                        $"[{testName}] NewDimension returned null.");
                }

                tx.RollBack();
                return new QuickDimensionFullMixedTestResult(
                    true,
                    gridCount,
                    wallCount,
                    doorCount,
                    windowCount,
                    totalReferences,
                    $"[{testName}] PASS: NewDimension accepted {totalReferences} refs " +
                    $"({gridCount} grids, {wallCount} walls, {doorCount} doors, {windowCount} windows). " +
                    "Transaction rolled back.");
            }
            catch (Exception ex)
            {
                tx.RollBack();
                return new QuickDimensionFullMixedTestResult(
                    false,
                    gridCount,
                    wallCount,
                    doorCount,
                    windowCount,
                    totalReferences,
                    $"[{testName}] FAIL: {ex.Message}");
            }
        }

        #endregion
    }
}
