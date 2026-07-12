using System;
using System.Collections.Generic;
using System.Linq;
using ArcTool.Core.Models;
using Autodesk.Revit.DB;
using RevitView = Autodesk.Revit.DB.View;

namespace ArcTool.Core.Services
{
    /// <summary>
    /// Session 1.3 spike service: tests mixed Grid + Wall reference arrays for NewDimension.
    /// Uses proven strategies from Session 1.1 (Grid: new Reference(grid)) and
    /// Session 1.2 (Wall: HostObjectUtils.GetSideFaces with closest-face selection).
    /// </summary>
    public static class QuickDimensionMixedReferenceProbeService
    {
        private const double MinimumDimensionLineLength = 1e-6;
        private const double ParallelDotTolerance = 0.98;
        private const double ProjectionTolerance = 1e-4;

        /// <summary>
        /// Runs the mixed reference probe, testing Grid + Wall references in the same ReferenceArray.
        /// </summary>
        public static QuickDimensionMixedProbeSummary RunMixedReferenceProbe(
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

            // Collect Grid candidates
            var gridCollectionResult = CollectGridCandidates(doc, view, firstPoint, dimensionDirection, dimensionLength);

            // Collect Wall candidates
            var wallCollectionResult = CollectWallCandidates(doc, view, firstPoint, dimensionDirection, dimensionLength);

            // Merge all candidates into unified list
            List<QuickDimensionMixedCandidate> allCandidates = new List<QuickDimensionMixedCandidate>();
            allCandidates.AddRange(gridCollectionResult.Candidates);
            allCandidates.AddRange(wallCollectionResult.Candidates);

            // Sort by position along dimension line
            List<QuickDimensionMixedCandidate> sortedCandidates = allCandidates
                .OrderBy(c => c.ParameterOnDimensionLine)
                .ToList();

            // Probe all scenarios
            QuickDimensionMixedScenarioResult sortedResult = ProbeScenario(
                doc, view, dimensionLine, sortedCandidates,
                QuickDimensionMixedTestScenario.SortedByPosition);

            QuickDimensionMixedScenarioResult reversedResult = ProbeScenario(
                doc, view, dimensionLine, sortedCandidates.AsEnumerable().Reverse().ToList(),
                QuickDimensionMixedTestScenario.ReversedOrder);

            QuickDimensionMixedScenarioResult gridsOnlyResult = ProbeScenario(
                doc, view, dimensionLine,
                sortedCandidates.Where(c => c.SourceType == QuickDimensionMixedSourceType.Grid).ToList(),
                QuickDimensionMixedTestScenario.GridsOnly);

            QuickDimensionMixedScenarioResult wallsOnlyResult = ProbeScenario(
                doc, view, dimensionLine,
                sortedCandidates.Where(c => c.SourceType == QuickDimensionMixedSourceType.Wall).ToList(),
                QuickDimensionMixedTestScenario.WallsOnly);

            return new QuickDimensionMixedProbeSummary(
                gridCollectionResult.CollectedCount,
                wallCollectionResult.CollectedCount,
                gridCollectionResult.Candidates.Count,
                wallCollectionResult.Candidates.Count,
                gridCollectionResult.SkippedArcCount,
                gridCollectionResult.SkippedParallelCount,
                wallCollectionResult.SkippedCurtainCount,
                wallCollectionResult.SkippedParallelCount,
                wallCollectionResult.SkippedNoFaceReferenceCount,
                sortedResult,
                reversedResult,
                gridsOnlyResult,
                wallsOnlyResult);
        }

        private static Line CreateDimensionLine(XYZ firstPoint, XYZ secondPoint)
        {
            if (firstPoint.DistanceTo(secondPoint) < MinimumDimensionLineLength)
            {
                throw new InvalidOperationException("The two picked points are too close to define a dimension line.");
            }

            return Line.CreateBound(firstPoint, secondPoint);
        }

        #region Grid Collection

        private sealed class GridCollectionResult
        {
            public int CollectedCount { get; set; }
            public int SkippedArcCount { get; set; }
            public int SkippedParallelCount { get; set; }
            public List<QuickDimensionMixedCandidate> Candidates { get; } = new List<QuickDimensionMixedCandidate>();
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

                result.Candidates.Add(new QuickDimensionMixedCandidate(
                    grid.Id,
                    QuickDimensionMixedSourceType.Grid,
                    grid.Name,
                    parameter,
                    gridReference));
            }

            // Deduplicate by ElementId
            result.Candidates.RemoveAll(c =>
                result.Candidates.Any(other =>
                    other.ElementId.Value == c.ElementId.Value &&
                    other.ParameterOnDimensionLine < c.ParameterOnDimensionLine));

            return result;
        }

        #endregion

        #region Wall Collection

        private sealed class WallCollectionResult
        {
            public int CollectedCount { get; set; }
            public int SkippedCurtainCount { get; set; }
            public int SkippedParallelCount { get; set; }
            public int SkippedNoFaceReferenceCount { get; set; }
            public List<QuickDimensionMixedCandidate> Candidates { get; } = new List<QuickDimensionMixedCandidate>();
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

                result.Candidates.Add(new QuickDimensionMixedCandidate(
                    wall.Id,
                    QuickDimensionMixedSourceType.Wall,
                    wallTypeName,
                    parameter,
                    wallReference));
            }

            // Deduplicate by ElementId
            var seen = new HashSet<long>();
            result.Candidates.RemoveAll(c => !seen.Add(c.ElementId.Value));

            return result;
        }

        /// <summary>
        /// Gets the closest side face reference using HostObjectUtils.GetSideFaces().
        /// Session 1.2 locked strategy: pick the face whose centroid is closer to the dimension line.
        /// </summary>
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

        /// <summary>
        /// Calculates the perpendicular distance from a face's centroid to the dimension line.
        /// </summary>
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

        #region Scenario Probing

        /// <summary>
        /// Tests a scenario by attempting to create a dimension with the given candidates.
        /// </summary>
        private static QuickDimensionMixedScenarioResult ProbeScenario(
            Document doc,
            RevitView view,
            Line dimensionLine,
            IReadOnlyList<QuickDimensionMixedCandidate> candidates,
            QuickDimensionMixedTestScenario scenario)
        {
            int gridCount = candidates.Count(c => c.SourceType == QuickDimensionMixedSourceType.Grid);
            int wallCount = candidates.Count(c => c.SourceType == QuickDimensionMixedSourceType.Wall);

            ReferenceArray references = new ReferenceArray();
            foreach (QuickDimensionMixedCandidate candidate in candidates)
            {
                if (candidate.Reference != null)
                {
                    references.Append(candidate.Reference);
                }
            }

            if (references.Size < 2)
            {
                return new QuickDimensionMixedScenarioResult(
                    scenario,
                    false,
                    gridCount,
                    wallCount,
                    $"Need at least 2 valid references. Got {references.Size}.");
            }

            using Transaction tx = new Transaction(doc, $"ArcTool: Probe {scenario}");
            tx.Start();

            try
            {
                Dimension dimension = doc.Create.NewDimension(view, dimensionLine, references);
                if (dimension == null)
                {
                    tx.RollBack();
                    return new QuickDimensionMixedScenarioResult(
                        scenario,
                        false,
                        gridCount,
                        wallCount,
                        "NewDimension returned null.");
                }

                tx.RollBack();
                return new QuickDimensionMixedScenarioResult(
                    scenario,
                    true,
                    gridCount,
                    wallCount,
                    $"NewDimension accepted {references.Size} references ({gridCount} grids, {wallCount} walls). Transaction rolled back.");
            }
            catch (Exception ex)
            {
                tx.RollBack();
                return new QuickDimensionMixedScenarioResult(
                    scenario,
                    false,
                    gridCount,
                    wallCount,
                    ex.Message);
            }
        }

        #endregion
    }
}
