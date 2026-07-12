using System;
using System.Collections.Generic;
using System.Linq;
using ArcTool.Core.Models;
using Autodesk.Revit.DB;
using RevitView = Autodesk.Revit.DB.View;

namespace ArcTool.Core.Services
{
    /// <summary>
    /// Session 1.2 spike service: tests wall face reference strategies for NewDimension.
    /// </summary>
    public static class QuickDimensionWallReferenceProbeService
    {
        private const double MinimumDimensionLineLength = 1e-6;
        private const double ParallelDotTolerance = 0.98;
        private const double ProjectionTolerance = 1e-4;

        /// <summary>
        /// Runs the wall face reference probe, testing both HostObjectUtils.GetSideFaces()
        /// and Options.ComputeReferences strategies.
        /// </summary>
        public static QuickDimensionWallProbeSummary RunWallReferenceProbe(
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

            // Collect walls visible in the view
            List<Wall> walls = new FilteredElementCollector(doc, view.Id)
                .OfClass(typeof(Wall))
                .Cast<Wall>()
                .ToList();

            List<QuickDimensionWallCandidate> candidates = new List<QuickDimensionWallCandidate>();
            int skippedCurtainWallCount = 0;
            int skippedParallelWallCount = 0;
            int skippedNoFaceReferenceCount = 0;

            XYZ dimensionDirection = dimensionLine.Direction;
            double dimensionLength = dimensionLine.Length;

            foreach (Wall wall in walls)
            {
                if (wall?.IsValidObject != true)
                {
                    continue;
                }

                // Skip curtain walls - not supported in V1
                if (wall.WallType?.Kind == WallKind.Curtain)
                {
                    skippedCurtainWallCount++;
                    continue;
                }

                // Get wall orientation (normal to the wall face)
                XYZ wallNormal = wall.Orientation;
                if (wallNormal == null)
                {
                    continue;
                }

                // Skip walls parallel to dimension line
                // Wall normal perpendicular to dimension direction means wall is parallel
                double dotProduct = Math.Abs(wallNormal.DotProduct(dimensionDirection));
                if (dotProduct < (1.0 - ParallelDotTolerance))
                {
                    skippedParallelWallCount++;
                    continue;
                }

                // Get wall location curve for projection test
                LocationCurve locationCurve = wall.Location as LocationCurve;
                if (locationCurve?.Curve is not Line wallLine)
                {
                    // Skip arc walls or walls without line-based location
                    continue;
                }

                // Project wall midpoint onto dimension line to check if it's within range
                XYZ wallMidpoint = (wallLine.GetEndPoint(0) + wallLine.GetEndPoint(1)) * 0.5;
                double parameter = (wallMidpoint - firstPoint).DotProduct(dimensionDirection);
                if (parameter < -ProjectionTolerance || parameter > dimensionLength + ProjectionTolerance)
                {
                    continue;
                }

                // Get references using both strategies
                Reference sideFaceRef = GetSideFaceReference(wall, firstPoint, dimensionDirection);
                Reference geometryFaceRef = GetGeometryFaceReference(wall, firstPoint, dimensionDirection);

                if (sideFaceRef == null && geometryFaceRef == null)
                {
                    skippedNoFaceReferenceCount++;
                    continue;
                }

                string wallTypeName = wall.WallType?.Name ?? "Unknown";

                candidates.Add(new QuickDimensionWallCandidate(
                    wall.Id,
                    wallTypeName,
                    parameter,
                    sideFaceRef,
                    geometryFaceRef));
            }

            // Sort by position along dimension line and deduplicate
            candidates = candidates
                .OrderBy(c => c.ParameterOnDimensionLine)
                .GroupBy(c => c.WallId.Value)
                .Select(g => g.First())
                .ToList();

            // Probe both strategies
            QuickDimensionWallStrategyProbeResult sideFacesResult = ProbeStrategy(
                doc,
                view,
                dimensionLine,
                candidates,
                QuickDimensionWallReferenceStrategy.HostObjectUtilsSideFaces);

            QuickDimensionWallStrategyProbeResult geometryResult = ProbeStrategy(
                doc,
                view,
                dimensionLine,
                candidates,
                QuickDimensionWallReferenceStrategy.GeometryComputeReferences);

            return new QuickDimensionWallProbeSummary(
                walls.Count,
                candidates.Count,
                skippedCurtainWallCount,
                skippedParallelWallCount,
                skippedNoFaceReferenceCount,
                sideFacesResult,
                geometryResult);
        }

        private static Line CreateDimensionLine(XYZ firstPoint, XYZ secondPoint)
        {
            if (firstPoint.DistanceTo(secondPoint) < MinimumDimensionLineLength)
            {
                throw new InvalidOperationException("The two picked points are too close to define a dimension line.");
            }

            return Line.CreateBound(firstPoint, secondPoint);
        }

        /// <summary>
        /// Gets a face reference using HostObjectUtils.GetSideFaces().
        /// Returns the face reference closest to the dimension line.
        /// </summary>
        private static Reference GetSideFaceReference(Wall wall, XYZ dimensionLinePoint, XYZ dimensionDirection)
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
        private static double GetFaceDistanceToDimensionLine(Document doc, Reference faceRef, XYZ dimensionLinePoint, XYZ dimensionDirection)
        {
            try
            {
                GeometryObject geomObj = doc.GetElement(faceRef).GetGeometryObjectFromReference(faceRef);
                if (geomObj is PlanarFace planarFace)
                {
                    // Get face centroid by averaging bounding box
                    BoundingBoxUV bbox = planarFace.GetBoundingBox();
                    UV midUV = (bbox.Min + bbox.Max) * 0.5;
                    XYZ faceCentroid = planarFace.Evaluate(midUV);

                    // Calculate perpendicular distance to dimension line
                    // Distance = |projection of (faceCentroid - dimensionLinePoint) onto perpendicular direction|
                    XYZ toFace = faceCentroid - dimensionLinePoint;

                    // Perpendicular direction in XY plane (rotate dimensionDirection 90 degrees)
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

        /// <summary>
        /// Gets a face reference using Options.ComputeReferences = true.
        /// Finds planar faces whose normal aligns with the dimension direction,
        /// then picks the one closest to the dimension line.
        /// </summary>
        private static Reference GetGeometryFaceReference(Wall wall, XYZ dimensionLinePoint, XYZ dimensionDirection)
        {
            try
            {
                Options options = new Options
                {
                    ComputeReferences = true,
                    IncludeNonVisibleObjects = false
                };

                GeometryElement geomElement = wall.get_Geometry(options);
                if (geomElement == null)
                {
                    return null;
                }

                List<(Reference faceRef, XYZ centroid)> candidateFaces = new List<(Reference, XYZ)>();

                foreach (GeometryObject geomObj in geomElement)
                {
                    if (geomObj is Solid solid)
                    {
                        CollectAlignedFaces(solid, dimensionDirection, candidateFaces);
                    }
                    else if (geomObj is GeometryInstance geomInstance)
                    {
                        GeometryElement instanceGeom = geomInstance.GetInstanceGeometry();
                        if (instanceGeom != null)
                        {
                            foreach (GeometryObject instanceObj in instanceGeom)
                            {
                                if (instanceObj is Solid instanceSolid)
                                {
                                    CollectAlignedFaces(instanceSolid, dimensionDirection, candidateFaces);
                                }
                            }
                        }
                    }
                }

                if (candidateFaces.Count == 0)
                {
                    return null;
                }

                // Pick the face closest to the dimension line
                XYZ perpDirection = new XYZ(-dimensionDirection.Y, dimensionDirection.X, 0).Normalize();
                Reference bestRef = null;
                double bestDist = double.MaxValue;

                foreach (var (faceRef, centroid) in candidateFaces)
                {
                    XYZ toFace = centroid - dimensionLinePoint;
                    double dist = Math.Abs(toFace.DotProduct(perpDirection));
                    if (dist < bestDist)
                    {
                        bestDist = dist;
                        bestRef = faceRef;
                    }
                }

                return bestRef;
            }
            catch
            {
                // Geometry extraction may fail for certain elements
            }

            return null;
        }

        /// <summary>
        /// Collects planar faces from a solid whose normals align with the dimension direction.
        /// </summary>
        private static void CollectAlignedFaces(Solid solid, XYZ dimensionDirection, List<(Reference, XYZ)> candidateFaces)
        {
            if (solid == null || solid.Faces == null || solid.Faces.Size == 0)
            {
                return;
            }

            foreach (Face face in solid.Faces)
            {
                if (face is not PlanarFace planarFace)
                {
                    continue;
                }

                Reference faceRef = planarFace.Reference;
                if (faceRef == null)
                {
                    continue;
                }

                // Check if face is vertical (normal is horizontal, Z component ≈ 0)
                // This matches what HostObjectUtils.GetSideFaces returns - the wall's side faces
                XYZ faceNormal = planarFace.FaceNormal;
                bool isVerticalFace = Math.Abs(faceNormal.Z) < 0.1;

                if (isVerticalFace)
                {
                    // Get face centroid
                    BoundingBoxUV bbox = planarFace.GetBoundingBox();
                    UV midUV = (bbox.Min + bbox.Max) * 0.5;
                    XYZ centroid = planarFace.Evaluate(midUV);

                    candidateFaces.Add((faceRef, centroid));
                }
            }
        }

        /// <summary>
        /// Tests a reference strategy by attempting to create a dimension.
        /// </summary>
        private static QuickDimensionWallStrategyProbeResult ProbeStrategy(
            Document doc,
            RevitView view,
            Line dimensionLine,
            IReadOnlyList<QuickDimensionWallCandidate> candidates,
            QuickDimensionWallReferenceStrategy strategy)
        {
            ReferenceArray references = new ReferenceArray();

            foreach (QuickDimensionWallCandidate candidate in candidates)
            {
                Reference reference = strategy == QuickDimensionWallReferenceStrategy.HostObjectUtilsSideFaces
                    ? candidate.SideFaceReference
                    : candidate.GeometryFaceReference;

                if (reference != null)
                {
                    references.Append(reference);
                }
            }

            if (references.Size < 2)
            {
                return new QuickDimensionWallStrategyProbeResult(
                    strategy,
                    false,
                    references.Size,
                    "Need at least 2 valid wall face references.");
            }

            using Transaction tx = new Transaction(doc, $"ArcTool: Probe {strategy}");
            tx.Start();

            try
            {
                Dimension dimension = doc.Create.NewDimension(view, dimensionLine, references);
                if (dimension == null)
                {
                    tx.RollBack();
                    return new QuickDimensionWallStrategyProbeResult(
                        strategy,
                        false,
                        references.Size,
                        "NewDimension returned null.");
                }

                tx.RollBack();
                return new QuickDimensionWallStrategyProbeResult(
                    strategy,
                    true,
                    references.Size,
                    "NewDimension accepted the references. Transaction was rolled back; no dimension was kept in the model.");
            }
            catch (Exception ex)
            {
                tx.RollBack();
                return new QuickDimensionWallStrategyProbeResult(
                    strategy,
                    false,
                    references.Size,
                    ex.Message);
            }
        }
    }
}
