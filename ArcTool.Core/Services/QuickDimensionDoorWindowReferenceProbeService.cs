using System;
using System.Collections.Generic;
using System.Linq;
using ArcTool.Core.Models;
using Autodesk.Revit.DB;
using RevitView = Autodesk.Revit.DB.View;

namespace ArcTool.Core.Services
{
    /// <summary>
    /// Session 1.4 spike service: tests Door/Window opening reference strategies for NewDimension.
    /// Tests three strategies:
    /// 1. FamilyInstance.GetReferences(FamilyInstanceReferenceType.Left/Right)
    /// 2. Options.ComputeReferences = true + Face.Reference
    /// 3. Host Wall FindInserts + Opening Geometry
    /// </summary>
    public static class QuickDimensionDoorWindowReferenceProbeService
    {
        private const double MinimumDimensionLineLength = 1e-6;
        private const double ParallelDotTolerance = 0.98;
        private const double ProjectionTolerance = 1e-4;

        /// <summary>
        /// Runs the Door/Window reference probe, testing all three strategies.
        /// </summary>
        public static QuickDimensionDoorWindowProbeSummary RunDoorWindowReferenceProbe(
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

            // Collect Door and Window candidates
            var collectionResult = CollectDoorWindowCandidates(
                doc, view, firstPoint, dimensionDirection, dimensionLength);

            // Extract references using all three strategies
            foreach (var candidate in collectionResult.Candidates)
            {
                ExtractFamilyInstanceReferences(doc, candidate);
                ExtractGeometryComputeReferences(doc, candidate, dimensionDirection);
                ExtractHostWallOpeningReferences(doc, candidate, dimensionDirection);
            }

            // Probe each strategy
            var familyInstanceResult = ProbeStrategy(
                doc, view, dimensionLine, collectionResult.Candidates,
                QuickDimensionDoorWindowReferenceStrategy.FamilyInstanceReferences);

            var geometryResult = ProbeStrategy(
                doc, view, dimensionLine, collectionResult.Candidates,
                QuickDimensionDoorWindowReferenceStrategy.GeometryComputeReferences);

            var openingGeometryResult = ProbeStrategy(
                doc, view, dimensionLine, collectionResult.Candidates,
                QuickDimensionDoorWindowReferenceStrategy.HostWallOpeningGeometry);

            return new QuickDimensionDoorWindowProbeSummary(
                collectionResult.CollectedDoorCount,
                collectionResult.CollectedWindowCount,
                collectionResult.AcceptedDoorCount,
                collectionResult.AcceptedWindowCount,
                collectionResult.SkippedNonHostedCount,
                collectionResult.SkippedParallelCount,
                collectionResult.SkippedOutsideSpanCount,
                familyInstanceResult,
                geometryResult,
                openingGeometryResult);
        }

        private static Line CreateDimensionLine(XYZ firstPoint, XYZ secondPoint)
        {
            if (firstPoint.DistanceTo(secondPoint) < MinimumDimensionLineLength)
            {
                throw new InvalidOperationException("The two picked points are too close to define a dimension line.");
            }

            return Line.CreateBound(firstPoint, secondPoint);
        }

        #region Candidate Collection

        private sealed class DoorWindowCollectionResult
        {
            public int CollectedDoorCount { get; set; }
            public int CollectedWindowCount { get; set; }
            public int AcceptedDoorCount { get; set; }
            public int AcceptedWindowCount { get; set; }
            public int SkippedNonHostedCount { get; set; }
            public int SkippedParallelCount { get; set; }
            public int SkippedOutsideSpanCount { get; set; }
            public List<QuickDimensionDoorWindowCandidate> Candidates { get; } = new List<QuickDimensionDoorWindowCandidate>();
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
                var candidate = TryCreateCandidate(
                    door, QuickDimensionDoorWindowSourceType.Door,
                    firstPoint, dimensionDirection, dimensionLength, result);

                if (candidate != null)
                {
                    result.Candidates.Add(candidate);
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
                var candidate = TryCreateCandidate(
                    window, QuickDimensionDoorWindowSourceType.Window,
                    firstPoint, dimensionDirection, dimensionLength, result);

                if (candidate != null)
                {
                    result.Candidates.Add(candidate);
                    result.AcceptedWindowCount++;
                }
            }

            // Sort by position along dimension line
            result.Candidates.Sort((a, b) => a.ParameterOnDimensionLine.CompareTo(b.ParameterOnDimensionLine));

            return result;
        }

        private static QuickDimensionDoorWindowCandidate TryCreateCandidate(
            FamilyInstance instance,
            QuickDimensionDoorWindowSourceType sourceType,
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

            string familyName = instance.Symbol?.Family?.Name ?? "Unknown";
            string typeName = instance.Symbol?.Name ?? "Unknown";

            return new QuickDimensionDoorWindowCandidate(
                instance.Id,
                sourceType,
                familyName,
                typeName,
                parameter,
                hostWall.Id);
        }

        #endregion

        #region Strategy 1: FamilyInstance.GetReferences

        /// <summary>
        /// Extracts references using FamilyInstance.GetReferences(FamilyInstanceReferenceType.Left/Right).
        /// This is the recommended API approach if the family has properly named reference planes.
        /// </summary>
        private static void ExtractFamilyInstanceReferences(Document doc, QuickDimensionDoorWindowCandidate candidate)
        {
            try
            {
                FamilyInstance instance = doc.GetElement(candidate.ElementId) as FamilyInstance;
                if (instance == null) return;

                // Try to get Left reference
                try
                {
                    IList<Reference> leftRefs = instance.GetReferences(FamilyInstanceReferenceType.Left);
                    if (leftRefs?.Count > 0)
                    {
                        candidate.LeftFamilyInstanceReference = leftRefs[0];
                    }
                }
                catch
                {
                    // Family may not have Left reference plane
                }

                // Try to get Right reference
                try
                {
                    IList<Reference> rightRefs = instance.GetReferences(FamilyInstanceReferenceType.Right);
                    if (rightRefs?.Count > 0)
                    {
                        candidate.RightFamilyInstanceReference = rightRefs[0];
                    }
                }
                catch
                {
                    // Family may not have Right reference plane
                }
            }
            catch
            {
                // GetReferences may throw for certain family types
            }
        }

        #endregion

        #region Strategy 2: Geometry ComputeReferences

        /// <summary>
        /// Extracts references using Options.ComputeReferences = true and Face.Reference.
        /// This is a general approach that doesn't depend on family definition.
        /// </summary>
        private static void ExtractGeometryComputeReferences(
            Document doc,
            QuickDimensionDoorWindowCandidate candidate,
            XYZ dimensionDirection)
        {
            try
            {
                FamilyInstance instance = doc.GetElement(candidate.ElementId) as FamilyInstance;
                if (instance == null) return;

                Options options = new Options
                {
                    ComputeReferences = true,
                    IncludeNonVisibleObjects = false,
                    View = doc.ActiveView
                };

                GeometryElement geomElement = instance.get_Geometry(options);
                if (geomElement == null) return;

                // Collect all planar faces perpendicular to dimension direction
                List<(PlanarFace face, double position)> perpendicularFaces = new List<(PlanarFace, double)>();

                foreach (GeometryObject geomObj in geomElement)
                {
                    CollectPerpendicularFaces(geomObj, dimensionDirection, perpendicularFaces);
                }

                if (perpendicularFaces.Count < 2) return;

                // Sort by position along dimension direction
                perpendicularFaces.Sort((a, b) => a.position.CompareTo(b.position));

                // Take leftmost and rightmost faces
                PlanarFace leftFace = perpendicularFaces.First().face;
                PlanarFace rightFace = perpendicularFaces.Last().face;

                if (leftFace.Reference != null)
                {
                    candidate.LeftGeometryReference = leftFace.Reference;
                }

                if (rightFace.Reference != null)
                {
                    candidate.RightGeometryReference = rightFace.Reference;
                }
            }
            catch
            {
                // Geometry extraction may fail for certain instances
            }
        }

        private static void CollectPerpendicularFaces(
            GeometryObject geomObj,
            XYZ dimensionDirection,
            List<(PlanarFace face, double position)> result)
        {
            if (geomObj is Solid solid)
            {
                foreach (Face face in solid.Faces)
                {
                    if (face is PlanarFace planarFace)
                    {
                        XYZ faceNormal = planarFace.FaceNormal;
                        double dot = Math.Abs(faceNormal.DotProduct(dimensionDirection));

                        // Face is perpendicular to dimension direction if normal is parallel
                        if (dot > ParallelDotTolerance)
                        {
                            BoundingBoxUV bbox = planarFace.GetBoundingBox();
                            UV midUV = (bbox.Min + bbox.Max) * 0.5;
                            XYZ faceCentroid = planarFace.Evaluate(midUV);
                            double position = faceCentroid.DotProduct(dimensionDirection);

                            result.Add((planarFace, position));
                        }
                    }
                }
            }
            else if (geomObj is GeometryInstance geomInstance)
            {
                GeometryElement instanceGeom = geomInstance.GetInstanceGeometry();
                if (instanceGeom != null)
                {
                    foreach (GeometryObject innerObj in instanceGeom)
                    {
                        CollectPerpendicularFaces(innerObj, dimensionDirection, result);
                    }
                }
            }
        }

        #endregion

        #region Strategy 3: Host Wall Opening Geometry

        /// <summary>
        /// Extracts references from the host wall's opening cut geometry.
        /// Uses Wall.FindInserts() to identify openings and extracts edge references.
        /// </summary>
        private static void ExtractHostWallOpeningReferences(
            Document doc,
            QuickDimensionDoorWindowCandidate candidate,
            XYZ dimensionDirection)
        {
            try
            {
                Wall hostWall = doc.GetElement(candidate.HostWallId) as Wall;
                if (hostWall == null) return;

                FamilyInstance instance = doc.GetElement(candidate.ElementId) as FamilyInstance;
                if (instance == null) return;

                // Get the instance's bounding box to identify its opening region
                BoundingBoxXYZ instanceBBox = instance.get_BoundingBox(null);
                if (instanceBBox == null) return;

                // Get wall geometry with openings
                Options options = new Options
                {
                    ComputeReferences = true,
                    IncludeNonVisibleObjects = false
                };

                GeometryElement wallGeom = hostWall.get_Geometry(options);
                if (wallGeom == null) return;

                // Find edges at the opening boundaries
                List<(Edge edge, double position)> openingEdges = new List<(Edge, double)>();

                foreach (GeometryObject geomObj in wallGeom)
                {
                    if (geomObj is Solid solid)
                    {
                        CollectOpeningEdges(solid, instanceBBox, dimensionDirection, openingEdges);
                    }
                }

                if (openingEdges.Count < 2) return;

                // Sort by position
                openingEdges.Sort((a, b) => a.position.CompareTo(b.position));

                // Take leftmost and rightmost edges within the instance's bounding box
                Edge leftEdge = openingEdges.First().edge;
                Edge rightEdge = openingEdges.Last().edge;

                if (leftEdge.Reference != null)
                {
                    candidate.LeftOpeningReference = leftEdge.Reference;
                }

                if (rightEdge.Reference != null)
                {
                    candidate.RightOpeningReference = rightEdge.Reference;
                }
            }
            catch
            {
                // Opening geometry extraction may fail
            }
        }

        private static void CollectOpeningEdges(
            Solid solid,
            BoundingBoxXYZ instanceBBox,
            XYZ dimensionDirection,
            List<(Edge edge, double position)> result)
        {
            // Expand instance bbox slightly for tolerance
            double tolerance = 0.5; // feet
            XYZ bboxMin = instanceBBox.Min - new XYZ(tolerance, tolerance, tolerance);
            XYZ bboxMax = instanceBBox.Max + new XYZ(tolerance, tolerance, tolerance);

            foreach (Edge edge in solid.Edges)
            {
                try
                {
                    Curve edgeCurve = edge.AsCurve();
                    if (edgeCurve is not Line edgeLine) continue;

                    // Check if edge is vertical (perpendicular to XY plane)
                    XYZ edgeDirection = edgeLine.Direction;
                    if (Math.Abs(edgeDirection.Z) < 0.9) continue;

                    // Check if edge is within instance's bounding box region
                    XYZ edgeMidpoint = (edgeLine.GetEndPoint(0) + edgeLine.GetEndPoint(1)) * 0.5;

                    if (edgeMidpoint.X < bboxMin.X || edgeMidpoint.X > bboxMax.X) continue;
                    if (edgeMidpoint.Y < bboxMin.Y || edgeMidpoint.Y > bboxMax.Y) continue;

                    double position = edgeMidpoint.DotProduct(dimensionDirection);
                    result.Add((edge, position));
                }
                catch
                {
                    // Skip problematic edges
                }
            }
        }

        #endregion

        #region Strategy Probing

        /// <summary>
        /// Tests a strategy by attempting to create a dimension with the extracted references.
        /// </summary>
        private static QuickDimensionDoorWindowStrategyResult ProbeStrategy(
            Document doc,
            RevitView view,
            Line dimensionLine,
            IReadOnlyList<QuickDimensionDoorWindowCandidate> candidates,
            QuickDimensionDoorWindowReferenceStrategy strategy)
        {
            int totalCandidates = candidates.Count;
            int candidatesWithRefs = candidates.Count(c => c.HasReferences(strategy));

            // Build reference array with both left and right references
            ReferenceArray references = new ReferenceArray();
            foreach (var candidate in candidates)
            {
                Reference leftRef = candidate.GetLeftReference(strategy);
                Reference rightRef = candidate.GetRightReference(strategy);

                if (leftRef != null)
                {
                    references.Append(leftRef);
                }
                if (rightRef != null)
                {
                    references.Append(rightRef);
                }
            }

            int referencesUsed = references.Size;

            if (references.Size < 2)
            {
                return new QuickDimensionDoorWindowStrategyResult(
                    strategy,
                    false,
                    totalCandidates,
                    candidatesWithRefs,
                    referencesUsed,
                    $"Need at least 2 valid references. Got {references.Size}.");
            }

            using Transaction tx = new Transaction(doc, $"ArcTool: Probe {strategy}");
            tx.Start();

            try
            {
                Dimension dimension = doc.Create.NewDimension(view, dimensionLine, references);
                if (dimension == null)
                {
                    tx.RollBack();
                    return new QuickDimensionDoorWindowStrategyResult(
                        strategy,
                        false,
                        totalCandidates,
                        candidatesWithRefs,
                        referencesUsed,
                        "NewDimension returned null.");
                }

                tx.RollBack();
                return new QuickDimensionDoorWindowStrategyResult(
                    strategy,
                    true,
                    totalCandidates,
                    candidatesWithRefs,
                    referencesUsed,
                    $"NewDimension accepted {references.Size} references from {candidatesWithRefs} candidates. Transaction rolled back.");
            }
            catch (Exception ex)
            {
                tx.RollBack();
                return new QuickDimensionDoorWindowStrategyResult(
                    strategy,
                    false,
                    totalCandidates,
                    candidatesWithRefs,
                    referencesUsed,
                    ex.Message);
            }
        }

        #endregion
    }
}
