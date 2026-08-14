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
    /// Production read-only Wall boundary candidate collector for Quick Dimension.
    /// This service performs no transactions and never creates dimensions.
    /// </summary>
    public static class QuickDimensionWallCandidateCollector
    {
        private const double VerticalFaceNormalZTolerance = 1e-3;

        /// <summary>
        /// Collects Wall side-face boundary candidates visible in the supplied plan view.
        /// The collector uses true 2D side-face segment/dimension-line intersection instead of midpoint projection.
        /// </summary>
        public static QuickDimensionReadOnlyResult CollectWallCandidates(
            Document doc,
            RevitView view,
            QuickDimensionLineContext lineContext,
            QuickDimensionOptions? options = null)
        {
            if (doc == null) throw new ArgumentNullException(nameof(doc));
            if (view == null) throw new ArgumentNullException(nameof(view));
            if (lineContext == null) throw new ArgumentNullException(nameof(lineContext));

            QuickDimensionOptions effectiveOptions = options ?? QuickDimensionOptions.Default;
            var diagnostics = new List<QuickDimensionDiagnostic>();
            var rawCandidates = new List<QuickDimensionCandidate>();
            int collectedCount = 0;

            if (!effectiveOptions.IncludesSource(QuickDimensionSourceType.Wall))
            {
                diagnostics.Add(new QuickDimensionDiagnostic(
                    QuickDimensionDiagnosticSeverity.Info,
                    QuickDimensionRejectedReason.None,
                    "Wall collection is disabled by Quick Dimension options.",
                    sourceType: QuickDimensionSourceType.Wall));

                return BuildResult(lineContext, rawCandidates, diagnostics, collectedCount, effectiveOptions);
            }

            if (!IsSupportedPlanView(view))
            {
                diagnostics.Add(new QuickDimensionDiagnostic(
                    QuickDimensionDiagnosticSeverity.Error,
                    QuickDimensionRejectedReason.UnsupportedView,
                    $"Quick Dimension Wall collection supports active plan views only. Current view type: {view.ViewType}.",
                    view.Id,
                    QuickDimensionSourceType.Wall,
                    view.Name));

                return BuildResult(lineContext, rawCandidates, diagnostics, collectedCount, effectiveOptions);
            }

            List<Wall> walls = new FilteredElementCollector(doc, view.Id)
                .OfClass(typeof(Wall))
                .OfCategory(BuiltInCategory.OST_Walls)
                .Cast<Wall>()
                .ToList();

            collectedCount = walls.Count;

            foreach (Wall wall in walls)
            {
                if (wall?.IsValidObject != true)
                {
                    diagnostics.Add(new QuickDimensionDiagnostic(
                        QuickDimensionDiagnosticSeverity.Warning,
                        QuickDimensionRejectedReason.InvalidGeometry,
                        "Skipped an invalid Wall object returned by the active-view collector.",
                        sourceType: QuickDimensionSourceType.Wall));
                    continue;
                }

                TryCollectWallCandidate(wall, lineContext, effectiveOptions, rawCandidates, diagnostics);
            }

            return BuildResult(lineContext, rawCandidates, diagnostics, collectedCount, effectiveOptions);
        }

        /// <summary>
        /// Wall-axis projection model (Session 2.7): resolves the selected wall's two picked-side anchors
        /// from the validated directional side-face boundary model. Wall end-cap faces are no longer used here.
        /// </summary>
        public static QuickDimensionReadOnlyResult CollectSelectedWallEndAnchors(
            Document doc,
            QuickDimensionLineContext lineContext,
            Wall selectedWall,
            QuickDimensionOptions? options = null)
        {
            if (doc == null) throw new ArgumentNullException(nameof(doc));
            if (lineContext == null) throw new ArgumentNullException(nameof(lineContext));
            if (selectedWall == null) throw new ArgumentNullException(nameof(selectedWall));

            QuickDimensionOptions effectiveOptions = options ?? QuickDimensionOptions.Default;
            var diagnostics = new List<QuickDimensionDiagnostic>();
            var rawCandidates = new List<QuickDimensionCandidate>();
            int collectedCount = selectedWall.IsValidObject ? 1 : 0;

            if (!effectiveOptions.IncludesSource(QuickDimensionSourceType.Wall))
            {
                diagnostics.Add(new QuickDimensionDiagnostic(
                    QuickDimensionDiagnosticSeverity.Info,
                    QuickDimensionRejectedReason.None,
                    "Wall collection is disabled by Quick Dimension options.",
                    sourceType: QuickDimensionSourceType.Wall));

                return BuildResult(lineContext, rawCandidates, diagnostics, collectedCount, effectiveOptions);
            }

            TryCollectWallEndAnchors(selectedWall, lineContext, rawCandidates, diagnostics);

            return BuildResult(lineContext, rawCandidates, diagnostics, collectedCount, effectiveOptions);
        }

        private static void TryCollectWallEndAnchors(
            Wall wall,
            QuickDimensionLineContext lineContext,
            List<QuickDimensionCandidate> candidates,
            List<QuickDimensionDiagnostic> diagnostics)
        {
            string displayName = BuildWallDisplayName(wall);

            try
            {
                if (lineContext.SideSign == 0 || lineContext.SideNormal == null)
                {
                    diagnostics.Add(new QuickDimensionDiagnostic(
                        QuickDimensionDiagnosticSeverity.Warning,
                        QuickDimensionRejectedReason.InvalidGeometry,
                        $"Skipped Wall '{displayName}' because the wall-axis side pick did not resolve to a placement side.",
                        wall.Id, QuickDimensionSourceType.Wall, displayName));
                    return;
                }

                // SideNormal is guaranteed non-null here: we checked SideSign != 0 above.
                // Adding the unit perpendicular vector to FirstPoint (= wall start) places the
                // pick point clearly on the correct side for GetSide() cross-product computation.
                XYZ sideNormal = lineContext.SideNormal!;
                XYZ sidePickPoint = lineContext.FirstPoint + sideNormal;
                QuickDimensionWallSpikeResult spikeResult = QuickDimensionWallReferenceProbeService.RunWallReferenceProbe(
                    wall,
                    sidePickPoint);

                if (!spikeResult.Succeeded || spikeResult.StartAnchor == null || spikeResult.FinishAnchor == null)
                {
                    diagnostics.Add(new QuickDimensionDiagnostic(
                        QuickDimensionDiagnosticSeverity.Warning,
                        QuickDimensionRejectedReason.MissingReference,
                        $"Skipped Wall '{displayName}' because side-face anchor resolution failed: {spikeResult.Message}",
                        wall.Id, QuickDimensionSourceType.Wall, displayName));
                    return;
                }

                if (spikeResult.StartAnchor.EdgeReference == null || spikeResult.FinishAnchor.EdgeReference == null)
                {
                    diagnostics.Add(new QuickDimensionDiagnostic(
                        QuickDimensionDiagnosticSeverity.Warning,
                        QuickDimensionRejectedReason.MissingReference,
                        $"Skipped Wall '{displayName}' because side-face anchor resolution did not return usable Edge.Reference values.",
                        wall.Id, QuickDimensionSourceType.Wall, displayName));
                    return;
                }

                string typeName = wall.WallType?.Name ?? string.Empty;

                // BUG-09 fix: the anchor's FinalCandidate.elementId must be the true owner of the
                // resolved Reference, not the selected wall. On extends-outward (Exterior shell)
                // anchors the reference lives on a joining wall, so hardcoding wall.Id mismatched the
                // stableReference owner. Pass the reference owner (EdgeReference.ElementId) as the
                // candidate ElementId and keep wall.Id as HostElementId to preserve the
                // "anchor belongs to which wall's chain" link. EdgeReference is non-null here
                // (guarded above), so ElementId access is safe.
                ElementId startAnchorOwnerId = spikeResult.StartAnchor.EdgeReference.ElementId;
                ElementId finishAnchorOwnerId = spikeResult.FinishAnchor.EdgeReference.ElementId;

                candidates.Add(new QuickDimensionCandidate(
                    startAnchorOwnerId, QuickDimensionSourceType.Wall,
                    $"{displayName} [Start Anchor]",
                    spikeResult.StartAnchor.EdgeReference, QuickDimensionReferenceStrategy.WallSideFace,
                    spikeResult.StartAnchor.Midpoint, spikeResult.StartAnchor.ParameterOnWallAxis,
                    hostElementId: wall.Id,
                    typeName: typeName));

                candidates.Add(new QuickDimensionCandidate(
                    finishAnchorOwnerId, QuickDimensionSourceType.Wall,
                    $"{displayName} [Finish Anchor]",
                    spikeResult.FinishAnchor.EdgeReference, QuickDimensionReferenceStrategy.WallSideFace,
                    spikeResult.FinishAnchor.Midpoint, spikeResult.FinishAnchor.ParameterOnWallAxis,
                    hostElementId: wall.Id,
                    typeName: typeName));
            }
            catch (Exception ex)
            {
                diagnostics.Add(new QuickDimensionDiagnostic(
                    QuickDimensionDiagnosticSeverity.Warning,
                    QuickDimensionRejectedReason.CollectorException,
                    $"Skipped Wall '{displayName}' because the collector caught an API exception: {ex.Message}",
                    wall.Id, QuickDimensionSourceType.Wall, displayName));
            }
        }

        private static void TryCollectWallCandidate(
            Wall wall,
            QuickDimensionLineContext lineContext,
            QuickDimensionOptions options,
            List<QuickDimensionCandidate> candidates,
            List<QuickDimensionDiagnostic> diagnostics)
        {
            string displayName = BuildWallDisplayName(wall);

            try
            {
                if (wall.WallType?.Kind == WallKind.Curtain)
                {
                    diagnostics.Add(new QuickDimensionDiagnostic(
                        QuickDimensionDiagnosticSeverity.Warning,
                        QuickDimensionRejectedReason.CurtainWallUnsupported,
                        $"Skipped Wall '{displayName}' because curtain walls are outside the Quick Dimension MVP scope.",
                        wall.Id,
                        QuickDimensionSourceType.Wall,
                        displayName));
                    return;
                }

                if (wall.Location is not LocationCurve locationCurve || locationCurve.Curve == null)
                {
                    diagnostics.Add(new QuickDimensionDiagnostic(
                        QuickDimensionDiagnosticSeverity.Warning,
                        QuickDimensionRejectedReason.InvalidGeometry,
                        $"Skipped Wall '{displayName}' because it does not expose a valid LocationCurve.",
                        wall.Id,
                        QuickDimensionSourceType.Wall,
                        displayName));
                    return;
                }

                Curve wallCurve = locationCurve.Curve;
                if (wallCurve is not Line)
                {
                    diagnostics.Add(new QuickDimensionDiagnostic(
                        QuickDimensionDiagnosticSeverity.Warning,
                        QuickDimensionRejectedReason.ArcWallUnsupported,
                        $"Skipped Wall '{displayName}' because arc or non-line walls are outside the Quick Dimension MVP scope.",
                        wall.Id,
                        QuickDimensionSourceType.Wall,
                        displayName));
                    return;
                }

                if (!QuickDimensionGeometryService.TryGetStraightCurveEndpoints(wallCurve, out XYZ startPoint, out XYZ endPoint))
                {
                    diagnostics.Add(new QuickDimensionDiagnostic(
                        QuickDimensionDiagnosticSeverity.Warning,
                        QuickDimensionRejectedReason.InvalidGeometry,
                        $"Skipped Wall '{displayName}' because its curve endpoints are not valid finite points.",
                        wall.Id,
                        QuickDimensionSourceType.Wall,
                        displayName));
                    return;
                }

                if (!QuickDimensionGeometryService.TryGetPlanarDirection(startPoint, endPoint, options.ProjectionTolerance, out XYZ wallDirection))
                {
                    diagnostics.Add(new QuickDimensionDiagnostic(
                        QuickDimensionDiagnosticSeverity.Warning,
                        QuickDimensionRejectedReason.InvalidGeometry,
                        $"Skipped Wall '{displayName}' because its curve does not define a valid plan-view direction.",
                        wall.Id,
                        QuickDimensionSourceType.Wall,
                        displayName));
                    return;
                }

                if (QuickDimensionGeometryService.IsNearlyParallel(lineContext.Direction, wallDirection, options.ProjectionTolerance))
                {
                    diagnostics.Add(new QuickDimensionDiagnostic(
                        QuickDimensionDiagnosticSeverity.Warning,
                        QuickDimensionRejectedReason.ParallelToDimensionLine,
                        $"Skipped Wall '{displayName}' because it is parallel to the picked dimension line.",
                        wall.Id,
                        QuickDimensionSourceType.Wall,
                        displayName));
                    return;
                }

                if (!TryGetClosestSideFaceCandidate(
                    wall.Document,
                    wall,
                    lineContext,
                    startPoint,
                    endPoint,
                    options.ProjectionTolerance,
                    out WallSideFaceCandidate sideFaceCandidate,
                    out QuickDimensionRejectedReason rejectedReason,
                    out string sideFaceFailureReason))
                {
                    diagnostics.Add(new QuickDimensionDiagnostic(
                        QuickDimensionDiagnosticSeverity.Warning,
                        rejectedReason,
                        $"Skipped Wall '{displayName}' because no side-face reference passed the boundary rules: {sideFaceFailureReason}",
                        wall.Id,
                        QuickDimensionSourceType.Wall,
                        displayName));
                    return;
                }

                string candidateDisplayName = $"{displayName} [{sideFaceCandidate.ShellLayerType}]";
                string typeName = wall.WallType?.Name ?? string.Empty;

                candidates.Add(new QuickDimensionCandidate(
                    wall.Id,
                    QuickDimensionSourceType.Wall,
                    candidateDisplayName,
                    sideFaceCandidate.Reference,
                    QuickDimensionReferenceStrategy.WallSideFace,
                    sideFaceCandidate.HitPoint,
                    sideFaceCandidate.ParameterOnDimensionLine,
                    typeName: typeName));
            }
            catch (Exception ex)
            {
                diagnostics.Add(new QuickDimensionDiagnostic(
                    QuickDimensionDiagnosticSeverity.Warning,
                    QuickDimensionRejectedReason.CollectorException,
                    $"Skipped Wall '{displayName}' because the collector caught an API exception: {ex.Message}",
                    wall.Id,
                    QuickDimensionSourceType.Wall,
                    displayName));
            }
        }

        private static bool TryGetClosestSideFaceCandidate(
            Document doc,
            Wall wall,
            QuickDimensionLineContext lineContext,
            XYZ wallStartPoint,
            XYZ wallEndPoint,
            double tolerance,
            out WallSideFaceCandidate selectedCandidate,
            out QuickDimensionRejectedReason rejectedReason,
            out string failureReason)
        {
            selectedCandidate = null!;
            rejectedReason = QuickDimensionRejectedReason.MissingReference;
            failureReason = string.Empty;

            var faceCandidates = new List<WallSideFaceCandidate>();
            int sideFaceReferenceCount = 0;
            int usableSideFaceCount = 0;

            CollectSideFaceCandidates(
                doc,
                wall,
                ShellLayerType.Exterior,
                lineContext,
                wallStartPoint,
                wallEndPoint,
                tolerance,
                ref sideFaceReferenceCount,
                ref usableSideFaceCount,
                faceCandidates);

            CollectSideFaceCandidates(
                doc,
                wall,
                ShellLayerType.Interior,
                lineContext,
                wallStartPoint,
                wallEndPoint,
                tolerance,
                ref sideFaceReferenceCount,
                ref usableSideFaceCount,
                faceCandidates);

            if (faceCandidates.Count == 0)
            {
                if (sideFaceReferenceCount == 0)
                {
                    rejectedReason = QuickDimensionRejectedReason.MissingReference;
                    failureReason = "HostObjectUtils.GetSideFaces() returned no side-face references for Exterior or Interior shell layers.";
                    return false;
                }

                if (usableSideFaceCount == 0)
                {
                    rejectedReason = QuickDimensionRejectedReason.InvalidGeometry;
                    failureReason = "side-face references did not resolve to supported vertical planar faces.";
                    return false;
                }

                rejectedReason = QuickDimensionRejectedReason.OutsidePickedSpan;
                failureReason = "resolved side-face boundary segments do not intersect the picked dimension span.";
                return false;
            }

            selectedCandidate = faceCandidates
                .OrderBy(candidate => candidate.DistanceToDimensionLine)
                .ThenBy(candidate => candidate.ShellLayerType == ShellLayerType.Exterior ? 0 : 1)
                .First();

            return true;
        }

        private static void CollectSideFaceCandidates(
            Document doc,
            Wall wall,
            ShellLayerType shellLayerType,
            QuickDimensionLineContext lineContext,
            XYZ wallStartPoint,
            XYZ wallEndPoint,
            double tolerance,
            ref int sideFaceReferenceCount,
            ref int usableSideFaceCount,
            List<WallSideFaceCandidate> faceCandidates)
        {
            IList<Reference> sideFaceReferences;
            try
            {
                sideFaceReferences = HostObjectUtils.GetSideFaces(wall, shellLayerType);
            }
            catch
            {
                return;
            }

            if (sideFaceReferences == null || sideFaceReferences.Count == 0)
            {
                return;
            }

            sideFaceReferenceCount += sideFaceReferences.Count;

            foreach (Reference sideFaceReference in sideFaceReferences)
            {
                if (sideFaceReference == null)
                {
                    continue;
                }

                if (!TryGetPlanarFace(doc, sideFaceReference, out PlanarFace planarFace))
                {
                    continue;
                }

                XYZ faceNormal = planarFace.FaceNormal;
                if (!IsSupportedVerticalSideFaceNormal(faceNormal, tolerance))
                {
                    continue;
                }

                if (!TryGetFaceCentroid(planarFace, out XYZ faceCentroid))
                {
                    continue;
                }

                if (!TryBuildPlanarSideFaceSegment(
                    wallStartPoint,
                    wallEndPoint,
                    faceNormal,
                    faceCentroid,
                    tolerance,
                    out XYZ faceSegmentStart,
                    out XYZ faceSegmentEnd))
                {
                    continue;
                }

                usableSideFaceCount++;

                if (!QuickDimensionGeometryService.TryIntersectSegmentWithDimensionLine2D(
                    lineContext,
                    faceSegmentStart,
                    faceSegmentEnd,
                    tolerance,
                    out XYZ hitPoint,
                    out double parameterOnDimensionLine))
                {
                    continue;
                }

                double distanceToDimensionLine = QuickDimensionGeometryService.DistanceToDimensionLine2D(lineContext, faceCentroid);
                faceCandidates.Add(new WallSideFaceCandidate(
                    shellLayerType,
                    sideFaceReference,
                    hitPoint,
                    parameterOnDimensionLine,
                    distanceToDimensionLine));
            }
        }

        private static bool TryGetPlanarFace(Document doc, Reference faceReference, out PlanarFace planarFace)
        {
            planarFace = null!;

            try
            {
                Element element = doc.GetElement(faceReference);
                GeometryObject? geometryObject = element?.GetGeometryObjectFromReference(faceReference);
                if (geometryObject is not PlanarFace face)
                {
                    return false;
                }

                planarFace = face;
                return true;
            }
            catch
            {
                return false;
            }
        }

        private static bool TryGetFaceCentroid(PlanarFace face, out XYZ centroid)
        {
            centroid = null!;

            try
            {
                BoundingBoxUV boundingBox = face.GetBoundingBox();
                UV midpoint = (boundingBox.Min + boundingBox.Max) * 0.5;
                XYZ evaluatedPoint = face.Evaluate(midpoint);
                if (!QuickDimensionGeometryService.IsFinite(evaluatedPoint))
                {
                    return false;
                }

                centroid = evaluatedPoint;
                return true;
            }
            catch
            {
                return false;
            }
        }

        private static bool TryBuildPlanarSideFaceSegment(
            XYZ wallStartPoint,
            XYZ wallEndPoint,
            XYZ faceNormal,
            XYZ facePoint,
            double tolerance,
            out XYZ faceSegmentStart,
            out XYZ faceSegmentEnd)
        {
            faceSegmentStart = null!;
            faceSegmentEnd = null!;

            if (!QuickDimensionGeometryService.IsFinite(wallStartPoint) ||
                !QuickDimensionGeometryService.IsFinite(wallEndPoint) ||
                !QuickDimensionGeometryService.IsFinite(faceNormal) ||
                !QuickDimensionGeometryService.IsFinite(facePoint))
            {
                return false;
            }

            double normalLength = Math.Sqrt((faceNormal.X * faceNormal.X) + (faceNormal.Y * faceNormal.Y));
            if (normalLength <= tolerance)
            {
                return false;
            }

            XYZ planarNormal = new XYZ(faceNormal.X / normalLength, faceNormal.Y / normalLength, 0.0);
            double signedOffset = (facePoint - wallStartPoint).DotProduct(planarNormal);

            faceSegmentStart = wallStartPoint + (planarNormal * signedOffset);
            faceSegmentEnd = wallEndPoint + (planarNormal * signedOffset);
            return QuickDimensionGeometryService.IsFinite(faceSegmentStart) && QuickDimensionGeometryService.IsFinite(faceSegmentEnd);
        }

        private static bool IsSupportedVerticalSideFaceNormal(XYZ faceNormal, double tolerance)
        {
            if (!QuickDimensionGeometryService.IsFinite(faceNormal))
            {
                return false;
            }

            if (Math.Abs(faceNormal.Z) > VerticalFaceNormalZTolerance)
            {
                return false;
            }

            double planarLength = Math.Sqrt((faceNormal.X * faceNormal.X) + (faceNormal.Y * faceNormal.Y));
            return planarLength > tolerance;
        }

        private static QuickDimensionReadOnlyResult BuildResult(
            QuickDimensionLineContext lineContext,
            IReadOnlyList<QuickDimensionCandidate> rawCandidates,
            List<QuickDimensionDiagnostic> diagnostics,
            int collectedCount,
            QuickDimensionOptions options)
        {
            IReadOnlyList<QuickDimensionCandidate> candidates = QuickDimensionGeometryService
                .DeduplicateCandidates(rawCandidates, options.DuplicateTolerance);

            AddDuplicateDiagnostics(rawCandidates, candidates, diagnostics);
            AddAcceptedDiagnostics(candidates, diagnostics);

            var summaries = new[]
            {
                new QuickDimensionSourceSummary(
                    QuickDimensionSourceType.Wall,
                    collectedCount,
                    candidates.Count,
                    Math.Max(0, collectedCount - candidates.Count))
            };

            return new QuickDimensionReadOnlyResult(
                lineContext,
                candidates,
                diagnostics,
                summaries,
                options);
        }

        private static void AddDuplicateDiagnostics(
            IReadOnlyList<QuickDimensionCandidate> rawCandidates,
            IReadOnlyList<QuickDimensionCandidate> deduplicatedCandidates,
            List<QuickDimensionDiagnostic> diagnostics)
        {
            if (rawCandidates.Count == deduplicatedCandidates.Count)
            {
                return;
            }

            foreach (QuickDimensionCandidate rawCandidate in rawCandidates)
            {
                bool kept = deduplicatedCandidates.Any(candidate => ReferenceEquals(candidate, rawCandidate));
                if (kept)
                {
                    continue;
                }

                diagnostics.Add(new QuickDimensionDiagnostic(
                    QuickDimensionDiagnosticSeverity.Warning,
                    QuickDimensionRejectedReason.DuplicateCandidate,
                    $"Removed duplicate Wall candidate '{rawCandidate.DisplayName}' during source-aware deduplication.",
                    rawCandidate.ElementId,
                    QuickDimensionSourceType.Wall,
                    rawCandidate.DisplayName));
            }
        }

        private static void AddAcceptedDiagnostics(
            IEnumerable<QuickDimensionCandidate> candidates,
            List<QuickDimensionDiagnostic> diagnostics)
        {
            foreach (QuickDimensionCandidate candidate in candidates)
            {
                diagnostics.Add(new QuickDimensionDiagnostic(
                    QuickDimensionDiagnosticSeverity.Info,
                    QuickDimensionRejectedReason.None,
                    $"Accepted Wall boundary candidate '{candidate.DisplayName}' at parameter {candidate.ParameterOnDimensionLine:0.####}.",
                    candidate.ElementId,
                    QuickDimensionSourceType.Wall,
                    candidate.DisplayName));
            }
        }

        private static bool IsSupportedPlanView(RevitView view)
        {
            if (view.IsTemplate)
            {
                return false;
            }

            return view.ViewType == ViewType.FloorPlan
                || view.ViewType == ViewType.CeilingPlan
                || view.ViewType == ViewType.EngineeringPlan
                || view.ViewType == ViewType.AreaPlan;
        }

        private static string BuildWallDisplayName(Wall wall)
        {
            string typeName = wall.WallType?.Name ?? string.Empty;
            return string.IsNullOrWhiteSpace(typeName)
                ? $"Wall {wall.Id.Value}"
                : $"Wall: {typeName}";
        }

        private sealed class WallSideFaceCandidate
        {
            public WallSideFaceCandidate(
                ShellLayerType shellLayerType,
                Reference reference,
                XYZ hitPoint,
                double parameterOnDimensionLine,
                double distanceToDimensionLine)
            {
                ShellLayerType = shellLayerType;
                Reference = reference ?? throw new ArgumentNullException(nameof(reference));
                HitPoint = hitPoint ?? throw new ArgumentNullException(nameof(hitPoint));
                ParameterOnDimensionLine = parameterOnDimensionLine;
                DistanceToDimensionLine = distanceToDimensionLine;
            }

            public ShellLayerType ShellLayerType { get; }
            public Reference Reference { get; }
            public XYZ HitPoint { get; }
            public double ParameterOnDimensionLine { get; }
            public double DistanceToDimensionLine { get; }
        }
    }
}
