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
    /// Production read-only Door/Window opening candidate collector for Quick Dimension.
    /// This service performs no transactions and never creates dimensions.
    /// </summary>
    public static class QuickDimensionDoorWindowCandidateCollector
    {
        private const double OpeningGeometryBoundingBoxPadding = 0.5;
        private const double VerticalEdgeDirectionZMinimum = 0.9;
        private const double OpeningSpanMismatchMinimumMm = 10.0;
        private const double OpeningSpanMismatchRatio = 0.02;

        /// <summary>
        /// Collects wall-hosted Door and Window opening-edge candidates visible in the supplied plan view.
        /// FamilyInstance Left/Right references are preferred; host-wall opening edge references are used as fallback.
        /// </summary>
        public static QuickDimensionReadOnlyResult CollectDoorWindowCandidates(
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
            int collectedDoorCount = 0;
            int collectedWindowCount = 0;

            bool includeDoors = effectiveOptions.IncludesSource(QuickDimensionSourceType.Door);
            bool includeWindows = effectiveOptions.IncludesSource(QuickDimensionSourceType.Window);

            if (!includeDoors)
            {
                diagnostics.Add(new QuickDimensionDiagnostic(
                    QuickDimensionDiagnosticSeverity.Info,
                    QuickDimensionRejectedReason.None,
                    "Door opening collection is disabled by Quick Dimension options.",
                    sourceType: QuickDimensionSourceType.Door));
            }

            if (!includeWindows)
            {
                diagnostics.Add(new QuickDimensionDiagnostic(
                    QuickDimensionDiagnosticSeverity.Info,
                    QuickDimensionRejectedReason.None,
                    "Window opening collection is disabled by Quick Dimension options.",
                    sourceType: QuickDimensionSourceType.Window));
            }

            if (!includeDoors && !includeWindows)
            {
                return BuildResult(
                    lineContext,
                    rawCandidates,
                    diagnostics,
                    collectedDoorCount,
                    collectedWindowCount,
                    effectiveOptions);
            }

            if (!IsSupportedPlanView(view))
            {
                diagnostics.Add(new QuickDimensionDiagnostic(
                    QuickDimensionDiagnosticSeverity.Error,
                    QuickDimensionRejectedReason.UnsupportedView,
                    $"Quick Dimension Door/Window collection supports active plan views only. Current view type: {view.ViewType}.",
                    view.Id,
                    displayName: view.Name));

                return BuildResult(
                    lineContext,
                    rawCandidates,
                    diagnostics,
                    collectedDoorCount,
                    collectedWindowCount,
                    effectiveOptions);
            }

            if (includeDoors)
            {
                collectedDoorCount = CollectOpeningSourceCandidates(
                    doc,
                    view,
                    BuiltInCategory.OST_Doors,
                    QuickDimensionSourceType.Door,
                    lineContext,
                    effectiveOptions,
                    rawCandidates,
                    diagnostics);
            }

            if (includeWindows)
            {
                collectedWindowCount = CollectOpeningSourceCandidates(
                    doc,
                    view,
                    BuiltInCategory.OST_Windows,
                    QuickDimensionSourceType.Window,
                    lineContext,
                    effectiveOptions,
                    rawCandidates,
                    diagnostics);
            }

            return BuildResult(
                lineContext,
                rawCandidates,
                diagnostics,
                collectedDoorCount,
                collectedWindowCount,
                effectiveOptions);
        }

        private static int CollectOpeningSourceCandidates(
            Document doc,
            RevitView view,
            BuiltInCategory category,
            QuickDimensionSourceType sourceType,
            QuickDimensionLineContext lineContext,
            QuickDimensionOptions options,
            List<QuickDimensionCandidate> candidates,
            List<QuickDimensionDiagnostic> diagnostics)
        {
            List<FamilyInstance> instances = new FilteredElementCollector(doc, view.Id)
                .OfCategory(category)
                .OfClass(typeof(FamilyInstance))
                .Cast<FamilyInstance>()
                .ToList();

            foreach (FamilyInstance instance in instances)
            {
                if (instance?.IsValidObject != true)
                {
                    diagnostics.Add(new QuickDimensionDiagnostic(
                        QuickDimensionDiagnosticSeverity.Warning,
                        QuickDimensionRejectedReason.InvalidGeometry,
                        $"Skipped an invalid {sourceType} object returned by the active-view collector.",
                        sourceType: sourceType));
                    continue;
                }

                TryCollectOpeningCandidates(instance, sourceType, lineContext, options, candidates, diagnostics);
            }

            return instances.Count;
        }

        /// <summary>
        /// Wall-axis projection model (ADR-2026-06-11): collects Door/Window opening jamb candidates whose host
        /// is the selected wall, projecting BOTH left and right jambs onto the wall axis. The participation test is
        /// projection within the axis span (not 2D intersection), and the parallel guard is intentionally NOT applied
        /// because the host wall is, by definition, the axis itself.
        /// </summary>
        public static QuickDimensionReadOnlyResult CollectOpeningsAlongWallAxis(
            Document doc,
            RevitView view,
            QuickDimensionLineContext lineContext,
            ElementId selectedWallId,
            QuickDimensionOptions? options = null)
        {
            if (doc == null) throw new ArgumentNullException(nameof(doc));
            if (view == null) throw new ArgumentNullException(nameof(view));
            if (lineContext == null) throw new ArgumentNullException(nameof(lineContext));
            if (selectedWallId == null) throw new ArgumentNullException(nameof(selectedWallId));

            QuickDimensionOptions effectiveOptions = options ?? QuickDimensionOptions.Default;
            var diagnostics = new List<QuickDimensionDiagnostic>();
            var rawCandidates = new List<QuickDimensionCandidate>();
            int collectedDoorCount = 0;
            int collectedWindowCount = 0;

            bool includeDoors = effectiveOptions.IncludesSource(QuickDimensionSourceType.Door);
            bool includeWindows = effectiveOptions.IncludesSource(QuickDimensionSourceType.Window);

            if (includeDoors)
            {
                collectedDoorCount = CollectWallAxisOpeningSource(
                    doc, view, BuiltInCategory.OST_Doors, QuickDimensionSourceType.Door,
                    lineContext, selectedWallId, effectiveOptions, rawCandidates, diagnostics);
            }

            if (includeWindows)
            {
                collectedWindowCount = CollectWallAxisOpeningSource(
                    doc, view, BuiltInCategory.OST_Windows, QuickDimensionSourceType.Window,
                    lineContext, selectedWallId, effectiveOptions, rawCandidates, diagnostics);
            }

            return BuildResult(
                lineContext, rawCandidates, diagnostics, collectedDoorCount, collectedWindowCount, effectiveOptions);
        }

        private static int CollectWallAxisOpeningSource(
            Document doc,
            RevitView view,
            BuiltInCategory category,
            QuickDimensionSourceType sourceType,
            QuickDimensionLineContext lineContext,
            ElementId selectedWallId,
            QuickDimensionOptions options,
            List<QuickDimensionCandidate> candidates,
            List<QuickDimensionDiagnostic> diagnostics)
        {
            List<FamilyInstance> instances = new FilteredElementCollector(doc, view.Id)
                .OfCategory(category)
                .OfClass(typeof(FamilyInstance))
                .Cast<FamilyInstance>()
                .Where(instance => instance?.Host is Wall hostWall && hostWall.Id == selectedWallId)
                .ToList();

            foreach (FamilyInstance instance in instances)
            {
                if (instance?.IsValidObject != true)
                {
                    continue;
                }

                TryProjectOpeningOntoWallAxis(instance, sourceType, lineContext, options, candidates, diagnostics);
            }

            return instances.Count;
        }

        private static void TryProjectOpeningOntoWallAxis(
            FamilyInstance instance,
            QuickDimensionSourceType sourceType,
            QuickDimensionLineContext lineContext,
            QuickDimensionOptions options,
            List<QuickDimensionCandidate> candidates,
            List<QuickDimensionDiagnostic> diagnostics)
        {
            string familyName = GetFamilyName(instance);
            string typeName = GetTypeName(instance);
            string displayName = BuildOpeningDisplayName(sourceType, familyName, typeName, instance.Id);

            try
            {
                Element? hostElement = instance.Host;
                if (hostElement is not Wall hostWall)
                {
                    return;
                }

                // In the wall-axis model the host wall IS the axis, so the axis direction is the projection direction.
                XYZ wallDirection = lineContext.Direction;

                OpeningReferencePair familyReferences = TryGetFamilyInstanceReferences(instance);
                OpeningReferencePair fallbackReferences = options.EnableHostWallOpeningFallback
                    ? TryGetHostWallOpeningReferences(hostWall, instance, wallDirection, options.DuplicateTolerance)
                    : OpeningReferencePair.Empty;

                var referenceInfos = new List<OpeningReferenceInfo>();
                AddOpeningGeometryValidationDiagnostic(
                    diagnostics, fallbackReferences.Validation, instance.Id, sourceType, displayName);

                bool useFallbackReferences = fallbackReferences.Validation?.UseFallbackReferences == true;
                AddOpeningReferenceInfo(
                    referenceInfos, familyReferences.LeftReference, familyReferences.LeftPoint,
                    fallbackReferences.LeftReference, fallbackReferences.LeftPoint, "Left",
                    options.EnableHostWallOpeningFallback, useFallbackReferences);
                AddOpeningReferenceInfo(
                    referenceInfos, familyReferences.RightReference, familyReferences.RightPoint,
                    fallbackReferences.RightReference, fallbackReferences.RightPoint, "Right",
                    options.EnableHostWallOpeningFallback, useFallbackReferences);

                if (referenceInfos.Count == 0)
                {
                    diagnostics.Add(new QuickDimensionDiagnostic(
                        QuickDimensionDiagnosticSeverity.Warning,
                        QuickDimensionRejectedReason.MissingReference,
                        $"Skipped {displayName} because no valid opening-edge reference was available. Host Wall: {hostWall.Id.Value}.",
                        instance.Id, sourceType, displayName));
                    return;
                }

                int acceptedReferenceCount = 0;
                foreach (OpeningReferenceInfo referenceInfo in referenceInfos)
                {
                    if (!QuickDimensionGeometryService.TryProjectPointToPickedSpan(
                        lineContext, referenceInfo.ReferencePoint, options.ProjectionTolerance,
                        out XYZ projectedPoint, out double parameterOnDimensionLine))
                    {
                        diagnostics.Add(new QuickDimensionDiagnostic(
                            QuickDimensionDiagnosticSeverity.Warning,
                            QuickDimensionRejectedReason.OutsidePickedSpan,
                            $"Skipped {displayName} [{referenceInfo.Label}] because its projected jamb falls outside the wall axis span. Host Wall: {hostWall.Id.Value}.",
                            instance.Id, sourceType, $"{displayName} [{referenceInfo.Label}]"));
                        continue;
                    }

                    // [PHASE2_INSTRUMENTATION] Log reference identity + projection point + provenance
                    string referenceIdentity = referenceInfo.Reference?.ElementId.Value.ToString() ?? "null";
                    string provenance = DeriveProvenance(referenceInfo);
                    diagnostics.Add(new QuickDimensionDiagnostic(
                        QuickDimensionDiagnosticSeverity.Info,
                        QuickDimensionRejectedReason.None,
                        $"[PHASE2_INSTRUMENTATION] {displayName} [{referenceInfo.Label}]: " +
                        $"Reference={referenceIdentity}, " +
                        $"ReferenceOwner={referenceInfo.Reference?.ElementId.Value ?? instance.Id.Value}, " +
                        $"ProjectionPoint=({referenceInfo.ReferencePoint.X:F6},{referenceInfo.ReferencePoint.Y:F6},{referenceInfo.ReferencePoint.Z:F6}), " +
                        $"ProjectedStation={parameterOnDimensionLine:F6}, " +
                        $"PointProvenance={provenance}, " +
                        $"Strategy={referenceInfo.ReferenceStrategy}",
                        instance.Id, sourceType, $"{displayName} [{referenceInfo.Label}]"));

                    ElementId candidateElementId = ResolveCandidateElementId(instance.Id, referenceInfo);
                    candidates.Add(new QuickDimensionCandidate(
                        candidateElementId, sourceType, $"{displayName} [{referenceInfo.Label}]",
                        referenceInfo.Reference!, referenceInfo.ReferenceStrategy,
                        projectedPoint, parameterOnDimensionLine,
                        hostWall.Id, familyName, typeName));

                    acceptedReferenceCount++;
                }

                if (acceptedReferenceCount == 0)
                {
                    diagnostics.Add(new QuickDimensionDiagnostic(
                        QuickDimensionDiagnosticSeverity.Warning,
                        QuickDimensionRejectedReason.OutsidePickedSpan,
                        $"Skipped {displayName} because no projected opening-edge fell inside the wall axis span. Host Wall: {hostWall.Id.Value}.",
                        instance.Id, sourceType, displayName));
                }
            }
            catch (Exception ex)
            {
                diagnostics.Add(new QuickDimensionDiagnostic(
                    QuickDimensionDiagnosticSeverity.Warning,
                    QuickDimensionRejectedReason.CollectorException,
                    $"Skipped {displayName} because the collector caught an API exception: {ex.Message}",
                    instance.Id, sourceType, displayName));
            }
        }

        /// <summary>
        /// [PHASE2_INSTRUMENTATION] Derive provenance label for the projection point.
        /// </summary>
        private static string DeriveProvenance(OpeningReferenceInfo referenceInfo)
        {
            if (referenceInfo.ReferenceStrategy == QuickDimensionReferenceStrategy.HostWallOpeningGeometry)
                return "Reference-Owned (HostWallOpeningGeometry)";

            if (referenceInfo.ReferenceStrategy == QuickDimensionReferenceStrategy.FamilyInstanceLeftRight)
                return "Reference-Owned (FamilyInstanceLeftRight)";

            return "Unknown Strategy";
        }

        private static ElementId ResolveCandidateElementId(ElementId instanceId, OpeningReferenceInfo referenceInfo)
        {
            if (referenceInfo.ReferenceStrategy == QuickDimensionReferenceStrategy.HostWallOpeningGeometry)
            {
                ElementId? referenceOwnerId = referenceInfo.Reference?.ElementId;
                if (referenceOwnerId != null && referenceOwnerId != ElementId.InvalidElementId)
                {
                    return referenceOwnerId;
                }
            }

            return instanceId;
        }

        private static void TryCollectOpeningCandidates(
            FamilyInstance instance,
            QuickDimensionSourceType sourceType,
            QuickDimensionLineContext lineContext,
            QuickDimensionOptions options,
            List<QuickDimensionCandidate> candidates,
            List<QuickDimensionDiagnostic> diagnostics)
        {
            string familyName = GetFamilyName(instance);
            string typeName = GetTypeName(instance);
            string displayName = BuildOpeningDisplayName(sourceType, familyName, typeName, instance.Id);

            try
            {
                Element? hostElement = instance.Host;
                if (hostElement is not Wall hostWall)
                {
                    diagnostics.Add(new QuickDimensionDiagnostic(
                        QuickDimensionDiagnosticSeverity.Warning,
                        QuickDimensionRejectedReason.NonWallHostedOpening,
                        $"Skipped {displayName} because it is not hosted by a Wall.",
                        instance.Id,
                        sourceType,
                        displayName));
                    return;
                }

                if (hostWall.Location is not LocationCurve locationCurve || locationCurve.Curve == null)
                {
                    diagnostics.Add(new QuickDimensionDiagnostic(
                        QuickDimensionDiagnosticSeverity.Warning,
                        QuickDimensionRejectedReason.InvalidGeometry,
                        $"Skipped {displayName} because host Wall {hostWall.Id.Value} does not expose a valid LocationCurve.",
                        instance.Id,
                        sourceType,
                        displayName));
                    return;
                }

                Curve hostCurve = locationCurve.Curve;
                if (hostCurve is not Line)
                {
                    diagnostics.Add(new QuickDimensionDiagnostic(
                        QuickDimensionDiagnosticSeverity.Warning,
                        QuickDimensionRejectedReason.ArcWallUnsupported,
                        $"Skipped {displayName} because host Wall {hostWall.Id.Value} is arc or non-line, which is outside the Quick Dimension MVP scope.",
                        instance.Id,
                        sourceType,
                        displayName));
                    return;
                }

                if (!QuickDimensionGeometryService.TryGetStraightCurveEndpoints(hostCurve, out XYZ hostStartPoint, out XYZ hostEndPoint))
                {
                    diagnostics.Add(new QuickDimensionDiagnostic(
                        QuickDimensionDiagnosticSeverity.Warning,
                        QuickDimensionRejectedReason.InvalidGeometry,
                        $"Skipped {displayName} because host Wall {hostWall.Id.Value} curve endpoints are not valid finite points.",
                        instance.Id,
                        sourceType,
                        displayName));
                    return;
                }

                if (!QuickDimensionGeometryService.TryGetPlanarDirection(hostStartPoint, hostEndPoint, options.ProjectionTolerance, out XYZ hostWallDirection))
                {
                    diagnostics.Add(new QuickDimensionDiagnostic(
                        QuickDimensionDiagnosticSeverity.Warning,
                        QuickDimensionRejectedReason.InvalidGeometry,
                        $"Skipped {displayName} because host Wall {hostWall.Id.Value} does not define a valid plan-view direction.",
                        instance.Id,
                        sourceType,
                        displayName));
                    return;
                }

                if (QuickDimensionGeometryService.IsNearlyParallel(lineContext.Direction, hostWallDirection, options.ProjectionTolerance))
                {
                    diagnostics.Add(new QuickDimensionDiagnostic(
                        QuickDimensionDiagnosticSeverity.Warning,
                        QuickDimensionRejectedReason.ParallelToDimensionLine,
                        $"Skipped {displayName} because host Wall {hostWall.Id.Value} is parallel to the picked dimension line.",
                        instance.Id,
                        sourceType,
                        displayName));
                    return;
                }

                OpeningReferencePair familyReferences = TryGetFamilyInstanceReferences(instance);
                OpeningReferencePair fallbackReferences = options.EnableHostWallOpeningFallback
                    ? TryGetHostWallOpeningReferences(hostWall, instance, lineContext.Direction, options.DuplicateTolerance)
                    : OpeningReferencePair.Empty;

                var referenceInfos = new List<OpeningReferenceInfo>();
                AddOpeningGeometryValidationDiagnostic(
                    diagnostics, fallbackReferences.Validation, instance.Id, sourceType, displayName);

                bool useFallbackReferences = fallbackReferences.Validation?.UseFallbackReferences == true;
                AddOpeningReferenceInfo(
                    referenceInfos,
                    familyReferences.LeftReference,
                    familyReferences.LeftPoint,
                    fallbackReferences.LeftReference,
                    fallbackReferences.LeftPoint,
                    "Left",
                    options.EnableHostWallOpeningFallback,
                    useFallbackReferences);

                AddOpeningReferenceInfo(
                    referenceInfos,
                    familyReferences.RightReference,
                    familyReferences.RightPoint,
                    fallbackReferences.RightReference,
                    fallbackReferences.RightPoint,
                    "Right",
                    options.EnableHostWallOpeningFallback,
                    useFallbackReferences);

                if (referenceInfos.Count == 0)
                {
                    string fallbackState = options.EnableHostWallOpeningFallback
                        ? "FamilyInstance Left/Right references and host-wall opening fallback both failed."
                        : "FamilyInstance Left/Right references were missing and host-wall opening fallback is disabled.";

                    diagnostics.Add(new QuickDimensionDiagnostic(
                        QuickDimensionDiagnosticSeverity.Warning,
                        QuickDimensionRejectedReason.MissingReference,
                        $"Skipped {displayName} because no valid opening-edge reference was available. {fallbackState} Host Wall: {hostWall.Id.Value}.",
                        instance.Id,
                        sourceType,
                        displayName));
                    return;
                }

                int acceptedReferenceCount = 0;
                foreach (OpeningReferenceInfo referenceInfo in referenceInfos)
                {
                    if (!TryBuildReferenceSegment(
                        instance,
                        referenceInfo.ReferencePoint,
                        hostWallDirection,
                        options.ProjectionTolerance,
                        out XYZ segmentStart,
                        out XYZ segmentEnd))
                    {
                        diagnostics.Add(new QuickDimensionDiagnostic(
                            QuickDimensionDiagnosticSeverity.Warning,
                            QuickDimensionRejectedReason.InvalidGeometry,
                            $"Skipped {displayName} [{referenceInfo.Label}] because its opening-edge segment could not be reconstructed from bounding-box geometry. Host Wall: {hostWall.Id.Value}.",
                            instance.Id,
                            sourceType,
                            $"{displayName} [{referenceInfo.Label}]"));
                        continue;
                    }

                    if (!QuickDimensionGeometryService.TryIntersectSegmentWithDimensionLine2D(
                        lineContext,
                        segmentStart,
                        segmentEnd,
                        options.ProjectionTolerance,
                        out XYZ hitPoint,
                        out double parameterOnDimensionLine))
                    {
                        diagnostics.Add(new QuickDimensionDiagnostic(
                            QuickDimensionDiagnosticSeverity.Warning,
                            QuickDimensionRejectedReason.OutsidePickedSpan,
                            $"Skipped {displayName} [{referenceInfo.Label}] because its opening edge does not intersect the picked dimension span. Host Wall: {hostWall.Id.Value}.",
                            instance.Id,
                            sourceType,
                            $"{displayName} [{referenceInfo.Label}]"));
                        continue;
                    }

                    ElementId candidateElementId = ResolveCandidateElementId(instance.Id, referenceInfo);
                    candidates.Add(new QuickDimensionCandidate(
                        candidateElementId,
                        sourceType,
                        $"{displayName} [{referenceInfo.Label}]",
                        referenceInfo.Reference,
                        referenceInfo.ReferenceStrategy,
                        hitPoint,
                        parameterOnDimensionLine,
                        hostWall.Id,
                        familyName,
                        typeName));

                    acceptedReferenceCount++;
                }

                if (acceptedReferenceCount == 0)
                {
                    diagnostics.Add(new QuickDimensionDiagnostic(
                        QuickDimensionDiagnosticSeverity.Warning,
                        QuickDimensionRejectedReason.OutsidePickedSpan,
                        $"Skipped {displayName} because no available opening-edge reference intersected the picked dimension span. Host Wall: {hostWall.Id.Value}.",
                        instance.Id,
                        sourceType,
                        displayName));
                }
            }
            catch (Exception ex)
            {
                diagnostics.Add(new QuickDimensionDiagnostic(
                    QuickDimensionDiagnosticSeverity.Warning,
                    QuickDimensionRejectedReason.CollectorException,
                    $"Skipped {displayName} because the collector caught an API exception: {ex.Message}",
                    instance.Id,
                    sourceType,
                    displayName));
            }
        }

        private static void AddOpeningReferenceInfo(
            List<OpeningReferenceInfo> referenceInfos,
            Reference? familyReference,
            XYZ? familyPoint,
            Reference? fallbackReference,
            XYZ? fallbackPoint,
            string label,
            bool fallbackEnabled,
            bool useFallbackReference)
        {
            if (fallbackEnabled && useFallbackReference && fallbackReference != null && fallbackPoint != null)
            {
                referenceInfos.Add(new OpeningReferenceInfo(
                    fallbackReference,
                    fallbackPoint,
                    label,
                    QuickDimensionReferenceStrategy.HostWallOpeningGeometry));
                return;
            }

            if (familyReference != null && familyPoint != null)
            {
                referenceInfos.Add(new OpeningReferenceInfo(
                    familyReference,
                    familyPoint,
                    label,
                    QuickDimensionReferenceStrategy.FamilyInstanceLeftRight));
                return;
            }

            if (fallbackEnabled && fallbackReference != null && fallbackPoint != null)
            {
                referenceInfos.Add(new OpeningReferenceInfo(
                    fallbackReference,
                    fallbackPoint,
                    label,
                    QuickDimensionReferenceStrategy.HostWallOpeningGeometry));
            }
        }

        private static void AddOpeningGeometryValidationDiagnostic(
            List<QuickDimensionDiagnostic> diagnostics,
            OpeningGeometryValidation? validation,
            ElementId instanceId,
            QuickDimensionSourceType sourceType,
            string displayName)
        {
            if (validation == null)
            {
                return;
            }

            string expectedSpan = FormatOptionalLengthMm(validation.ExpectedSpan);
            string boundingBoxSpan = FormatOptionalLengthMm(validation.BoundingBoxSpan);
            string rawSpan = FormatOptionalLengthMm(validation.RawSpan);
            string selectedSpan = FormatOptionalLengthMm(validation.SelectedSpan);
            string typeWidth = FormatOptionalLengthMm(validation.TypeWidth);
            string selectedStrategy = validation.UseFallbackReferences
                ? QuickDimensionReferenceStrategy.HostWallOpeningGeometry.ToString()
                : QuickDimensionReferenceStrategy.FamilyInstanceLeftRight.ToString();

            QuickDimensionDiagnosticSeverity severity = validation.SelectedPairValidated
                ? QuickDimensionDiagnosticSeverity.Info
                : QuickDimensionDiagnosticSeverity.Warning;

            diagnostics.Add(new QuickDimensionDiagnostic(
                severity,
                QuickDimensionRejectedReason.None,
                $"Opening geometry check for {displayName}: expected span {expectedSpan} (type width {typeWidth}, instance bbox {boundingBoxSpan}), raw padded-bbox extrema {rawSpan}, selected jamb pair {selectedSpan}; selected strategy: {selectedStrategy}.",
                instanceId,
                sourceType,
                displayName));
        }

        private static string FormatOptionalLengthMm(double? value)
        {
            if (!value.HasValue)
            {
                return "n/a";
            }

            double millimeters = UnitUtils.ConvertFromInternalUnits(value.Value, UnitTypeId.Millimeters);
            return $"{millimeters:0.##} mm";
        }

        private static OpeningReferencePair TryGetFamilyInstanceReferences(FamilyInstance instance)
        {
            Reference? leftReference = null;
            XYZ? leftPoint = null;
            Reference? rightReference = null;
            XYZ? rightPoint = null;
            XYZ? openingCenterPoint = GetOpeningCenterPoint(instance);

            try
            {
                IList<Reference>? leftReferences = instance.GetReferences(FamilyInstanceReferenceType.Left);
                if (leftReferences != null && leftReferences.Count > 0)
                {
                    leftReference = leftReferences[0];
                    if (TryResolveReferenceOwnedPoint(instance, leftReference, openingCenterPoint, out XYZ resolvedLeftPoint))
                    {
                        leftPoint = resolvedLeftPoint;
                    }
                }
            }
            catch
            {
                leftReference = null;
                leftPoint = null;
            }

            try
            {
                IList<Reference>? rightReferences = instance.GetReferences(FamilyInstanceReferenceType.Right);
                if (rightReferences != null && rightReferences.Count > 0)
                {
                    rightReference = rightReferences[0];
                    if (TryResolveReferenceOwnedPoint(instance, rightReference, openingCenterPoint, out XYZ resolvedRightPoint))
                    {
                        rightPoint = resolvedRightPoint;
                    }
                }
            }
            catch
            {
                rightReference = null;
                rightPoint = null;
            }

            return new OpeningReferencePair(leftReference, rightReference, leftPoint, rightPoint);
        }

        private static bool TryResolveReferenceOwnedPoint(
            FamilyInstance instance,
            Reference reference,
            XYZ? openingCenterPoint,
            out XYZ referencePoint)
        {
            referencePoint = null!;

            try
            {
                GeometryObject? geometryObject = instance.GetGeometryObjectFromReference(reference);
                if (geometryObject is Face face)
                {
                    if (openingCenterPoint == null || !QuickDimensionGeometryService.IsFinite(openingCenterPoint))
                    {
                        return false;
                    }

                    IntersectionResult? projection = face.Project(openingCenterPoint);
                    XYZ? projectedPoint = projection?.XYZPoint;
                    if (projectedPoint == null || !QuickDimensionGeometryService.IsFinite(projectedPoint))
                    {
                        return false;
                    }

                    referencePoint = projectedPoint;
                    return true;
                }

                if (geometryObject is Edge edge)
                {
                    Curve edgeCurve = edge.AsCurve();
                    if (edgeCurve is not Line edgeLine)
                    {
                        return false;
                    }

                    XYZ midpoint = (edgeLine.GetEndPoint(0) + edgeLine.GetEndPoint(1)) * 0.5;
                    if (!QuickDimensionGeometryService.IsFinite(midpoint))
                    {
                        return false;
                    }

                    referencePoint = midpoint;
                    return true;
                }
            }
            catch
            {
                // Keep the collector read-only and tolerant of non-stationable named references.
            }

            return false;
        }

        private static OpeningReferencePair TryGetHostWallOpeningReferences(
            Wall hostWall,
            FamilyInstance instance,
            XYZ dimensionDirection,
            double duplicateTolerance)
        {
            try
            {
                BoundingBoxXYZ? instanceBoundingBox = instance.get_BoundingBox(null);
                if (instanceBoundingBox == null ||
                    !TryGetBoundingBoxProjectionRange(
                        instanceBoundingBox,
                        dimensionDirection,
                        out double boundingBoxLeftPosition,
                        out double boundingBoxRightPosition))
                {
                    return OpeningReferencePair.Empty;
                }

                double boundingBoxSpan = boundingBoxRightPosition - boundingBoxLeftPosition;
                double? typeWidth = TryGetOpeningTypeWidth(instance);
                double expectedSpan = typeWidth ?? boundingBoxSpan;
                if (expectedSpan <= duplicateTolerance)
                {
                    return OpeningReferencePair.Empty;
                }

                XYZ? openingCenterPoint = GetOpeningCenterPoint(instance);
                double expectedCenterPosition = openingCenterPoint != null
                    ? openingCenterPoint.DotProduct(dimensionDirection)
                    : (boundingBoxLeftPosition + boundingBoxRightPosition) * 0.5;
                double expectedLeftPosition = expectedCenterPosition - (expectedSpan * 0.5);
                double expectedRightPosition = expectedCenterPosition + (expectedSpan * 0.5);

                Options geometryOptions = new Options
                {
                    ComputeReferences = true,
                    IncludeNonVisibleObjects = false
                };

                GeometryElement? wallGeometry = hostWall.get_Geometry(geometryOptions);
                if (wallGeometry == null)
                {
                    return OpeningReferencePair.Empty;
                }

                var edgeCandidates = new List<OpeningEdgeCandidate>();
                foreach (GeometryObject geometryObject in wallGeometry)
                {
                    if (geometryObject is Solid solid)
                    {
                        CollectOpeningEdgeCandidates(solid, instanceBoundingBox, dimensionDirection, edgeCandidates);
                    }
                }

                if (edgeCandidates.Count < 2)
                {
                    return new OpeningReferencePair(
                        null,
                        null,
                        null,
                        null,
                        new OpeningGeometryValidation(
                            expectedSpan,
                            boundingBoxSpan,
                            rawSpan: null,
                            selectedSpan: null,
                            typeWidth,
                            selectedPairValidated: false,
                            useFallbackReferences: false));
                }

                List<OpeningEdgeCandidate> orderedEdges = edgeCandidates
                    .OrderBy(candidate => candidate.PositionOnDimensionDirection)
                    .ToList();

                double rawSpan = orderedEdges.Last().PositionOnDimensionDirection
                    - orderedEdges.First().PositionOnDimensionDirection;

                if (!TrySelectNearestOpeningEdgePair(
                    orderedEdges,
                    expectedLeftPosition,
                    expectedRightPosition,
                    duplicateTolerance,
                    out OpeningEdgeCandidate selectedLeftEdge,
                    out OpeningEdgeCandidate selectedRightEdge))
                {
                    return new OpeningReferencePair(
                        null,
                        null,
                        null,
                        null,
                        new OpeningGeometryValidation(
                            expectedSpan,
                            boundingBoxSpan,
                            rawSpan,
                            selectedSpan: null,
                            typeWidth,
                            selectedPairValidated: false,
                            useFallbackReferences: false));
                }

                double selectedSpan = selectedRightEdge.PositionOnDimensionDirection
                    - selectedLeftEdge.PositionOnDimensionDirection;
                double mismatchTolerance = GetOpeningSpanMismatchTolerance(expectedSpan);
                bool selectedPairValidated =
                    Math.Abs(selectedLeftEdge.PositionOnDimensionDirection - expectedLeftPosition) <= mismatchTolerance &&
                    Math.Abs(selectedRightEdge.PositionOnDimensionDirection - expectedRightPosition) <= mismatchTolerance &&
                    Math.Abs(selectedSpan - expectedSpan) <= mismatchTolerance;

                bool rawExtremaMismatch = Math.Abs(rawSpan - expectedSpan) > mismatchTolerance;
                bool useFallbackReferences = selectedPairValidated && rawExtremaMismatch;

                return new OpeningReferencePair(
                    selectedPairValidated ? selectedLeftEdge.Reference : null,
                    selectedPairValidated ? selectedRightEdge.Reference : null,
                    selectedPairValidated ? selectedLeftEdge.Midpoint : null,
                    selectedPairValidated ? selectedRightEdge.Midpoint : null,
                    new OpeningGeometryValidation(
                        expectedSpan,
                        boundingBoxSpan,
                        rawSpan,
                        selectedSpan,
                        typeWidth,
                        selectedPairValidated,
                        useFallbackReferences));
            }
            catch
            {
                return OpeningReferencePair.Empty;
            }
        }

        private static bool TrySelectNearestOpeningEdgePair(
            IReadOnlyList<OpeningEdgeCandidate> orderedEdges,
            double expectedLeftPosition,
            double expectedRightPosition,
            double duplicateTolerance,
            out OpeningEdgeCandidate leftEdge,
            out OpeningEdgeCandidate rightEdge)
        {
            leftEdge = null!;
            rightEdge = null!;
            double bestScore = double.MaxValue;

            for (int leftIndex = 0; leftIndex < orderedEdges.Count - 1; leftIndex++)
            {
                OpeningEdgeCandidate leftCandidate = orderedEdges[leftIndex];
                for (int rightIndex = leftIndex + 1; rightIndex < orderedEdges.Count; rightIndex++)
                {
                    OpeningEdgeCandidate rightCandidate = orderedEdges[rightIndex];
                    double span = rightCandidate.PositionOnDimensionDirection
                        - leftCandidate.PositionOnDimensionDirection;
                    if (span <= duplicateTolerance)
                    {
                        continue;
                    }

                    double score =
                        Math.Abs(leftCandidate.PositionOnDimensionDirection - expectedLeftPosition) +
                        Math.Abs(rightCandidate.PositionOnDimensionDirection - expectedRightPosition);
                    if (score >= bestScore)
                    {
                        continue;
                    }

                    bestScore = score;
                    leftEdge = leftCandidate;
                    rightEdge = rightCandidate;
                }
            }

            return leftEdge != null && rightEdge != null;
        }

        private static double GetOpeningSpanMismatchTolerance(double expectedSpan)
        {
            double minimumTolerance = UnitUtils.ConvertToInternalUnits(
                OpeningSpanMismatchMinimumMm,
                UnitTypeId.Millimeters);
            return Math.Max(minimumTolerance, expectedSpan * OpeningSpanMismatchRatio);
        }

        private static double? TryGetOpeningTypeWidth(FamilyInstance instance)
        {
            Parameter? widthParameter = instance.Symbol?.get_Parameter(BuiltInParameter.FAMILY_WIDTH_PARAM)
                ?? instance.get_Parameter(BuiltInParameter.FAMILY_WIDTH_PARAM);
            if (widthParameter?.StorageType != StorageType.Double)
            {
                return null;
            }

            double width = widthParameter.AsDouble();
            return width > 0.0 && !double.IsNaN(width) && !double.IsInfinity(width)
                ? width
                : null;
        }

        private static void CollectOpeningEdgeCandidates(
            Solid solid,
            BoundingBoxXYZ instanceBoundingBox,
            XYZ dimensionDirection,
            List<OpeningEdgeCandidate> result)
        {
            XYZ bboxMin = instanceBoundingBox.Min - new XYZ(
                OpeningGeometryBoundingBoxPadding,
                OpeningGeometryBoundingBoxPadding,
                OpeningGeometryBoundingBoxPadding);

            XYZ bboxMax = instanceBoundingBox.Max + new XYZ(
                OpeningGeometryBoundingBoxPadding,
                OpeningGeometryBoundingBoxPadding,
                OpeningGeometryBoundingBoxPadding);

            foreach (Edge edge in solid.Edges)
            {
                try
                {
                    Reference? edgeReference = edge.Reference;
                    if (edgeReference == null)
                    {
                        continue;
                    }

                    Curve edgeCurve = edge.AsCurve();
                    if (edgeCurve is not Line edgeLine)
                    {
                        continue;
                    }

                    XYZ edgeDirection = edgeLine.Direction;
                    if (!QuickDimensionGeometryService.IsFinite(edgeDirection) || Math.Abs(edgeDirection.Z) < VerticalEdgeDirectionZMinimum)
                    {
                        continue;
                    }

                    XYZ midpoint = (edgeLine.GetEndPoint(0) + edgeLine.GetEndPoint(1)) * 0.5;
                    if (!IsPointInsideBoundingBox(midpoint, bboxMin, bboxMax))
                    {
                        continue;
                    }

                    double position = midpoint.DotProduct(dimensionDirection);
                    result.Add(new OpeningEdgeCandidate(edgeReference, midpoint, position));
                }
                catch
                {
                    // Skip problematic edges; one bad edge must not stop the read-only collector.
                }
            }
        }

        private static XYZ? GetOpeningCenterPoint(FamilyInstance instance)
        {
            if (instance.Location is LocationPoint locationPoint && QuickDimensionGeometryService.IsFinite(locationPoint.Point))
            {
                return locationPoint.Point;
            }

            BoundingBoxXYZ? boundingBox = instance.get_BoundingBox(null);
            if (boundingBox == null)
            {
                return null;
            }

            XYZ center = (boundingBox.Min + boundingBox.Max) * 0.5;
            return QuickDimensionGeometryService.IsFinite(center) ? center : null;
        }

        private static bool TryGetBoundingBoxProjectionRange(
            BoundingBoxXYZ boundingBox,
            XYZ direction,
            out double minimumProjection,
            out double maximumProjection)
        {
            minimumProjection = double.MaxValue;
            maximumProjection = double.MinValue;

            foreach (XYZ corner in GetBoundingBoxCorners(boundingBox))
            {
                if (!QuickDimensionGeometryService.IsFinite(corner))
                {
                    return false;
                }

                double projection = corner.DotProduct(direction);
                if (projection < minimumProjection)
                {
                    minimumProjection = projection;
                }

                if (projection > maximumProjection)
                {
                    maximumProjection = projection;
                }
            }

            return minimumProjection <= maximumProjection;
        }

        private static bool TryBuildReferenceSegment(
            FamilyInstance instance,
            XYZ referencePoint,
            XYZ hostWallDirection,
            double tolerance,
            out XYZ segmentStart,
            out XYZ segmentEnd)
        {
            segmentStart = null!;
            segmentEnd = null!;

            if (!QuickDimensionGeometryService.IsFinite(referencePoint) || !QuickDimensionGeometryService.IsFinite(hostWallDirection))
            {
                return false;
            }

            double halfSpan = GetOpeningHalfSpanAlongDirection(instance, hostWallDirection, tolerance);
            if (halfSpan <= tolerance)
            {
                return false;
            }

            segmentStart = referencePoint - (hostWallDirection * halfSpan);
            segmentEnd = referencePoint + (hostWallDirection * halfSpan);
            return QuickDimensionGeometryService.IsFinite(segmentStart) && QuickDimensionGeometryService.IsFinite(segmentEnd);
        }

        private static double GetOpeningHalfSpanAlongDirection(FamilyInstance instance, XYZ direction, double tolerance)
        {
            BoundingBoxXYZ? boundingBox = instance.get_BoundingBox(null);
            if (boundingBox == null || !TryGetBoundingBoxProjectionRange(boundingBox, direction, out double minimumProjection, out double maximumProjection))
            {
                return 0.0;
            }

            double span = maximumProjection - minimumProjection;
            if (span <= tolerance)
            {
                return 0.0;
            }

            return (span * 0.5) + tolerance;
        }

        private static IEnumerable<XYZ> GetBoundingBoxCorners(BoundingBoxXYZ boundingBox)
        {
            XYZ min = boundingBox.Min;
            XYZ max = boundingBox.Max;

            yield return new XYZ(min.X, min.Y, min.Z);
            yield return new XYZ(min.X, min.Y, max.Z);
            yield return new XYZ(min.X, max.Y, min.Z);
            yield return new XYZ(min.X, max.Y, max.Z);
            yield return new XYZ(max.X, min.Y, min.Z);
            yield return new XYZ(max.X, min.Y, max.Z);
            yield return new XYZ(max.X, max.Y, min.Z);
            yield return new XYZ(max.X, max.Y, max.Z);
        }

        private static QuickDimensionReadOnlyResult BuildResult(
            QuickDimensionLineContext lineContext,
            IReadOnlyList<QuickDimensionCandidate> rawCandidates,
            List<QuickDimensionDiagnostic> diagnostics,
            int collectedDoorCount,
            int collectedWindowCount,
            QuickDimensionOptions options)
        {
            IReadOnlyList<QuickDimensionCandidate> candidates = QuickDimensionGeometryService
                .DeduplicateCandidates(rawCandidates, options.DuplicateTolerance);

            AddDuplicateDiagnostics(rawCandidates, candidates, diagnostics);
            AddAcceptedDiagnostics(candidates, diagnostics);

            int acceptedDoorElementCount = CountAcceptedElements(candidates, QuickDimensionSourceType.Door);
            int acceptedWindowElementCount = CountAcceptedElements(candidates, QuickDimensionSourceType.Window);

            var summaries = new[]
            {
                new QuickDimensionSourceSummary(
                    QuickDimensionSourceType.Door,
                    collectedDoorCount,
                    acceptedDoorElementCount,
                    Math.Max(0, collectedDoorCount - acceptedDoorElementCount)),
                new QuickDimensionSourceSummary(
                    QuickDimensionSourceType.Window,
                    collectedWindowCount,
                    acceptedWindowElementCount,
                    Math.Max(0, collectedWindowCount - acceptedWindowElementCount))
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
                    $"Removed duplicate {rawCandidate.SourceType} opening candidate '{rawCandidate.DisplayName}' during source-aware deduplication.",
                    rawCandidate.ElementId,
                    rawCandidate.SourceType,
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
                    $"Accepted {candidate.SourceType} opening candidate '{candidate.DisplayName}' at parameter {candidate.ParameterOnDimensionLine:0.####}. Strategy: {candidate.ReferenceStrategy}. Host Wall: {candidate.HostElementValue}.",
                    candidate.ElementId,
                    candidate.SourceType,
                    candidate.DisplayName));
            }
        }

        private static int CountAcceptedElements(IReadOnlyList<QuickDimensionCandidate> candidates, QuickDimensionSourceType sourceType)
        {
            return candidates
                .Where(candidate => candidate.SourceType == sourceType)
                .Select(candidate => candidate.ElementValue)
                .Distinct()
                .Count();
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

        private static bool IsPointInsideBoundingBox(XYZ point, XYZ minimum, XYZ maximum)
        {
            return point.X >= minimum.X && point.X <= maximum.X
                && point.Y >= minimum.Y && point.Y <= maximum.Y
                && point.Z >= minimum.Z && point.Z <= maximum.Z;
        }

        private static string BuildOpeningDisplayName(
            QuickDimensionSourceType sourceType,
            string familyName,
            string typeName,
            ElementId elementId)
        {
            string sourceLabel = sourceType == QuickDimensionSourceType.Door ? "Door" : "Window";
            string familyLabel = string.IsNullOrWhiteSpace(familyName) ? "Unknown Family" : familyName;
            string typeLabel = string.IsNullOrWhiteSpace(typeName) ? "Unknown Type" : typeName;
            return $"{sourceLabel}: {familyLabel} - {typeLabel} ({elementId.Value})";
        }

        private static string GetFamilyName(FamilyInstance instance)
        {
            FamilySymbol? symbol = instance.Symbol;
            Family? family = symbol?.Family;
            return family?.Name ?? string.Empty;
        }

        private static string GetTypeName(FamilyInstance instance)
        {
            FamilySymbol? symbol = instance.Symbol;
            return symbol?.Name ?? string.Empty;
        }

        private sealed class OpeningReferencePair
        {
            public static OpeningReferencePair Empty { get; } = new OpeningReferencePair(null, null, null, null, null);

            public OpeningReferencePair(
                Reference? leftReference,
                Reference? rightReference,
                XYZ? leftPoint,
                XYZ? rightPoint,
                OpeningGeometryValidation? validation = null)
            {
                LeftReference = leftReference;
                RightReference = rightReference;
                LeftPoint = leftPoint;
                RightPoint = rightPoint;
                Validation = validation;
            }

            public Reference? LeftReference { get; }
            public Reference? RightReference { get; }
            public XYZ? LeftPoint { get; }
            public XYZ? RightPoint { get; }
            public OpeningGeometryValidation? Validation { get; }
        }

        /// <summary>
        /// Read-only evidence about how a hosted opening's jamb pair was resolved from host-wall geometry.
        /// Spans are stored in Revit internal feet; diagnostics convert to millimeters for operators.
        /// </summary>
        private sealed class OpeningGeometryValidation
        {
            public OpeningGeometryValidation(
                double expectedSpan,
                double boundingBoxSpan,
                double? rawSpan,
                double? selectedSpan,
                double? typeWidth,
                bool selectedPairValidated,
                bool useFallbackReferences)
            {
                ExpectedSpan = expectedSpan;
                BoundingBoxSpan = boundingBoxSpan;
                RawSpan = rawSpan;
                SelectedSpan = selectedSpan;
                TypeWidth = typeWidth;
                SelectedPairValidated = selectedPairValidated;
                UseFallbackReferences = useFallbackReferences;
            }

            public double ExpectedSpan { get; }
            public double BoundingBoxSpan { get; }
            public double? RawSpan { get; }
            public double? SelectedSpan { get; }
            public double? TypeWidth { get; }
            public bool SelectedPairValidated { get; }
            public bool UseFallbackReferences { get; }
        }

        private sealed class OpeningReferenceInfo
        {
            public OpeningReferenceInfo(
                Reference reference,
                XYZ referencePoint,
                string label,
                QuickDimensionReferenceStrategy referenceStrategy)
            {
                Reference = reference ?? throw new ArgumentNullException(nameof(reference));
                ReferencePoint = referencePoint ?? throw new ArgumentNullException(nameof(referencePoint));
                Label = label ?? string.Empty;
                ReferenceStrategy = referenceStrategy;
            }

            public Reference Reference { get; }
            public XYZ ReferencePoint { get; }
            public string Label { get; }
            public QuickDimensionReferenceStrategy ReferenceStrategy { get; }
        }

        private sealed class OpeningEdgeCandidate
        {
            public OpeningEdgeCandidate(Reference reference, XYZ midpoint, double positionOnDimensionDirection)
            {
                Reference = reference ?? throw new ArgumentNullException(nameof(reference));
                Midpoint = midpoint ?? throw new ArgumentNullException(nameof(midpoint));
                PositionOnDimensionDirection = positionOnDimensionDirection;
            }

            public Reference Reference { get; }
            public XYZ Midpoint { get; }
            public double PositionOnDimensionDirection { get; }
        }
    }
}
