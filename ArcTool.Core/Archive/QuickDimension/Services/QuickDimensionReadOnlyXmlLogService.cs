#nullable enable
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
    /// Writes the production Quick Dimension read-only summary XML log.
    /// This service serializes the already-collected read-only result and wall-axis aggregation trace only;
    /// it performs no geometry collection, no transaction, and never creates a Revit dimension.
    /// </summary>
    public static class QuickDimensionReadOnlyXmlLogService
    {
        private const string DateTimeFormat = "yyyy-MM-ddTHH:mm:ss.fffK";

        public static string WriteReadOnlySummaryLog(
            Document doc,
            RevitView activeView,
            Wall selectedWall,
            QuickDimensionReadOnlyResult result)
        {
            if (doc == null) throw new ArgumentNullException(nameof(doc));
            if (activeView == null) throw new ArgumentNullException(nameof(activeView));
            if (selectedWall == null) throw new ArgumentNullException(nameof(selectedWall));
            if (result == null) throw new ArgumentNullException(nameof(result));

            if (string.IsNullOrWhiteSpace(doc.PathName))
            {
                throw new InvalidOperationException("The Revit document must be saved before writing the Quick Dimension read-only XML log next to the .rvt file.");
            }

            if (selectedWall.Location is not LocationCurve locationCurve || locationCurve.Curve is not Line wallLine)
            {
                throw new InvalidOperationException("The selected wall does not expose a straight LocationCurve line for Quick Dimension read-only XML logging.");
            }

            QuickDimensionLineContext lineContext = result.LineContext;
            XYZ wallStart = wallLine.GetEndPoint(0);
            XYZ wallEnd = wallLine.GetEndPoint(1);
            string sideLabel = GetSideLabel(lineContext.SideSign);
            XYZ? sideNormal = lineContext.SideNormal;

            XElement root = new XElement(
                "QuickDimensionReadOnlySummaryLog",
                new XAttribute("createdAt", DateTimeOffset.Now.ToString(DateTimeFormat, CultureInfo.InvariantCulture)),
                new XAttribute("coordinateBasis", "Survey coordinates from Document.ActiveProjectLocation.GetProjectPosition"),
                new XAttribute("coordinateUnits", "meters primary; millimeters included as *_mm attributes"),
                new XAttribute("northField", "ProjectPosition.NorthSouth"),
                new XAttribute("eastField", "ProjectPosition.EastWest"),
                new XAttribute("selectedWallId", selectedWall.Id.Value),
                new XAttribute("selectedWallType", selectedWall.WallType?.Name ?? string.Empty),
                new XAttribute("activeViewId", activeView.Id.Value),
                new XAttribute("activeViewName", activeView.Name ?? string.Empty),
                new XAttribute("activeViewType", activeView.ViewType),
                new XAttribute("sideSign", lineContext.SideSign),
                new XAttribute("sideLabel", sideLabel),
                new XAttribute("selectedShellLayer", result.WallAxisAggregationTrace?.ShellLayer?.ToString() ?? string.Empty),
                new XAttribute("axisLengthMm", FormatMm(UnitUtils.ConvertFromInternalUnits(lineContext.Length, UnitTypeId.Millimeters))),
                new XAttribute("finalCandidateCount", result.CandidateCount),
                new XAttribute("canCreateChainDimension", result.CanCreateChainDimension));

            // Keep the top-level block order close to QuickDimensionWallSpikeXmlLogService:
            // result summary + anchors, selected wall, mid-run evidence, then read-only production extras.
            root.Add(BuildReadOnlyResultElement(doc, result));
            root.Add(BuildSelectedWallElement(doc, selectedWall, wallStart, wallEnd, lineContext.Direction, sideNormal));
            root.Add(BuildWallMidRunAggregationElement(doc, result.WallAxisAggregationTrace));
            root.Add(BuildFinalCandidatesElement(doc, result.Candidates));
            root.Add(BuildDiagnosticsElement(result.Diagnostics));

            string directory = Path.GetDirectoryName(doc.PathName);
            if (string.IsNullOrWhiteSpace(directory))
            {
                throw new InvalidOperationException("Could not resolve the folder of the active .rvt file.");
            }

            string safeTimestamp = DateTime.Now.ToString("yyyyMMdd_HHmmss", CultureInfo.InvariantCulture);
            string fileName = $"ArcTool_QD_ReadOnlySummary_{selectedWall.Id.Value}_{sideLabel}_{safeTimestamp}.xml";
            string path = Path.Combine(directory, fileName);

            XDocument document = new XDocument(new XDeclaration("1.0", "utf-8", "yes"), root);
            document.Save(path);
            return path;
        }

        private static XElement BuildReadOnlyResultElement(Document doc, QuickDimensionReadOnlyResult result)
        {
            QuickDimensionWallAxisAggregationTrace? trace = result.WallAxisAggregationTrace;
            return new XElement(
                "ReadOnlyResult",
                new XAttribute("succeeded", true),
                new XAttribute("side", trace?.SideLabel ?? GetSideLabel(result.LineContext.SideSign)),
                new XAttribute("sideSign", result.LineContext.SideSign),
                new XAttribute("selectedShellLayer", trace?.ShellLayer?.ToString() ?? string.Empty),
                new XAttribute("axisLengthMm", FormatMm(UnitUtils.ConvertFromInternalUnits(result.LineContext.Length, UnitTypeId.Millimeters))),
                new XAttribute("finalCandidateCount", result.CandidateCount),
                new XAttribute("canCreateChainDimension", result.CanCreateChainDimension),
                new XAttribute("diagnosticCount", result.DiagnosticCount),
                new XAttribute("message", trace?.Message ?? "Quick Dimension read-only engine completed."),
                BuildAnchorElement(doc, "StartAnchor", trace?.StartAnchor),
                BuildAnchorElement(doc, "FinishAnchor", trace?.FinishAnchor),
                BuildOptionsElement(result.Options, result.LineContext),
                BuildPerformanceTimingsElement(result.TimingTrace));
        }

        private static XElement BuildOptionsElement(
            QuickDimensionOptions options,
            QuickDimensionLineContext lineContext)
        {
            bool effectiveIncludeGrids = options.IncludeGrids && !lineContext.IsWallAxis;

            return new XElement(
                "Options",
                new XAttribute("includeGrids", effectiveIncludeGrids),
                new XAttribute("includeWalls", options.IncludeWalls),
                new XAttribute("includeDoors", options.IncludeDoors),
                new XAttribute("includeWindows", options.IncludeWindows),
                new XAttribute("enableHostWallOpeningFallback", options.EnableHostWallOpeningFallback),
                new XAttribute("projectionTolerance", FormatDouble(options.ProjectionTolerance)),
                new XAttribute("duplicateTolerance", FormatDouble(options.DuplicateTolerance)),
                new XAttribute("minimumDimensionLineLength", FormatDouble(options.MinimumDimensionLineLength)),
                new XAttribute("wallEndStationTolerance", FormatDouble(options.WallEndStationTolerance)));
        }

        private static XElement BuildPerformanceTimingsElement(QuickDimensionCollectionTimingTrace? timingTrace)
        {
            if (timingTrace == null)
            {
                return new XElement("PerformanceTimings", new XAttribute("captured", false));
            }

            return new XElement(
                "PerformanceTimings",
                new XAttribute("captured", true),
                new XAttribute("totalWallAxisCollectionMs", FormatDouble(timingTrace.TotalWallAxisCollectionMilliseconds)),
                new XAttribute("wallEndAnchorCollectionMs", FormatDouble(timingTrace.WallEndAnchorCollectionMilliseconds)),
                new XAttribute("midRunAggregationMs", FormatDouble(timingTrace.MidRunAggregationMilliseconds)),
                new XAttribute("openingCollectionMs", FormatDouble(timingTrace.OpeningCollectionMilliseconds)),
                new XAttribute("duplicateStationReductionMs", FormatDouble(timingTrace.DuplicateStationReductionMilliseconds)));
        }

        private static XElement BuildSelectedWallElement(
            Document doc,
            Wall selectedWall,
            XYZ wallStart,
            XYZ wallEnd,
            XYZ axisDirection,
            XYZ? sideNormal)
        {
            XElement element = new XElement(
                "SelectedWall",
                new XAttribute("id", selectedWall.Id.Value),
                new XAttribute("typeName", selectedWall.WallType?.Name ?? string.Empty),
                new XAttribute("widthMm", FormatMm(UnitUtils.ConvertFromInternalUnits(selectedWall.Width, UnitTypeId.Millimeters))),
                BuildPointElement(doc, "LocationCurveStart", wallStart, null, "location-curve"),
                BuildPointElement(doc, "LocationCurveEnd", wallEnd, null, "location-curve"),
                BuildVectorElement("AxisDirection", axisDirection));

            if (sideNormal != null)
            {
                element.Add(BuildVectorElement("SideNormal", sideNormal));
            }
            else
            {
                element.Add(new XElement("SideNormal", new XAttribute("resolved", false)));
            }

            return element;
        }

        private static XElement BuildResolvedAnchorsElement(
            Document doc,
            QuickDimensionWallAxisAggregationTrace? trace)
        {
            return new XElement(
                "ResolvedAnchors",
                BuildAnchorElement(doc, "StartAnchor", trace?.StartAnchor),
                BuildAnchorElement(doc, "FinishAnchor", trace?.FinishAnchor));
        }

        private static XElement BuildAnchorElement(
            Document doc,
            string name,
            QuickDimensionWallAxisAnchorTrace? anchor)
        {
            XElement element = new XElement(name);
            if (anchor == null)
            {
                element.SetAttributeValue("resolved", false);
                return element;
            }

            element.SetAttributeValue("resolved", true);
            element.SetAttributeValue("label", anchor.Label ?? string.Empty);
            element.SetAttributeValue("stationMm", FormatMm(UnitUtils.ConvertFromInternalUnits(anchor.StationOnWallAxis, UnitTypeId.Millimeters)));
            element.SetAttributeValue("elementId", GetReferenceElementValue(anchor.EdgeReference));
            element.SetAttributeValue("hasReference", anchor.HasReference);
            AddStableReferenceAttributes(doc, element, anchor.EdgeReference);

            if (anchor.Point != null)
            {
                element.Add(BuildPointElement(doc, "Point", anchor.Point, anchor.StationOnWallAxis, "resolved-anchor"));
            }

            return element;
        }

        private static XElement BuildWallMidRunAggregationElement(
            Document doc,
            QuickDimensionWallAxisAggregationTrace? trace)
        {
            if (trace == null)
            {
                return new XElement(
                    "WallMidRunAggregation",
                    new XAttribute("supported", false),
                    new XAttribute("message", "No wall-axis aggregation trace was attached to the read-only result."));
            }

            XElement element = new XElement(
                "WallMidRunAggregation",
                new XAttribute("supported", trace.Supported),
                new XAttribute("side", trace.SideLabel ?? string.Empty),
                new XAttribute("sideSign", trace.SideSign),
                new XAttribute("shell", trace.ShellLayer?.ToString() ?? string.Empty),
                new XAttribute("axisLengthMm", FormatMm(UnitUtils.ConvertFromInternalUnits(trace.AxisLength, UnitTypeId.Millimeters))),
                new XAttribute("candidateCount", trace.Candidates.Count),
                new XAttribute("message", trace.Message ?? string.Empty),
                new XAttribute("selectedWallId", trace.SelectedWallId),
                new XAttribute("resolvedAnchorMinStationMm", FormatMm(UnitUtils.ConvertFromInternalUnits(trace.ResolvedAnchorMinStation, UnitTypeId.Millimeters))),
                new XAttribute("resolvedAnchorMaxStationMm", FormatMm(UnitUtils.ConvertFromInternalUnits(trace.ResolvedAnchorMaxStation, UnitTypeId.Millimeters))),
                BuildJoinProvenanceElement("ElementsAtJoinStart", trace.ElementsAtJoinStartIds),
                BuildJoinProvenanceElement("ElementsAtJoinEnd", trace.ElementsAtJoinEndIds),
                BuildJoinProvenanceElement("GeometryJoin", trace.GeometryJoinIds));

            if (trace.SideNormal != null)
            {
                element.Add(BuildVectorElement("SideNormal", trace.SideNormal));
            }

            if (trace.SideLinePoint != null)
            {
                element.Add(BuildPointElement(doc, "SideLinePoint", trace.SideLinePoint, null, "selected-side-line"));
            }

            element.Add(new XElement(
                "Candidates",
                new XAttribute("count", trace.Candidates.Count),
                trace.Candidates.Select((candidate, index) => BuildMidRunCandidateElement(doc, index + 1, candidate))));

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
            QuickDimensionWallAxisCandidateTrace candidate)
        {
            XElement element = new XElement(
                "Candidate",
                new XAttribute("index", index),
                new XAttribute("wallId", candidate.CandidateWallId),
                new XAttribute("typeName", candidate.CandidateTypeName ?? string.Empty),
                new XAttribute("relation", candidate.Relation),
                new XAttribute("classification", candidate.Relation),
                new XAttribute("inElementsAtJoinStart", candidate.InElementsAtJoinStart),
                new XAttribute("inElementsAtJoinEnd", candidate.InElementsAtJoinEnd),
                new XAttribute("inGeometryJoin", candidate.InGeometryJoin),
                new XAttribute("isPerpendicular", candidate.IsPerpendicular),
                new XAttribute("isParallel", candidate.IsParallel),
                new XAttribute("referenceHitCount", candidate.ReferenceHitCount),
                new XAttribute("acceptedMidRunStationCount", candidate.AcceptedMidRunStationCount),
                new XAttribute("candidateWallExposesRefAtStation", candidate.ReferenceHits.Count > 0),
                new XAttribute("candidateReferenceNormalAlongAxis", candidate.ReferenceHits.Any(hit => hit.CandidateReferenceNormalAlongAxis)),
                new XAttribute("selectedWallExposesRefAtStation", candidate.ReferenceHits.Any(hit => hit.SelectedWallExposesRefAtStation)),
                new XAttribute("selectedReferenceNormalAlongAxis", candidate.ReferenceHits.Any(hit => hit.SelectedReferenceNormalAlongAxis)),
                new XAttribute("fallbackStationMm", FormatMm(UnitUtils.ConvertFromInternalUnits(candidate.FallbackStationOnSelectedAxis, UnitTypeId.Millimeters))),
                new XAttribute("fallbackDistanceToSideLineMm", FormatMm(UnitUtils.ConvertFromInternalUnits(candidate.FallbackDistanceToSideLine, UnitTypeId.Millimeters))),
                new XAttribute("source", candidate.ReferenceHits.Count > 0 ? "vertical-edge-on-side-line" : "location-curve-endpoint"),
                new XAttribute("reason", candidate.RejectedReason ?? string.Empty));

            element.Add(new XElement(
                "ReferenceHits",
                new XAttribute("count", candidate.ReferenceHits.Count),
                candidate.ReferenceHits.Select(hit => BuildMidRunReferenceHitElement(doc, hit))));
            return element;
        }

        private static XElement BuildMidRunReferenceHitElement(
            Document doc,
            QuickDimensionWallAxisReferenceHitTrace hit)
        {
            XElement element = new XElement(
                "ReferenceHit",
                new XAttribute("index", hit.Index),
                new XAttribute("stationMm", FormatMm(UnitUtils.ConvertFromInternalUnits(hit.StationOnSelectedAxis, UnitTypeId.Millimeters))),
                new XAttribute("distanceToSideLineMm", FormatMm(UnitUtils.ConvertFromInternalUnits(hit.DistanceToSideLine, UnitTypeId.Millimeters))),
                new XAttribute("candidateReferenceNormalAlongAxis", hit.CandidateReferenceNormalAlongAxis),
                new XAttribute("selectedWallExposesRefAtStation", hit.SelectedWallExposesRefAtStation),
                new XAttribute("selectedReferenceNormalAlongAxis", hit.SelectedReferenceNormalAlongAxis),
                new XAttribute("accepted", hit.Accepted),
                new XAttribute("rejectedReason", hit.RejectedReason ?? string.Empty),
                new XAttribute("hasReference", hit.EdgeReference != null),
                new XAttribute("elementId", GetReferenceElementValue(hit.EdgeReference)));

            AddStableReferenceAttributes(doc, element, hit.EdgeReference);

            if (hit.Point != null)
            {
                element.Add(BuildPointElement(doc, "Point", hit.Point, hit.StationOnSelectedAxis, "vertical-edge-on-side-line"));
            }

            return element;
        }

        private static XElement BuildFinalCandidatesElement(Document doc, IReadOnlyList<QuickDimensionCandidate> candidates)
        {
            return new XElement(
                "FinalCandidates",
                new XAttribute("count", candidates.Count),
                candidates.Select((candidate, index) => BuildFinalCandidateElement(doc, index + 1, candidate)));
        }

        private static XElement BuildFinalCandidateElement(Document doc, int index, QuickDimensionCandidate candidate)
        {
            XElement element = new XElement(
                "Candidate",
                new XAttribute("index", index),
                new XAttribute("stationMm", FormatMm(UnitUtils.ConvertFromInternalUnits(candidate.ParameterOnDimensionLine, UnitTypeId.Millimeters))),
                new XAttribute("sourceType", candidate.SourceType),
                new XAttribute("displayName", candidate.DisplayName ?? string.Empty),
                new XAttribute("elementId", candidate.ElementValue),
                new XAttribute("hostElementId", candidate.HostElementValue?.ToString(CultureInfo.InvariantCulture) ?? string.Empty),
                new XAttribute("referenceStrategy", candidate.ReferenceStrategy),
                new XAttribute("hasReference", candidate.Reference != null),
                new XAttribute("familyName", candidate.FamilyName ?? string.Empty),
                new XAttribute("typeName", candidate.TypeName ?? string.Empty),
                new XAttribute("dedupeStatus", "Kept"));

            AddStableReferenceAttributes(doc, element, candidate.Reference);
            element.Add(BuildPointElement(doc, "Point", candidate.HitPoint, candidate.ParameterOnDimensionLine, "final-candidate"));
            return element;
        }

        private static XElement BuildDiagnosticsElement(IReadOnlyList<QuickDimensionDiagnostic> diagnostics)
        {
            return new XElement(
                "Diagnostics",
                new XAttribute("count", diagnostics.Count),
                diagnostics.Select((diagnostic, index) => new XElement(
                    "Diagnostic",
                    new XAttribute("index", index + 1),
                    new XAttribute("severity", diagnostic.Severity),
                    new XAttribute("reason", diagnostic.Reason),
                    new XAttribute("isRejected", diagnostic.IsRejected),
                    new XAttribute("elementId", diagnostic.ElementValue?.ToString(CultureInfo.InvariantCulture) ?? string.Empty),
                    new XAttribute("sourceType", diagnostic.SourceType?.ToString() ?? string.Empty),
                    new XAttribute("displayName", diagnostic.DisplayName ?? string.Empty),
                    new XAttribute("message", diagnostic.Message ?? string.Empty))));
        }

        private static XElement BuildPointElement(Document doc, string name, XYZ point, double? station, string source)
        {
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

        private static XElement BuildVectorElement(string name, XYZ vector)
        {
            return new XElement(
                name,
                new XAttribute("x", FormatDouble(vector.X)),
                new XAttribute("y", FormatDouble(vector.Y)),
                new XAttribute("z", FormatDouble(vector.Z)));
        }

        private static void AddStableReferenceAttributes(Document doc, XElement element, Reference? reference)
        {
            if (reference == null)
            {
                element.SetAttributeValue("stableReference", string.Empty);
                element.SetAttributeValue("stableReferenceError", string.Empty);
                return;
            }

            try
            {
                element.SetAttributeValue("stableReference", reference.ConvertToStableRepresentation(doc) ?? string.Empty);
                element.SetAttributeValue("stableReferenceError", string.Empty);
            }
            catch (Exception ex)
            {
                element.SetAttributeValue("stableReference", string.Empty);
                element.SetAttributeValue("stableReferenceError", ex.Message ?? string.Empty);
            }
        }

        // === CHAIN CREATION AUDIT (Phase 3 NewDimension runtime evidence) ===
        // Reads back the committed Dimension by id after QuickDimensionChainCreationService returns.
        // Read-only: no transaction, no mutation of Candidates/aggregation. Appended to the SAME file
        // already written by WriteReadOnlySummaryLog via atomic temp+replace, so a failed append never
        // corrupts the original read-only XML. Creation status and audit status are reported independently.
        private const double ChainCreationAuditSegmentToleranceMm = 0.1;

        public static string TryAppendChainCreationAudit(
            Document doc,
            string readOnlyXmlPath,
            QuickDimensionReadOnlyResult result,
            QuickDimensionChainCreationResult creationResult)
        {
            if (doc == null) throw new ArgumentNullException(nameof(doc));
            if (result == null) throw new ArgumentNullException(nameof(result));
            if (creationResult == null) throw new ArgumentNullException(nameof(creationResult));

            if (string.IsNullOrWhiteSpace(readOnlyXmlPath) || !File.Exists(readOnlyXmlPath))
            {
                return "Chain creation audit skipped: the read-only XML log path is unavailable.";
            }

            try
            {
                XDocument document = XDocument.Load(readOnlyXmlPath);
                if (document.Root == null)
                {
                    return "Chain creation audit skipped: the read-only XML log has no root element.";
                }

                document.Root.Add(BuildChainCreationAuditElement(doc, result, creationResult));

                string tempPath = readOnlyXmlPath + ".audit.tmp";
                document.Save(tempPath);
                File.Replace(tempPath, readOnlyXmlPath, destinationBackupFileName: null);

                return $"Chain creation audit appended to {readOnlyXmlPath}.";
            }
            catch (Exception ex)
            {
                return $"Chain creation audit failed (original read-only XML left untouched): {ex.Message}";
            }
        }
        private static XElement BuildChainCreationAuditElement(
            Document doc,
            QuickDimensionReadOnlyResult result,
            QuickDimensionChainCreationResult creationResult)
        {
            IReadOnlyList<QuickDimensionCandidate> expectedCandidates = result.Candidates;
            int expectedReferenceCount = expectedCandidates.Count;
            int expectedSegmentCount = Math.Max(0, expectedReferenceCount - 1);

            List<string> expectedStableReferences = expectedCandidates
                .Select(candidate => TryGetStableReference(doc, candidate.Reference, out _))
                .ToList();
            List<long?> expectedCandidateOwners = expectedCandidates
                .Select(candidate => (long?)candidate.ElementValue)
                .ToList();
            List<long?> expectedReferenceOwners = expectedCandidates
                .Select(candidate => TryGetReferenceElementValue(candidate.Reference))
                .ToList();

            Dimension? dimension = null;
            string dimensionReadError = string.Empty;
            if (creationResult.DimensionId != null)
            {
                try
                {
                    dimension = doc.GetElement(creationResult.DimensionId) as Dimension;
                    if (creationResult.Succeeded && dimension == null)
                    {
                        dimensionReadError = "The committed dimension could not be resolved from its ElementId.";
                    }
                }
                catch (Exception ex)
                {
                    dimensionReadError = ex.Message ?? string.Empty;
                }
            }

            List<Reference> createdReferences = GetDimensionReferences(dimension);
            List<string> createdStableReferences = createdReferences
                .Select(reference => TryGetStableReference(doc, reference, out _))
                .ToList();
            List<long?> createdReferenceOwners = createdReferences
                .Select(TryGetReferenceElementValue)
                .ToList();

            string referenceOrderRelation = GetReferenceOrderRelation(
                expectedStableReferences,
                createdStableReferences);
            bool referenceIdentityMatched = referenceOrderRelation == "Exact" || referenceOrderRelation == "Reversed";
            bool referenceOwnersMatched = referenceIdentityMatched && ReferenceOwnersMatch(
                expectedReferenceOwners,
                createdReferenceOwners,
                referenceOrderRelation);

            DimensionSegmentArray? segmentArray = dimension?.Segments;
            int segmentArraySize = segmentArray?.Size ?? 0;
            List<(double? ValueInternal, string ValueSource)> segmentMeasurements = ExtractSegmentValues(dimension, segmentArray, expectedReferenceCount);
            List<double?> segmentValuesInternal = segmentMeasurements.Select(static measurement => measurement.ValueInternal).ToList();
            int actualSegmentCount = segmentValuesInternal.Count;
            bool segmentValuesMatched = SegmentValuesMatch(
                result,
                segmentValuesInternal,
                referenceOrderRelation);

            XElement audit = new XElement(
                "ChainCreationAudit",
                new XAttribute("attempted", creationResult.TransactionStatus != TransactionStatus.Uninitialized),
                new XAttribute("succeeded", creationResult.Succeeded),
                new XAttribute("message", creationResult.Message ?? string.Empty),
                new XAttribute("transactionStatus", creationResult.TransactionStatus),
                new XAttribute("dimensionId", creationResult.DimensionId?.Value.ToString(CultureInfo.InvariantCulture) ?? string.Empty),
                new XAttribute("expectedReferenceCount", expectedReferenceCount),
                new XAttribute("createdReferenceCount", createdReferences.Count),
                new XAttribute("expectedSegmentCount", expectedSegmentCount),
                new XAttribute("actualSegmentCount", actualSegmentCount),
                new XAttribute("referenceOrderRelation", referenceOrderRelation),
                new XAttribute("referenceIdentityMatched", referenceIdentityMatched),
                new XAttribute("referenceOwnersMatched", referenceOwnersMatched),
                new XAttribute("segmentValuesMatched", segmentValuesMatched),
                new XAttribute("dimensionReadError", dimensionReadError),
                BuildResolvedDimensionLineElement(creationResult),
                BuildExpectedCandidatesAuditElement(doc, expectedCandidates, expectedReferenceOwners),
                BuildCreatedReferencesAuditElement(
                    doc,
                    expectedCandidates,
                    expectedStableReferences,
                    expectedCandidateOwners,
                    createdReferences,
                    createdStableReferences,
                    createdReferenceOwners,
                    referenceOrderRelation),
                BuildSegmentsAuditElement(result, segmentMeasurements, actualSegmentCount, referenceOrderRelation));

            audit.Element("Segments")?.SetAttributeValue("segmentArraySize", segmentArraySize);
            audit.Element("Segments")?.SetAttributeValue("toleranceMm", FormatMm(ChainCreationAuditSegmentToleranceMm));
            return audit;
        }

        private static XElement BuildResolvedDimensionLineElement(QuickDimensionChainCreationResult creationResult)
        {
            double? minimumStationMm = ConvertInternalToMm(creationResult.MinimumStation);
            double? maximumStationMm = ConvertInternalToMm(creationResult.MaximumStation);
            double? spanMm = minimumStationMm.HasValue && maximumStationMm.HasValue
                ? maximumStationMm.Value - minimumStationMm.Value
                : null;

            return new XElement(
                "ResolvedDimensionLine",
                new XAttribute("minimumStationMm", FormatNullableMm(minimumStationMm)),
                new XAttribute("maximumStationMm", FormatNullableMm(maximumStationMm)),
                new XAttribute("spanMm", FormatNullableMm(spanMm)),
                new XAttribute("sideOffsetMm", FormatNullableMm(ConvertInternalToMm(creationResult.SideOffset))));
        }

        private static XElement BuildExpectedCandidatesAuditElement(
            Document doc,
            IReadOnlyList<QuickDimensionCandidate> candidates,
            IReadOnlyList<long?> expectedReferenceOwners)
        {
            return new XElement(
                "ExpectedCandidates",
                new XAttribute("count", candidates.Count),
                candidates.Select((candidate, index) =>
                {
                    double? nextDeltaMm = index + 1 < candidates.Count
                        ? UnitUtils.ConvertFromInternalUnits(
                            Math.Abs(candidates[index + 1].ParameterOnDimensionLine - candidate.ParameterOnDimensionLine),
                            UnitTypeId.Millimeters)
                        : null;
                    long? referenceOwner = expectedReferenceOwners[index];
                    string stableReference = TryGetStableReference(doc, candidate.Reference, out string stableReferenceError);

                    return new XElement(
                        "Candidate",
                        new XAttribute("index", index + 1),
                        new XAttribute("stationMm", FormatMm(UnitUtils.ConvertFromInternalUnits(candidate.ParameterOnDimensionLine, UnitTypeId.Millimeters))),
                        new XAttribute("sourceType", candidate.SourceType),
                        new XAttribute("displayName", candidate.DisplayName ?? string.Empty),
                        new XAttribute("elementId", candidate.ElementValue),
                        new XAttribute("hostElementId", candidate.HostElementValue?.ToString(CultureInfo.InvariantCulture) ?? string.Empty),
                        new XAttribute("referenceStrategy", candidate.ReferenceStrategy),
                        new XAttribute("stableReference", stableReference),
                        new XAttribute("stableReferenceError", stableReferenceError),
                        new XAttribute("referenceOwnerElementId", FormatNullableElementId(referenceOwner)),
                        new XAttribute("elementIdMatchesReferenceOwner", referenceOwner.HasValue && candidate.ElementValue == referenceOwner.Value),
                        new XAttribute("expectedNextSegmentDeltaMm", FormatNullableMm(nextDeltaMm)));
                }));
        }

        private static XElement BuildCreatedReferencesAuditElement(
            Document doc,
            IReadOnlyList<QuickDimensionCandidate> expectedCandidates,
            IReadOnlyList<string> expectedStableReferences,
            IReadOnlyList<long?> expectedCandidateOwners,
            IReadOnlyList<Reference> createdReferences,
            IReadOnlyList<string> createdStableReferences,
            IReadOnlyList<long?> createdReferenceOwners,
            string referenceOrderRelation)
        {
            int expectedCount = expectedCandidates.Count;

            return new XElement(
                "CreatedReferences",
                new XAttribute("count", createdReferences.Count),
                createdReferences.Select((reference, index) =>
                {
                    string stableReference = createdStableReferences[index];
                    int? matchedExpectedIndex = FindMatchedExpectedIndex(
                        index,
                        stableReference,
                        expectedStableReferences,
                        expectedCount,
                        referenceOrderRelation);
                    long? owner = createdReferenceOwners[index];

                    // Owner equality compares the COMMITTED reference's live ElementId against the
                    // matched candidate's DECLARED metadata ElementId (QuickDimensionCandidate.ElementId),
                    // not against that candidate's own live reference owner (which would be tautological,
                    // since it is the same Reference instance round-tripped through NewDimension). This is
                    // the check that surfaces BUG-10 (HostWallOpeningGeometry fallback elementId = opening
                    // instance id, while the live reference is owned by the host wall) under real creation.
                    bool stableEqualityForward = matchedExpectedIndex.HasValue
                        && StableReferencesEqual(stableReference, expectedStableReferences[matchedExpectedIndex.Value]);
                    bool ownerEqualityMatched = matchedExpectedIndex.HasValue
                        && owner.HasValue
                        && expectedCandidateOwners[matchedExpectedIndex.Value].HasValue
                        && owner.Value == expectedCandidateOwners[matchedExpectedIndex.Value]!.Value;

                    string stableReferenceError = string.Empty;
                    try
                    {
                        if (reference != null)
                        {
                            reference.ConvertToStableRepresentation(doc);
                        }
                    }
                    catch (Exception ex)
                    {
                        stableReferenceError = ex.Message ?? string.Empty;
                    }

                    return new XElement(
                        "Reference",
                        new XAttribute("index", index + 1),
                        new XAttribute("elementId", FormatNullableElementId(TryGetReferenceElementValue(reference))),
                        new XAttribute("stableReference", stableReference),
                        new XAttribute("stableReferenceError", stableReferenceError),
                        new XAttribute("matchedExpectedCandidateIndex", matchedExpectedIndex.HasValue ? (matchedExpectedIndex.Value + 1).ToString(CultureInfo.InvariantCulture) : string.Empty),
                        new XAttribute("stableReferenceEqualToExpected", stableEqualityForward),
                        new XAttribute("ownerEqualToExpected", ownerEqualityMatched));
                }));
        }

        private static XElement BuildSegmentsAuditElement(
            QuickDimensionReadOnlyResult result,
            List<(double? ValueInternal, string ValueSource)> segmentMeasurements,
            int actualSegmentCount,
            string referenceOrderRelation)
        {
            List<double> expectedDeltasMm = BuildExpectedDeltasMm(result.Candidates);
            int expectedSegmentCount = expectedDeltasMm.Count;

            var segmentElements = new List<XElement>();
            for (int i = 0; i < segmentMeasurements.Count; i++)
            {
                double? expectedDeltaMm = MapSegmentIndexToExpectedDelta(i, segmentMeasurements.Count, expectedDeltasMm, referenceOrderRelation);
                double? actualMm = segmentMeasurements[i].ValueInternal.HasValue
                    ? UnitUtils.ConvertFromInternalUnits(segmentMeasurements[i].ValueInternal!.Value, UnitTypeId.Millimeters)
                    : (double?)null;
                double? deltaMm = actualMm.HasValue && expectedDeltaMm.HasValue
                    ? Math.Abs(actualMm.Value - expectedDeltaMm.Value)
                    : null;
                bool matched = deltaMm.HasValue && deltaMm.Value <= ChainCreationAuditSegmentToleranceMm;

                segmentElements.Add(new XElement(
                    "Segment",
                    new XAttribute("index", i + 1),
                    new XAttribute("valueSource", segmentMeasurements[i].ValueSource),
                    new XAttribute("actualValueMm", FormatNullableMm(actualMm)),
                    new XAttribute("expectedDeltaMm", FormatNullableMm(expectedDeltaMm)),
                    new XAttribute("differenceMm", FormatNullableMm(deltaMm)),
                    new XAttribute("matched", matched)));
            }

            return new XElement(
                "Segments",
                new XAttribute("expectedSegmentCount", expectedSegmentCount),
                new XAttribute("actualSegmentCount", actualSegmentCount),
                segmentElements);
        }

        private static List<Reference> GetDimensionReferences(Dimension? dimension)
        {
            var result = new List<Reference>();
            if (dimension == null)
            {
                return result;
            }

            try
            {
                ReferenceArray referenceArray = dimension.References;
                if (referenceArray == null)
                {
                    return result;
                }

                foreach (Reference reference in referenceArray)
                {
                    if (reference != null)
                    {
                        result.Add(reference);
                    }
                }
            }
            catch (Exception)
            {
                // Leave result as-is; the caller reports created count vs expected count as a mismatch.
            }

            return result;
        }

        /// <summary>
        /// Single-segment Revit dimensions (2 references) report NumberOfSegments == 1 with an EMPTY
        /// Segments array; the length lives on Dimension.Value instead. Multi-segment dimensions (3+
        /// references) populate Segments with NumberOfSegments entries. This helper normalizes both
        /// shapes into one ordered list of nullable internal-unit lengths with the original value source.
        /// </summary>
        private static List<(double? ValueInternal, string ValueSource)> ExtractSegmentValues(
            Dimension? dimension,
            DimensionSegmentArray? segmentArray,
            int expectedReferenceCount)
        {
            var values = new List<(double? ValueInternal, string ValueSource)>();
            if (dimension == null)
            {
                return values;
            }

            try
            {
                if (segmentArray != null && segmentArray.Size > 0)
                {
                    foreach (DimensionSegment segment in segmentArray)
                    {
                        values.Add((segment?.Value, "DimensionSegment.Value"));
                    }
                    return values;
                }

                if (expectedReferenceCount == 2)
                {
                    // Single-segment case: Segments is empty by design; read the whole-dimension value.
                    values.Add((dimension.Value, "Dimension.Value"));
                }
            }
            catch (Exception)
            {
                // Leave values as-is; caller reports actualSegmentCount vs expected as a mismatch.
            }

            return values;
        }

        private static string TryGetStableReference(Document doc, Reference? reference, out string error)
        {
            error = string.Empty;
            if (reference == null)
            {
                return string.Empty;
            }

            try
            {
                return reference.ConvertToStableRepresentation(doc) ?? string.Empty;
            }
            catch (Exception ex)
            {
                error = ex.Message ?? string.Empty;
                return string.Empty;
            }
        }

        private static long? TryGetReferenceElementValue(Reference? reference)
        {
            try
            {
                return reference?.ElementId?.Value;
            }
            catch
            {
                return null;
            }
        }

        private static string GetReferenceOrderRelation(
            IReadOnlyList<string> expectedStableReferences,
            IReadOnlyList<string> createdStableReferences)
        {
            if (expectedStableReferences.Count != createdStableReferences.Count)
            {
                return "Mismatch";
            }

            bool exact = true;
            bool reversed = true;
            int count = expectedStableReferences.Count;
            for (int i = 0; i < count; i++)
            {
                exact &= StableReferencesEqual(createdStableReferences[i], expectedStableReferences[i]);
                reversed &= StableReferencesEqual(createdStableReferences[i], expectedStableReferences[count - 1 - i]);
            }

            if (exact)
            {
                return "Exact";
            }

            return reversed ? "Reversed" : "Mismatch";
        }

        private static bool ReferenceOwnersMatch(
            IReadOnlyList<long?> expectedOwners,
            IReadOnlyList<long?> createdOwners,
            string referenceOrderRelation)
        {
            if (expectedOwners.Count != createdOwners.Count)
            {
                return false;
            }

            for (int i = 0; i < createdOwners.Count; i++)
            {
                int? expectedIndex = MapCreatedIndexToExpectedIndex(i, expectedOwners.Count, referenceOrderRelation);
                if (!expectedIndex.HasValue
                    || !createdOwners[i].HasValue
                    || !expectedOwners[expectedIndex.Value].HasValue
                    || createdOwners[i]!.Value != expectedOwners[expectedIndex.Value]!.Value)
                {
                    return false;
                }
            }

            return true;
        }

        private static int? FindMatchedExpectedIndex(
            int createdIndex,
            string createdStableReference,
            IReadOnlyList<string> expectedStableReferences,
            int expectedCount,
            string referenceOrderRelation)
        {
            int? mappedIndex = MapCreatedIndexToExpectedIndex(createdIndex, expectedCount, referenceOrderRelation);
            if (mappedIndex.HasValue)
            {
                return mappedIndex;
            }

            var matches = new List<int>();
            for (int i = 0; i < expectedStableReferences.Count; i++)
            {
                if (StableReferencesEqual(createdStableReference, expectedStableReferences[i]))
                {
                    matches.Add(i);
                }
            }

            return matches.Count == 1 ? matches[0] : null;
        }

        private static int? MapCreatedIndexToExpectedIndex(
            int createdIndex,
            int expectedCount,
            string referenceOrderRelation)
        {
            if (createdIndex < 0 || createdIndex >= expectedCount)
            {
                return null;
            }

            return referenceOrderRelation switch
            {
                "Exact" => createdIndex,
                "Reversed" => expectedCount - 1 - createdIndex,
                _ => null
            };
        }

        private static bool StableReferencesEqual(string left, string right)
        {
            return !string.IsNullOrWhiteSpace(left)
                && !string.IsNullOrWhiteSpace(right)
                && string.Equals(left, right, StringComparison.Ordinal);
        }

        private static List<double> BuildExpectedDeltasMm(IReadOnlyList<QuickDimensionCandidate> candidates)
        {
            var deltas = new List<double>();
            for (int i = 0; i + 1 < candidates.Count; i++)
            {
                deltas.Add(UnitUtils.ConvertFromInternalUnits(
                    Math.Abs(candidates[i + 1].ParameterOnDimensionLine - candidates[i].ParameterOnDimensionLine),
                    UnitTypeId.Millimeters));
            }

            return deltas;
        }

        private static double? MapSegmentIndexToExpectedDelta(
            int segmentIndex,
            int actualSegmentCount,
            IReadOnlyList<double> expectedDeltasMm,
            string referenceOrderRelation)
        {
            if (actualSegmentCount != expectedDeltasMm.Count
                || segmentIndex < 0
                || segmentIndex >= actualSegmentCount)
            {
                return null;
            }

            return referenceOrderRelation switch
            {
                "Exact" => expectedDeltasMm[segmentIndex],
                "Reversed" => expectedDeltasMm[expectedDeltasMm.Count - 1 - segmentIndex],
                _ => null
            };
        }

        private static bool SegmentValuesMatch(
            QuickDimensionReadOnlyResult result,
            IReadOnlyList<double?> actualSegmentValuesInternal,
            string referenceOrderRelation)
        {
            List<double> expectedDeltasMm = BuildExpectedDeltasMm(result.Candidates);
            if ((referenceOrderRelation != "Exact" && referenceOrderRelation != "Reversed")
                || actualSegmentValuesInternal.Count != expectedDeltasMm.Count)
            {
                return false;
            }

            for (int i = 0; i < actualSegmentValuesInternal.Count; i++)
            {
                double? expectedDeltaMm = MapSegmentIndexToExpectedDelta(
                    i,
                    actualSegmentValuesInternal.Count,
                    expectedDeltasMm,
                    referenceOrderRelation);
                double? actualValueInternal = actualSegmentValuesInternal[i];
                if (!expectedDeltaMm.HasValue || !actualValueInternal.HasValue)
                {
                    return false;
                }

                double actualMm = UnitUtils.ConvertFromInternalUnits(actualValueInternal.Value, UnitTypeId.Millimeters);
                if (Math.Abs(actualMm - expectedDeltaMm.Value) > ChainCreationAuditSegmentToleranceMm)
                {
                    return false;
                }
            }

            return true;
        }

        private static double? ConvertInternalToMm(double? value)
        {
            return value.HasValue
                ? UnitUtils.ConvertFromInternalUnits(value.Value, UnitTypeId.Millimeters)
                : (double?)null;
        }

        private static string FormatNullableMm(double? value)
        {
            return value.HasValue ? FormatMm(value.Value) : string.Empty;
        }

        private static string FormatNullableElementId(long? value)
        {
            return value?.ToString(CultureInfo.InvariantCulture) ?? string.Empty;
        }

        private static string GetReferenceElementValue(Reference? reference)
        {
            try
            {
                return reference?.ElementId?.Value.ToString(CultureInfo.InvariantCulture) ?? string.Empty;
            }
            catch
            {
                return string.Empty;
            }
        }

        private static string GetSideLabel(int sideSign)
        {
            return sideSign > 0 ? "Left" : sideSign < 0 ? "Right" : "Unspecified";
        }

        private static string FormatMm(double value)
        {
            return value.ToString("0.###", CultureInfo.InvariantCulture);
        }

        private static string FormatMeters(double millimeters)
        {
            return (millimeters / 1000.0).ToString("0.###", CultureInfo.InvariantCulture);
        }

        private static string FormatDouble(double value)
        {
            return value.ToString("0.#######", CultureInfo.InvariantCulture);
        }
    }
}
