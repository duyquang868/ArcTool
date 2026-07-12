#nullable enable
using System;
using System.Collections.Generic;
using System.Linq;
using Autodesk.Revit.DB;
using ArcTool.Core.Models;

namespace ArcTool.Core.Services
{
    /// <summary>
    /// Transaction-free geometry helpers for the Quick Dimension read-only engine.
    /// This service performs only math and candidate ordering/deduplication; it does not access documents or create dimensions.
    /// </summary>
    public static class QuickDimensionGeometryService
    {
        /// <summary>
        /// Returns true when all coordinates are finite real numbers.
        /// </summary>
        public static bool IsFinite(XYZ point)
        {
            if (point == null)
            {
                return false;
            }

            return IsFinite(point.X) && IsFinite(point.Y) && IsFinite(point.Z);
        }

        /// <summary>
        /// Attempts to extract a straight curve's endpoints. Arc and other non-line curves are rejected by design.
        /// </summary>
        public static bool TryGetStraightCurveEndpoints(Curve curve, out XYZ startPoint, out XYZ endPoint)
        {
            startPoint = null!;
            endPoint = null!;

            if (curve is not Line line)
            {
                return false;
            }

            XYZ start = line.GetEndPoint(0);
            XYZ end = line.GetEndPoint(1);
            if (!IsFinite(start) || !IsFinite(end))
            {
                return false;
            }

            startPoint = start;
            endPoint = end;
            return true;
        }

        /// <summary>
        /// Attempts to normalize the XY direction from start to end.
        /// Z is intentionally ignored because the Quick Dimension MVP is plan-view only.
        /// </summary>
        public static bool TryGetPlanarDirection(XYZ startPoint, XYZ endPoint, double tolerance, out XYZ direction)
        {
            direction = null!;
            ValidateNonNegativeTolerance(tolerance, nameof(tolerance));

            if (!IsFinite(startPoint) || !IsFinite(endPoint))
            {
                return false;
            }

            double deltaX = endPoint.X - startPoint.X;
            double deltaY = endPoint.Y - startPoint.Y;
            double length = Math.Sqrt(deltaX * deltaX + deltaY * deltaY);
            if (length <= tolerance)
            {
                return false;
            }

            direction = new XYZ(deltaX / length, deltaY / length, 0.0);
            return true;
        }

        /// <summary>
        /// Returns true when two XY directions are nearly parallel or anti-parallel.
        /// </summary>
        public static bool IsNearlyParallel(XYZ firstDirection, XYZ secondDirection, double tolerance)
        {
            ValidateNonNegativeTolerance(tolerance, nameof(tolerance));

            if (!TryNormalizePlanarVector(firstDirection, tolerance, out XYZ first))
            {
                return false;
            }

            if (!TryNormalizePlanarVector(secondDirection, tolerance, out XYZ second))
            {
                return false;
            }

            return Math.Abs(Cross2D(first, second)) <= tolerance;
        }

        /// <summary>
        /// Returns the perpendicular XY distance from a point to the picked dimension line.
        /// </summary>
        public static double DistanceToDimensionLine2D(QuickDimensionLineContext context, XYZ point)
        {
            if (context == null) throw new ArgumentNullException(nameof(context));
            if (!IsFinite(point)) throw new ArgumentException("Point must contain finite coordinates.", nameof(point));

            XYZ dimensionDirection = GetPlanarDirectionOrThrow(context.Direction, nameof(context));
            XYZ offset = point - context.FirstPoint;
            return Math.Abs(Cross2D(offset, dimensionDirection));
        }

        /// <summary>
        /// Projects a point to the picked dimension line in XY and checks whether it lands inside the picked span.
        /// </summary>
        public static bool TryProjectPointToPickedSpan(
            QuickDimensionLineContext context,
            XYZ point,
            double tolerance,
            out XYZ projectedPoint,
            out double parameterOnDimensionLine)
        {
            if (context == null) throw new ArgumentNullException(nameof(context));
            ValidateNonNegativeTolerance(tolerance, nameof(tolerance));

            projectedPoint = null!;
            parameterOnDimensionLine = 0.0;

            if (!IsFinite(point))
            {
                return false;
            }

            XYZ dimensionDirection = GetPlanarDirectionOrThrow(context.Direction, nameof(context));
            XYZ offset = point - context.FirstPoint;
            parameterOnDimensionLine = Dot2D(offset, dimensionDirection);

            if (!context.IsInsidePickedSpan(parameterOnDimensionLine, tolerance))
            {
                return false;
            }

            projectedPoint = context.Evaluate(parameterOnDimensionLine);
            return true;
        }

        /// <summary>
        /// Intersects a straight source segment with the picked dimension line in XY.
        /// The dimension line is accepted only within the picked span; the source is accepted only within its segment endpoints.
        /// </summary>
        public static bool TryIntersectSegmentWithDimensionLine2D(
            QuickDimensionLineContext context,
            XYZ segmentStart,
            XYZ segmentEnd,
            double tolerance,
            out XYZ hitPoint,
            out double parameterOnDimensionLine)
        {
            if (context == null) throw new ArgumentNullException(nameof(context));
            ValidateNonNegativeTolerance(tolerance, nameof(tolerance));

            hitPoint = null!;
            parameterOnDimensionLine = 0.0;

            if (!IsFinite(segmentStart) || !IsFinite(segmentEnd))
            {
                return false;
            }

            XYZ dimensionDirection = GetPlanarDirectionOrThrow(context.Direction, nameof(context));
            if (!TryGetPlanarDirection(segmentStart, segmentEnd, tolerance, out XYZ segmentDirection))
            {
                return false;
            }

            double segmentLength = GetPlanarDistance(segmentStart, segmentEnd);
            double denominator = Cross2D(dimensionDirection, segmentDirection);
            if (Math.Abs(denominator) <= tolerance)
            {
                return false;
            }

            XYZ delta = segmentStart - context.FirstPoint;
            double dimensionParameter = Cross2D(delta, segmentDirection) / denominator;
            double segmentParameter = Cross2D(delta, dimensionDirection) / denominator;

            if (!context.IsInsidePickedSpan(dimensionParameter, tolerance))
            {
                return false;
            }

            if (segmentParameter < -tolerance || segmentParameter > segmentLength + tolerance)
            {
                return false;
            }

            parameterOnDimensionLine = Clamp(dimensionParameter, 0.0, context.Length);
            hitPoint = context.Evaluate(parameterOnDimensionLine);
            return true;
        }

        /// <summary>
        /// Sorts valid candidates by physical order on the picked dimension line with stable source-aware tie-breakers.
        /// </summary>
        public static IReadOnlyList<QuickDimensionCandidate> SortByDimensionParameter(IEnumerable<QuickDimensionCandidate> candidates)
        {
            if (candidates == null) throw new ArgumentNullException(nameof(candidates));

            return candidates
                .Where(candidate => candidate != null)
                .OrderBy(candidate => candidate.ParameterOnDimensionLine)
                .ThenBy(candidate => candidate.SourceType)
                .ThenBy(candidate => candidate.ElementValue)
                .ThenBy(candidate => candidate.HostElementValue ?? long.MinValue)
                .ThenBy(candidate => candidate.ReferenceStrategy)
                .ThenBy(candidate => candidate.DisplayName, StringComparer.Ordinal)
                .ToList()
                .AsReadOnly();
        }

        /// <summary>
        /// Conservatively removes only candidates that are effectively the same source/reference position.
        /// Door and Window left/right records are preserved unless all identity fields and positions match.
        /// </summary>
        public static IReadOnlyList<QuickDimensionCandidate> DeduplicateCandidates(
            IEnumerable<QuickDimensionCandidate> candidates,
            double duplicateTolerance)
        {
            ValidateNonNegativeTolerance(duplicateTolerance, nameof(duplicateTolerance));

            IReadOnlyList<QuickDimensionCandidate> sortedCandidates = SortByDimensionParameter(candidates);
            var result = new List<QuickDimensionCandidate>();

            foreach (QuickDimensionCandidate candidate in sortedCandidates)
            {
                bool duplicate = result.Any(existing => AreDuplicateCandidates(existing, candidate, duplicateTolerance));
                if (!duplicate)
                {
                    result.Add(candidate);
                }
            }

            return result.AsReadOnly();
        }

        /// <summary>
        /// Returns true only when two candidates share the same source identity and near-identical position.
        /// This intentionally avoids broad ElementId-only dedupe because openings need two references per instance.
        /// </summary>
        public static bool AreDuplicateCandidates(
            QuickDimensionCandidate first,
            QuickDimensionCandidate second,
            double duplicateTolerance)
        {
            if (first == null) throw new ArgumentNullException(nameof(first));
            if (second == null) throw new ArgumentNullException(nameof(second));
            ValidateNonNegativeTolerance(duplicateTolerance, nameof(duplicateTolerance));

            if (Math.Abs(first.ParameterOnDimensionLine - second.ParameterOnDimensionLine) > duplicateTolerance)
            {
                return false;
            }

            if (first.ElementValue != second.ElementValue)
            {
                return false;
            }

            if (first.SourceType != second.SourceType)
            {
                return false;
            }

            if (first.ReferenceStrategy != second.ReferenceStrategy)
            {
                return false;
            }

            if ((first.HostElementValue ?? long.MinValue) != (second.HostElementValue ?? long.MinValue))
            {
                return false;
            }

            if (!string.Equals(first.DisplayName, second.DisplayName, StringComparison.Ordinal))
            {
                return false;
            }

            return first.HitPoint.DistanceTo(second.HitPoint) <= duplicateTolerance;
        }

        private static bool TryNormalizePlanarVector(XYZ vector, double tolerance, out XYZ normalized)
        {
            normalized = null!;

            if (!IsFinite(vector))
            {
                return false;
            }

            double length = Math.Sqrt(vector.X * vector.X + vector.Y * vector.Y);
            if (length <= tolerance)
            {
                return false;
            }

            normalized = new XYZ(vector.X / length, vector.Y / length, 0.0);
            return true;
        }

        private static XYZ GetPlanarDirectionOrThrow(XYZ vector, string parameterName)
        {
            if (!TryNormalizePlanarVector(vector, 0.0, out XYZ normalized))
            {
                throw new ArgumentException("Vector must define a non-zero planar direction.", parameterName);
            }

            return normalized;
        }

        private static double GetPlanarDistance(XYZ firstPoint, XYZ secondPoint)
        {
            double deltaX = secondPoint.X - firstPoint.X;
            double deltaY = secondPoint.Y - firstPoint.Y;
            return Math.Sqrt(deltaX * deltaX + deltaY * deltaY);
        }

        private static double Dot2D(XYZ first, XYZ second)
        {
            return (first.X * second.X) + (first.Y * second.Y);
        }

        private static double Cross2D(XYZ first, XYZ second)
        {
            return (first.X * second.Y) - (first.Y * second.X);
        }

        private static bool IsFinite(double value)
        {
            return !double.IsNaN(value) && !double.IsInfinity(value);
        }

        private static double Clamp(double value, double minimum, double maximum)
        {
            if (value < minimum)
            {
                return minimum;
            }

            if (value > maximum)
            {
                return maximum;
            }

            return value;
        }

        private static void ValidateNonNegativeTolerance(double tolerance, string parameterName)
        {
            if (double.IsNaN(tolerance) || double.IsInfinity(tolerance) || tolerance < 0.0)
            {
                throw new ArgumentOutOfRangeException(parameterName, "Tolerance must be a finite non-negative number.");
            }
        }
    }
}
