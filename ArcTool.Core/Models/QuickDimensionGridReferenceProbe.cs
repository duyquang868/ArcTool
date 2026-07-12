using Autodesk.Revit.DB;

namespace ArcTool.Core.Models
{
    public enum QuickDimensionGridReferenceStrategy
    {
        ElementReference,
        CurveReference
    }

    public sealed class QuickDimensionGridCandidate
    {
        public QuickDimensionGridCandidate(
            ElementId gridId,
            string gridName,
            double parameterOnDimensionLine,
            Reference elementReference,
            Reference curveReference)
        {
            GridId = gridId;
            GridName = gridName ?? string.Empty;
            ParameterOnDimensionLine = parameterOnDimensionLine;
            ElementReference = elementReference;
            CurveReference = curveReference;
        }

        public ElementId GridId { get; }
        public string GridName { get; }
        public double ParameterOnDimensionLine { get; }
        public Reference ElementReference { get; }
        public Reference CurveReference { get; }
    }

    public sealed class QuickDimensionGridStrategyProbeResult
    {
        public QuickDimensionGridStrategyProbeResult(
            QuickDimensionGridReferenceStrategy strategy,
            bool succeeded,
            int referenceCount,
            string message)
        {
            Strategy = strategy;
            Succeeded = succeeded;
            ReferenceCount = referenceCount;
            Message = message ?? string.Empty;
        }

        public QuickDimensionGridReferenceStrategy Strategy { get; }
        public bool Succeeded { get; }
        public int ReferenceCount { get; }
        public string Message { get; }
    }

    public sealed class QuickDimensionGridProbeSummary
    {
        public QuickDimensionGridProbeSummary(
            int collectedGridCount,
            int acceptedGridCount,
            int skippedArcGridCount,
            int skippedParallelGridCount,
            QuickDimensionGridStrategyProbeResult elementReferenceResult,
            QuickDimensionGridStrategyProbeResult curveReferenceResult)
        {
            CollectedGridCount = collectedGridCount;
            AcceptedGridCount = acceptedGridCount;
            SkippedArcGridCount = skippedArcGridCount;
            SkippedParallelGridCount = skippedParallelGridCount;
            ElementReferenceResult = elementReferenceResult;
            CurveReferenceResult = curveReferenceResult;
        }

        public int CollectedGridCount { get; }
        public int AcceptedGridCount { get; }
        public int SkippedArcGridCount { get; }
        public int SkippedParallelGridCount { get; }
        public QuickDimensionGridStrategyProbeResult ElementReferenceResult { get; }
        public QuickDimensionGridStrategyProbeResult CurveReferenceResult { get; }
    }
}
