using System;
using System.Collections.Generic;
using System.Linq;
using ArcTool.Core.Archive.QuickDimension.Models;
using Autodesk.Revit.DB;
using RevitView = Autodesk.Revit.DB.View;

namespace ArcTool.Core.Archive.QuickDimension.Services
{
    public static class QuickDimensionGridReferenceProbeService
    {
        private const double MinimumDimensionLineLength = 1e-6;
        private const double ParallelDotTolerance = 0.98;
        private const double ProjectionTolerance = 1e-4;

        public static QuickDimensionGridProbeSummary RunGridReferenceProbe(
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
            List<Grid> grids = new FilteredElementCollector(doc, view.Id)
                .OfClass(typeof(Grid))
                .Cast<Grid>()
                .ToList();

            List<QuickDimensionGridCandidate> candidates = new List<QuickDimensionGridCandidate>();
            int skippedArcGridCount = 0;
            int skippedParallelGridCount = 0;

            XYZ dimensionDirection = dimensionLine.Direction;
            double dimensionLength = dimensionLine.Length;

            foreach (Grid grid in grids)
            {
                if (grid?.IsValidObject != true)
                {
                    continue;
                }

                Curve gridCurve = grid.Curve;
                if (gridCurve is not Line gridLine)
                {
                    skippedArcGridCount++;
                    continue;
                }

                XYZ gridDirection = gridLine.Direction;
                if (Math.Abs(gridDirection.DotProduct(dimensionDirection)) > ParallelDotTolerance)
                {
                    skippedParallelGridCount++;
                    continue;
                }

                XYZ midpoint = (gridLine.GetEndPoint(0) + gridLine.GetEndPoint(1)) * 0.5;
                double parameter = (midpoint - firstPoint).DotProduct(dimensionDirection);
                if (parameter < -ProjectionTolerance || parameter > dimensionLength + ProjectionTolerance)
                {
                    continue;
                }

                candidates.Add(new QuickDimensionGridCandidate(
                    grid.Id,
                    grid.Name,
                    parameter,
                    new Reference(grid),
                    gridCurve.Reference));
            }

            candidates = candidates
                .OrderBy(c => c.ParameterOnDimensionLine)
                .GroupBy(c => c.GridId.Value)
                .Select(g => g.First())
                .ToList();

            QuickDimensionGridStrategyProbeResult elementReferenceResult = ProbeStrategy(
                doc,
                view,
                dimensionLine,
                candidates,
                QuickDimensionGridReferenceStrategy.ElementReference);

            QuickDimensionGridStrategyProbeResult curveReferenceResult = ProbeStrategy(
                doc,
                view,
                dimensionLine,
                candidates,
                QuickDimensionGridReferenceStrategy.CurveReference);

            return new QuickDimensionGridProbeSummary(
                grids.Count,
                candidates.Count,
                skippedArcGridCount,
                skippedParallelGridCount,
                elementReferenceResult,
                curveReferenceResult);
        }

        private static Line CreateDimensionLine(XYZ firstPoint, XYZ secondPoint)
        {
            if (firstPoint.DistanceTo(secondPoint) < MinimumDimensionLineLength)
            {
                throw new InvalidOperationException("The two picked points are too close to define a dimension line.");
            }

            return Line.CreateBound(firstPoint, secondPoint);
        }

        private static QuickDimensionGridStrategyProbeResult ProbeStrategy(
            Document doc,
            RevitView view,
            Line dimensionLine,
            IReadOnlyList<QuickDimensionGridCandidate> candidates,
            QuickDimensionGridReferenceStrategy strategy)
        {
            ReferenceArray references = new ReferenceArray();
            foreach (QuickDimensionGridCandidate candidate in candidates)
            {
                Reference reference = strategy == QuickDimensionGridReferenceStrategy.ElementReference
                    ? candidate.ElementReference
                    : candidate.CurveReference;

                if (reference != null)
                {
                    references.Append(reference);
                }
            }

            if (references.Size < 2)
            {
                return new QuickDimensionGridStrategyProbeResult(
                    strategy,
                    false,
                    references.Size,
                    "Need at least 2 valid grid references.");
            }

            using Transaction tx = new Transaction(doc, $"ArcTool: Probe {strategy}");
            tx.Start();

            try
            {
                Dimension dimension = doc.Create.NewDimension(view, dimensionLine, references);
                if (dimension == null)
                {
                    tx.RollBack();
                    return new QuickDimensionGridStrategyProbeResult(
                        strategy,
                        false,
                        references.Size,
                        "NewDimension returned null.");
                }

                tx.RollBack();
                return new QuickDimensionGridStrategyProbeResult(
                    strategy,
                    true,
                    references.Size,
                    "NewDimension accepted the references. Transaction was rolled back; no dimension was kept in the model.");
            }
            catch (Exception ex)
            {
                tx.RollBack();
                return new QuickDimensionGridStrategyProbeResult(
                    strategy,
                    false,
                    references.Size,
                    ex.Message);
            }
        }
    }
}
