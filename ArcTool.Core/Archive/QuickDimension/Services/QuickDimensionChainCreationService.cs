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
    /// Creates one production chain dimension from an already validated read-only result.
    /// Candidate collection remains transaction-free; this service owns only Phase 3 creation.
    /// </summary>
    public static class QuickDimensionChainCreationService
    {
        public static QuickDimensionChainCreationResult CreateChainDimension(
            Document doc,
            RevitView view,
            QuickDimensionReadOnlyResult result,
            XYZ sidePickPoint)
        {
            if (doc == null) throw new ArgumentNullException(nameof(doc));
            if (view == null) throw new ArgumentNullException(nameof(view));
            if (result == null) throw new ArgumentNullException(nameof(result));
            if (sidePickPoint == null) throw new ArgumentNullException(nameof(sidePickPoint));

            if (!result.LineContext.IsWallAxis)
            {
                return QuickDimensionChainCreationResult.CreateFailed(
                    "Chain creation requires a wall-axis read-only result.");
            }

            if (!result.CanCreateChainDimension)
            {
                return QuickDimensionChainCreationResult.CreateFailed(
                    $"Chain creation requires at least two final candidates with distinct projected stations. Current count: {result.CandidateCount}.");
            }

            XYZ? sideNormal = result.LineContext.SideNormal;
            if (sideNormal == null)
            {
                return QuickDimensionChainCreationResult.CreateFailed(
                    "Chain creation requires a resolved left/right placement side.");
            }

            double sideOffset = (sidePickPoint - result.LineContext.FirstPoint).DotProduct(sideNormal);
            if (sideOffset <= result.Options.ProjectionTolerance)
            {
                return QuickDimensionChainCreationResult.CreateFailed(
                    "The side pick point does not define a positive dimension placement offset.");
            }

            double minimumStation = result.Candidates.Min(candidate => candidate.ParameterOnDimensionLine);
            double maximumStation = result.Candidates.Max(candidate => candidate.ParameterOnDimensionLine);
            if (maximumStation - minimumStation <= result.Options.DuplicateTolerance)
            {
                return QuickDimensionChainCreationResult.CreateFailed(
                    "The final candidate span is too short to create a chain dimension.");
            }

            XYZ dimensionStart = result.LineContext.Evaluate(minimumStation) + (sideNormal * sideOffset);
            XYZ dimensionEnd = result.LineContext.Evaluate(maximumStation) + (sideNormal * sideOffset);
            Line dimensionLine;
            try
            {
                dimensionLine = Line.CreateBound(dimensionStart, dimensionEnd);
            }
            catch (Exception ex)
            {
                return QuickDimensionChainCreationResult.CreateFailed(
                    $"The resolved dimension line could not be created: {ex.Message}");
            }

            ReferenceArray references = new ReferenceArray();
            foreach (QuickDimensionCandidate candidate in result.Candidates)
            {
                references.Append(candidate.Reference);
            }

            int expectedReferenceCount = result.Candidates.Count;
            if (references.Size != expectedReferenceCount)
            {
                return QuickDimensionChainCreationResult.CreateFailed(
                    $"ReferenceArray count {references.Size} does not match final candidate count {expectedReferenceCount}.");
            }

            using Transaction transaction = new Transaction(doc, "ArcTool: Quick Dimension Chain");
            transaction.Start();

            try
            {
                Dimension? dimension = doc.Create.NewDimension(view, dimensionLine, references);
                if (dimension == null)
                {
                    RollBackIfStarted(transaction);
                    return QuickDimensionChainCreationResult.CreateFailed(
                        "Revit NewDimension returned null; the creation transaction was rolled back.",
                        minimumStation,
                        maximumStation,
                        sideOffset,
                        expectedReferenceCount,
                        transaction.GetStatus());
                }

                int createdReferenceCount = dimension.References.Size;
                if (createdReferenceCount != expectedReferenceCount)
                {
                    RollBackIfStarted(transaction);
                    return QuickDimensionChainCreationResult.CreateFailed(
                        $"Created dimension reference count {createdReferenceCount} does not match final candidate count {expectedReferenceCount}; the transaction was rolled back.",
                        minimumStation,
                        maximumStation,
                        sideOffset,
                        createdReferenceCount,
                        transaction.GetStatus());
                }

                TransactionStatus commitStatus = transaction.Commit();
                if (commitStatus != TransactionStatus.Committed)
                {
                    RollBackIfStarted(transaction);
                    return QuickDimensionChainCreationResult.CreateFailed(
                        $"Dimension transaction did not commit. Revit transaction status: {commitStatus}.",
                        minimumStation,
                        maximumStation,
                        sideOffset,
                        createdReferenceCount,
                        transaction.GetStatus());
                }

                return QuickDimensionChainCreationResult.CreateSucceeded(
                    dimension.Id,
                    minimumStation,
                    maximumStation,
                    sideOffset,
                    createdReferenceCount,
                    commitStatus);
            }
            catch (Exception ex)
            {
                RollBackIfStarted(transaction);
                return QuickDimensionChainCreationResult.CreateFailed(
                    $"NewDimension failed: {ex.Message}; the creation transaction was rolled back.",
                    minimumStation,
                    maximumStation,
                    sideOffset,
                    expectedReferenceCount,
                    transaction.GetStatus());
            }
        }

        private static void RollBackIfStarted(Transaction transaction)
        {
            if (transaction.GetStatus() == TransactionStatus.Started)
            {
                transaction.RollBack();
            }
        }
    }

    public sealed class QuickDimensionChainCreationResult
    {
        private QuickDimensionChainCreationResult(
            bool succeeded,
            string message,
            ElementId? dimensionId,
            double? minimumStation,
            double? maximumStation,
            double? sideOffset,
            int referenceCount,
            TransactionStatus transactionStatus)
        {
            Succeeded = succeeded;
            Message = message;
            DimensionId = dimensionId;
            MinimumStation = minimumStation;
            MaximumStation = maximumStation;
            SideOffset = sideOffset;
            ReferenceCount = referenceCount;
            TransactionStatus = transactionStatus;
        }

        public bool Succeeded { get; }
        public string Message { get; }
        public ElementId? DimensionId { get; }
        public double? MinimumStation { get; }
        public double? MaximumStation { get; }
        public double? SideOffset { get; }
        public int ReferenceCount { get; }
        public TransactionStatus TransactionStatus { get; }

        public static QuickDimensionChainCreationResult CreateSucceeded(
            ElementId dimensionId,
            double minimumStation,
            double maximumStation,
            double sideOffset,
            int referenceCount,
            TransactionStatus transactionStatus)
        {
            return new QuickDimensionChainCreationResult(
                true,
                "Quick Dimension chain created successfully.",
                dimensionId,
                minimumStation,
                maximumStation,
                sideOffset,
                referenceCount,
                transactionStatus);
        }

        public static QuickDimensionChainCreationResult CreateFailed(
            string message,
            double? minimumStation = null,
            double? maximumStation = null,
            double? sideOffset = null,
            int referenceCount = 0,
            TransactionStatus transactionStatus = TransactionStatus.Uninitialized)
        {
            return new QuickDimensionChainCreationResult(
                false,
                message,
                null,
                minimumStation,
                maximumStation,
                sideOffset,
                referenceCount,
                transactionStatus);
        }
    }
}
