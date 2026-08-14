#nullable enable
using System;
using System.Text;
using ArcTool.Core.Archive.QuickDimension.Models;
using ArcTool.Core.Archive.QuickDimension.Services;
using Autodesk.Revit.Attributes;
using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using Autodesk.Revit.UI.Selection;
using RevitTaskDialog = Autodesk.Revit.UI.TaskDialog;
using RevitView = Autodesk.Revit.DB.View;

namespace ArcTool.Core.Archive.QuickDimension.Commands
{
    /// <summary>
    /// Phase 3 smoke command: reuses the proven read-only wall-axis engine, then creates ONE chain
    /// dimension from the final ordered candidates. It never batches walls and never edits the
    /// read-only summary command. XML audit is written before any mutation so a failed creation
    /// still leaves the ordered candidates, options, and opening diagnostics on disk.
    /// </summary>
    [Transaction(TransactionMode.Manual)]
    [Regeneration(RegenerationOption.Manual)]
    public class QuickDimensionCreateChainSmokeCommand : IExternalCommand
    {
        public Result Execute(ExternalCommandData commandData, ref string message, ElementSet elements)
        {
            UIDocument? uidoc = commandData.Application.ActiveUIDocument;
            Document? doc = uidoc?.Document;

            if (uidoc == null || doc == null)
            {
                message = "No active document is available.";
                RevitTaskDialog.Show("ArcTool Error", message);
                return Result.Failed;
            }

            RevitView activeView = doc.ActiveView;
            if (!IsSupportedPlanView(activeView))
            {
                RevitTaskDialog.Show(
                    "ArcTool - Quick Dimension Create Chain (Smoke)",
                    "Quick Dimension chain creation supports active Plan Views only.\n\n" +
                    $"Current view type: {activeView.ViewType}");
                return Result.Cancelled;
            }

            try
            {
                Reference wallReference = uidoc.Selection.PickObject(
                    ObjectType.Element,
                    new WallAxisSelectionFilter(),
                    "Quick Dimension chain: select ONE host Wall (straight, non-curtain) as the dimension axis.");

                if (wallReference == null || doc.GetElement(wallReference.ElementId) is not Wall selectedWall || !selectedWall.IsValidObject)
                {
                    RevitTaskDialog.Show(
                        "ArcTool - Quick Dimension Create Chain (Smoke)",
                        "The selected element is not a valid host Wall.");
                    return Result.Cancelled;
                }

                if (selectedWall.WallType?.Kind == WallKind.Curtain)
                {
                    RevitTaskDialog.Show(
                        "ArcTool - Quick Dimension Create Chain (Smoke)",
                        "Curtain walls are outside the Quick Dimension MVP scope.");
                    return Result.Cancelled;
                }

                if (selectedWall.Location is not LocationCurve locationCurve || locationCurve.Curve is not Line wallLine)
                {
                    RevitTaskDialog.Show(
                        "ArcTool - Quick Dimension Create Chain (Smoke)",
                        "The selected wall does not expose a straight (Line) LocationCurve. Arc/non-line host walls are excluded from MVP.");
                    return Result.Cancelled;
                }

                XYZ wallStart = wallLine.GetEndPoint(0);
                XYZ wallEnd = wallLine.GetEndPoint(1);

                XYZ sidePickPoint = uidoc.Selection.PickPoint(
                    ObjectSnapTypes.None,
                    "Quick Dimension chain: pick a point on the LEFT or RIGHT side of the wall to set the dimension placement side and offset.");

                QuickDimensionOptions options = QuickDimensionOptions.Default;
                QuickDimensionLineContext lineContext = QuickDimensionLineContext.CreateFromWallAxis(
                    selectedWall.Id,
                    wallStart,
                    wallEnd,
                    sidePickPoint,
                    options.MinimumDimensionLineLength);

                if (lineContext.SideSign == 0)
                {
                    RevitTaskDialog.Show(
                        "ArcTool - Quick Dimension Create Chain (Smoke)",
                        "The side pick point lies on the wall axis. Please pick a point clearly on the LEFT or RIGHT side of the wall.");
                    return Result.Cancelled;
                }

                QuickDimensionReadOnlyResult result = QuickDimensionReadOnlyEngine.CollectCandidates(
                    doc,
                    activeView,
                    lineContext,
                    options);

                string xmlLogStatus = TryWriteXmlLog(doc, activeView, selectedWall, result);

                if (!result.CanCreateChainDimension)
                {
                    RevitTaskDialog.Show(
                        "ArcTool - Quick Dimension Create Chain (Smoke)",
                        BuildBlockedMessage(result, selectedWall, xmlLogStatus));
                    return Result.Cancelled;
                }

                QuickDimensionChainCreationResult creationResult = QuickDimensionChainCreationService.CreateChainDimension(
                    doc,
                    activeView,
                    result,
                    sidePickPoint);

                string chainCreationAuditStatus = TryAppendChainCreationAudit(doc, xmlLogStatus, result, creationResult);

                RevitTaskDialog.Show(
                    "ArcTool - Quick Dimension Create Chain (Smoke)",
                    BuildResultMessage(result, selectedWall, creationResult, xmlLogStatus, chainCreationAuditStatus));

                return creationResult.Succeeded ? Result.Succeeded : Result.Failed;
            }
            catch (Autodesk.Revit.Exceptions.OperationCanceledException)
            {
                return Result.Cancelled;
            }
            catch (InvalidOperationException ex)
            {
                RevitTaskDialog.Show(
                    "ArcTool - Quick Dimension Create Chain (Smoke)",
                    $"Quick Dimension could not build the wall axis:\n\n{ex.Message}");
                return Result.Cancelled;
            }
            catch (Exception ex)
            {
                message = ex.Message;
                RevitTaskDialog.Show(
                    "ArcTool Error",
                    $"Quick Dimension chain creation failed.\n\n{ex.Message}");
                return Result.Failed;
            }
        }

        private static string TryWriteXmlLog(
            Document doc,
            RevitView activeView,
            Wall selectedWall,
            QuickDimensionReadOnlyResult result)
        {
            try
            {
                return QuickDimensionReadOnlyXmlLogService.WriteReadOnlySummaryLog(doc, activeView, selectedWall, result);
            }
            catch (Exception ex)
            {
                return $"XML log failed: {ex.Message}";
            }
        }

        private static string TryAppendChainCreationAudit(
            Document doc,
            string xmlLogStatus,
            QuickDimensionReadOnlyResult result,
            QuickDimensionChainCreationResult creationResult)
        {
            try
            {
                return QuickDimensionReadOnlyXmlLogService.TryAppendChainCreationAudit(doc, xmlLogStatus, result, creationResult);
            }
            catch (Exception ex)
            {
                return $"Chain creation audit failed: {ex.Message}";
            }
        }

        private static string BuildBlockedMessage(
            QuickDimensionReadOnlyResult result,
            Wall selectedWall,
            string xmlLogPath)
        {
            StringBuilder sb = new StringBuilder();
            sb.AppendLine("=== QUICK DIMENSION CHAIN CREATION (BLOCKED) ===");
            sb.AppendLine();
            sb.AppendLine($"XML log: {xmlLogPath}");
            sb.AppendLine($"Selected Wall: {BuildWallLabel(selectedWall)}");
            sb.AppendLine($"Final candidate records: {result.CandidateCount}");
            sb.AppendLine("The read-only engine did not produce at least two distinct-station candidates, so no dimension was created.");
            return sb.ToString();
        }

        private static string BuildResultMessage(
            QuickDimensionReadOnlyResult result,
            Wall selectedWall,
            QuickDimensionChainCreationResult creationResult,
            string xmlLogPath,
            string chainCreationAuditStatus)
        {
            StringBuilder sb = new StringBuilder();
            sb.AppendLine("=== QUICK DIMENSION CHAIN CREATION (SMOKE) ===");
            sb.AppendLine();
            sb.AppendLine($"XML log: {xmlLogPath}");
            sb.AppendLine($"Selected Wall: {BuildWallLabel(selectedWall)}");
            sb.AppendLine($"Creation status: {(creationResult.Succeeded ? "SUCCESS" : "FAILED")}");
            sb.AppendLine($"Creation message: {creationResult.Message}");
            sb.AppendLine($"Transaction status: {creationResult.TransactionStatus}");
            sb.AppendLine($"Audit status: {chainCreationAuditStatus}");
            sb.AppendLine();

            sb.AppendLine($"Final candidate records: {result.CandidateCount}");
            sb.AppendLine($"References used: {creationResult.ReferenceCount}");
            if (creationResult.DimensionId != null)
            {
                sb.AppendLine($"Created dimension id: {creationResult.DimensionId.Value}");
            }

            if (creationResult.MinimumStation.HasValue && creationResult.MaximumStation.HasValue)
            {
                double minMm = UnitUtils.ConvertFromInternalUnits(creationResult.MinimumStation.Value, UnitTypeId.Millimeters);
                double maxMm = UnitUtils.ConvertFromInternalUnits(creationResult.MaximumStation.Value, UnitTypeId.Millimeters);
                sb.AppendLine($"Resolved line span: {minMm:0.##} mm .. {maxMm:0.##} mm (may extend beyond raw wall axis).");
            }

            if (creationResult.SideOffset.HasValue)
            {
                double offsetMm = UnitUtils.ConvertFromInternalUnits(creationResult.SideOffset.Value, UnitTypeId.Millimeters);
                sb.AppendLine($"Placement offset from axis: {offsetMm:0.##} mm.");
            }

            return sb.ToString();
        }

        private static string BuildWallLabel(Wall wall)
        {
            string typeName = wall.WallType?.Name ?? string.Empty;
            return string.IsNullOrWhiteSpace(typeName)
                ? $"Wall {wall.Id.Value}"
                : $"Wall {wall.Id.Value} ({typeName})";
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

        private sealed class WallAxisSelectionFilter : ISelectionFilter
        {
            public bool AllowElement(Element elem)
            {
                if (elem is not Wall wall || !wall.IsValidObject)
                {
                    return false;
                }

                if (wall.WallType?.Kind == WallKind.Curtain)
                {
                    return false;
                }

                return wall.Location is LocationCurve locationCurve && locationCurve.Curve is Line;
            }

            public bool AllowReference(Reference reference, XYZ position)
            {
                return false;
            }
        }
    }
}
