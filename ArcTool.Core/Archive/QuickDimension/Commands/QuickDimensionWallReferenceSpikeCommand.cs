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
    /// Wall-axis spike command. Isolated smoke path: pick one straight non-curtain host Wall,
    /// pick a side (left/right), and report the two vertical edge references of the matching
    /// side face at min/max projected station along the wall axis. No dimension is created.
    /// </summary>
    [Transaction(TransactionMode.Manual)]
    public class QuickDimensionWallReferenceSpikeCommand : IExternalCommand
    {
        public Result Execute(ExternalCommandData commandData, ref string message, ElementSet elements)
        {
            UIDocument uidoc = commandData.Application.ActiveUIDocument;
            Document doc = uidoc?.Document;

            if (doc == null)
            {
                message = "No active document is available.";
                RevitTaskDialog.Show("ArcTool Error", message);
                return Result.Failed;
            }

            RevitView activeView = doc.ActiveView;
            if (activeView is not ViewPlan || activeView.IsTemplate)
            {
                RevitTaskDialog.Show(
                    "ArcTool - Quick Dimension Wall Spike",
                    "This wall spike only supports active Plan Views.");
                return Result.Cancelled;
            }

            try
            {
                Reference wallReference = uidoc.Selection.PickObject(
                    ObjectType.Element,
                    new WallSpikeSelectionFilter(),
                    "Wall spike: select ONE straight non-curtain host Wall.");

                if (wallReference == null || doc.GetElement(wallReference.ElementId) is not Wall wall || !wall.IsValidObject)
                {
                    RevitTaskDialog.Show(
                        "ArcTool - Quick Dimension Wall Spike",
                        "The selected element is not a valid host Wall.");
                    return Result.Cancelled;
                }

                XYZ sidePickPoint = uidoc.Selection.PickPoint(
                    ObjectSnapTypes.None,
                    "Wall spike: pick a point clearly on the LEFT or RIGHT side of the wall.");

                QuickDimensionWallSpikeResult result = QuickDimensionWallReferenceProbeService.RunWallReferenceProbe(
                    wall,
                    sidePickPoint);

                string logStatus = TryWriteXmlLog(doc, wall, sidePickPoint, result);

                RevitTaskDialog.Show(
                    "ArcTool - Quick Dimension Wall Spike",
                    BuildSummaryMessage(result) + Environment.NewLine + logStatus);

                return result.Succeeded ? Result.Succeeded : Result.Cancelled;
            }
            catch (Autodesk.Revit.Exceptions.OperationCanceledException)
            {
                return Result.Cancelled;
            }
            catch (Exception ex)
            {
                message = ex.Message;
                RevitTaskDialog.Show(
                    "ArcTool Error",
                    $"Quick Dimension wall reference spike failed.\n\n{ex.Message}");
                return Result.Failed;
            }
        }

        private static string TryWriteXmlLog(Document doc, Wall wall, XYZ sidePickPoint, QuickDimensionWallSpikeResult result)
        {
            try
            {
                string logPath = QuickDimensionWallSpikeXmlLogService.WriteWallSpikeLog(doc, wall, sidePickPoint, result);
                return $"XML log: {logPath}";
            }
            catch (Exception ex)
            {
                return $"XML log failed: {ex.Message}";
            }
        }

        private static string BuildSummaryMessage(QuickDimensionWallSpikeResult result)
        {
            StringBuilder sb = new StringBuilder();

            sb.AppendLine("=== QUICK DIMENSION WALL SPIKE (WALL-AXIS EDGE MODEL) ===");
            sb.AppendLine();
            sb.AppendLine($"Selected Wall: {FormatWallLabel(result)}");
            sb.AppendLine($"Wall-axis length: {FormatMillimeters(result.WallAxisLength)}");
            sb.AppendLine($"Placement side: {FormatSide(result.Side)}");
            sb.AppendLine($"Selected side face: {FormatShellLayer(result.SelectedShellLayer)}");
            sb.AppendLine($"Vertical edges on side face: {result.TotalVerticalEdgesOnSide}");
            sb.AppendLine($"Succeeded: {(result.Succeeded ? "Yes" : "No")}");
            sb.AppendLine();

            sb.AppendLine("--- WALL END ANCHORS ---");
            if (result.StartAnchor != null)
            {
                sb.AppendLine($"01. {result.StartAnchor.Label} | t={FormatMillimeters(result.StartAnchor.ParameterOnWallAxis)} | Midpoint XYZ (mm) = {FormatMidpoint(result.StartAnchor.Midpoint)}");
            }
            else
            {
                sb.AppendLine("01. Start anchor: not resolved.");
            }

            if (result.FinishAnchor != null)
            {
                sb.AppendLine($"02. {result.FinishAnchor.Label} | t={FormatMillimeters(result.FinishAnchor.ParameterOnWallAxis)} | Midpoint XYZ (mm) = {FormatMidpoint(result.FinishAnchor.Midpoint)}");
            }
            else
            {
                sb.AppendLine("02. Finish anchor: not resolved.");
            }
            sb.AppendLine();

            sb.AppendLine("--- MESSAGE ---");
            sb.AppendLine(string.IsNullOrWhiteSpace(result.Message) ? "(no message)" : result.Message);
            return sb.ToString();
        }

        private static string FormatWallLabel(QuickDimensionWallSpikeResult result)
        {
            string typeName = string.IsNullOrWhiteSpace(result.WallTypeName) ? "Unknown Type" : result.WallTypeName;
            return $"Wall {result.WallId.Value} ({typeName})";
        }

        private static string FormatSide(QuickDimensionWallSpikeSide side)
        {
            return side switch
            {
                QuickDimensionWallSpikeSide.Left => "Left (+CCW normal of wall direction)",
                QuickDimensionWallSpikeSide.Right => "Right (-CCW normal of wall direction)",
                _ => "Unspecified"
            };
        }

        private static string FormatShellLayer(ShellLayerType? shellLayer)
        {
            return shellLayer switch
            {
                ShellLayerType.Exterior => "Exterior",
                ShellLayerType.Interior => "Interior",
                _ => "None"
            };
        }

        private static string FormatMillimeters(double internalUnits)
        {
            double millimeters = UnitUtils.ConvertFromInternalUnits(internalUnits, UnitTypeId.Millimeters);
            return $"{millimeters:0.##} mm";
        }

        private static string FormatMidpoint(XYZ midpoint)
        {
            if (midpoint == null)
            {
                return "(null)";
            }

            double x = UnitUtils.ConvertFromInternalUnits(midpoint.X, UnitTypeId.Millimeters);
            double y = UnitUtils.ConvertFromInternalUnits(midpoint.Y, UnitTypeId.Millimeters);
            double z = UnitUtils.ConvertFromInternalUnits(midpoint.Z, UnitTypeId.Millimeters);
            return $"({x:0.##}, {y:0.##}, {z:0.##})";
        }

        private sealed class WallSpikeSelectionFilter : ISelectionFilter
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
