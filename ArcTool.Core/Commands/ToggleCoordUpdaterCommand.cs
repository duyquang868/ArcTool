using ArcTool.Core.Services;
using Autodesk.Revit.Attributes;
using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using RevitTaskDialog = Autodesk.Revit.UI.TaskDialog;

namespace ArcTool.Core.Commands
{
    /// <summary>
    /// Toggles real-time coordinate auto-update for the active Revit document.
    /// </summary>
    [Transaction(TransactionMode.Manual)]
    public class ToggleCoordUpdaterCommand : IExternalCommand
    {
        /// <summary>
        /// Enables or disables the document-scoped coordinate updater without opening a Revit transaction.
        /// </summary>
        /// <param name="commandData">Revit command context used to access the active document.</param>
        /// <param name="message">Failure message returned to Revit when the command cannot continue.</param>
        /// <param name="elements">Unused element set required by the Revit external-command contract.</param>
        /// <returns>A Revit result describing whether the toggle command completed.</returns>
        public Result Execute(ExternalCommandData commandData, ref string message, ElementSet elements)
        {
            UIDocument uidoc = commandData.Application.ActiveUIDocument;
            if (uidoc == null)
            {
                RevitTaskDialog.Show("ArcTool", "No document is open.");
                return Result.Failed;
            }

            Document doc = uidoc.Document;
            AddInId addInId = App.AddInId;
            if (addInId == null)
            {
                RevitTaskDialog.Show("ArcTool Error",
                    "AddInId is not available. Restart Revit and try again.");
                return Result.Failed;
            }

            bool wasRegistered = CoordinateUpdaterService.IsRegisteredForDocument(doc, addInId);

            if (wasRegistered)
            {
                CoordinateUpdaterService.UnregisterForDocument(doc, addInId);
                CoordinateLogService.LogToggle(doc, isNowEnabled: false);
            }
            else
            {
                CoordinateUpdaterService.RegisterForDocument(doc, addInId);
                bool isNowRegistered = CoordinateUpdaterService.IsRegisteredForDocument(doc, addInId);
                CoordinateLogService.LogToggle(doc, isNowEnabled: isNowRegistered);
            }

            bool currentState = CoordinateUpdaterService.IsRegisteredForDocument(doc, addInId);
            string stateText = currentState
                ? "Auto-Update is now ENABLED.\n\nCoordinates will update automatically when registered coordinate elements are moved."
                : "Auto-Update is now DISABLED.\n\nUse 'Write Coordinates' to update manually.";
            RevitTaskDialog.Show("ArcTool — Coordinate Auto-Update", stateText);

            return Result.Succeeded;
        }
    }
}
