using System;
using System.IO;
using Autodesk.Revit.Attributes;
using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using ArcTool.Core.Models;
using ArcTool.Core.Services;
using ArcTool.UI;
using RevitTaskDialog = Autodesk.Revit.UI.TaskDialog;

namespace ArcTool.Core.Commands
{
    /// <summary>
    /// Registers the ArcTool coordinate shared parameters for supported 3D element categories.
    /// Detail Items use RegisterDetailItemCoordTypeCommand so the two operator pipelines stay separate.
    /// </summary>
    [Transaction(TransactionMode.Manual)]
    public class RegisterCoordParamsCommand : IExternalCommand
    {
        /// <summary>
        /// Ensures the shared parameter definitions exist and are bound to supported 3D coordinate categories.
        /// The command is intentionally idempotent so support teams can rerun it safely without creating duplicate bindings.
        /// </summary>
        /// <param name="commandData">Revit command context used to access the active document and application services.</param>
        /// <param name="message">Failure message returned to Revit when registration cannot complete safely.</param>
        /// <param name="elements">Unused selection container required by the Revit external-command contract.</param>
        /// <returns>A Revit result describing whether registration succeeded, failed, or was cancelled.</returns>
        public Result Execute(ExternalCommandData commandData, ref string message, ElementSet elements)
        {
            UIApplication uiApp = commandData.Application;
            Document doc = uiApp.ActiveUIDocument?.Document;

            if (doc == null)
            {
                message = "No active document is available.";
                return Result.Failed;
            }

            // ── PHASE E: Open settings dialog first ───────────────────────────────
            // Read current settings to pre-populate the dialog.
            string currentAxis = doc.ProjectInformation
                .LookupParameter("AT_CoordAxisMapping")?.AsString() ?? "VN-2000";
            string currentUnit = doc.ProjectInformation
                .LookupParameter("AT_CoordUnit")?.AsString() ?? "Meters";
            string currentFilter = doc.ProjectInformation
                .LookupParameter("AT_CoordTriggerFilter")?.AsString() ?? "StructuralColumns";
            var dialog = new CoordSettingsDialog(currentAxis, currentUnit, currentFilter);
            var helper = new System.Windows.Interop.WindowInteropHelper(dialog);
            helper.Owner = Autodesk.Windows.ComponentManager.ApplicationWindow;
            if (dialog.ShowDialog() != true)
            {
                return Result.Cancelled;
            }


            string sharedParametersFilename = doc.Application.SharedParametersFilename;
            if (string.IsNullOrWhiteSpace(sharedParametersFilename) || !File.Exists(sharedParametersFilename))
            {
                RevitTaskDialog.Show(
                    "ArcTool — Coordinate Parameters",
                    "No shared parameter file is set. Open Manage → Shared Parameters, create or select a .txt file, then run this command again.");
                message = "No shared parameter file is configured.";
                return Result.Failed;
            }

            using var tx = new Transaction(doc, "ArcTool: Register Coordinate Parameters");
            tx.Start();

            DefinitionFile sharedParameterFile;
            DefinitionGroup coordinateGroup;

            try
            {
                sharedParameterFile = doc.Application.OpenSharedParameterFile();
                if (sharedParameterFile == null)
                {
                    throw new InvalidOperationException("Revit could not open the configured shared parameter file.");
                }
            }
            catch (InvalidOperationException ex)
            {
                RollBackSafely(tx);
                message = ex.Message;
                return Result.Failed;
            }
            catch (IOException ex)
            {
                RollBackSafely(tx);
                message = ex.Message;
                return Result.Failed;
            }
            catch (Exception ex)
            {
                RollBackSafely(tx);
                message = ex.Message;
                return Result.Failed;
            }

            try
            {
                coordinateGroup = sharedParameterFile.Groups.get_Item(CoordGroupName.GroupName)
                    ?? sharedParameterFile.Groups.Create(CoordGroupName.GroupName);
            }
            catch (InvalidOperationException ex)
            {
                RollBackSafely(tx);
                message = ex.Message;
                return Result.Failed;
            }
            catch (IOException ex)
            {
                RollBackSafely(tx);
                message = ex.Message;
                return Result.Failed;
            }
            catch (Exception ex)
            {
                RollBackSafely(tx);
                message = ex.Message;
                return Result.Failed;
            }

            try
            {
                CoordinateParameterBindingService.EnsureCoordinateParameters(
                    doc,
                    coordinateGroup,
                    CoordV1Scope.ElementTypeCategories,
                    CoordV1Scope.GetElementTypeCategoryLabel());
                CoordinateProjectSettingsService.EnsureProjectInformationParameters(doc, coordinateGroup);
                WriteProjectInfoString(doc, "AT_CoordAxisMapping", dialog.SelectedAxisMappingKey);
                WriteProjectInfoString(doc, "AT_CoordUnit", dialog.SelectedOutputUnitKey);
                WriteProjectInfoString(doc, "AT_CoordTriggerFilter", dialog.SelectedTriggerFilterKey);
            }
            catch (InvalidOperationException ex)
            {
                RollBackSafely(tx);
                message = ex.Message;
                return Result.Failed;
            }
            catch (IOException ex)
            {
                RollBackSafely(tx);
                message = ex.Message;
                return Result.Failed;
            }
            catch (Exception ex)
            {
                RollBackSafely(tx);
                message = ex.Message;
                return Result.Failed;
            }

            try
            {
                tx.Commit();
            }
            catch (InvalidOperationException ex)
            {
                RollBackSafely(tx);
                message = ex.Message;
                return Result.Failed;
            }
            catch (IOException ex)
            {
                RollBackSafely(tx);
                message = ex.Message;
                return Result.Failed;
            }
            catch (Exception ex)
            {
                RollBackSafely(tx);
                message = ex.Message;
                return Result.Failed;
            }

            CoordinateLogService.LogSettingsChange(
                doc,
                currentAxis, dialog.SelectedAxisMappingKey,
                currentUnit, dialog.SelectedOutputUnitKey);
            RefreshUpdaterRegistration(doc);

            RevitTaskDialog.Show(
                "ArcTool — Coordinate Parameters",
                "Element coordinate registration complete.\n" +
                $"AT_CoordX / AT_CoordY / AT_CoordZ are now available on {CoordV1Scope.GetElementTypeCategoryLabel()}.\n" +
                "AT_CoordAxisMapping / AT_CoordUnit / AT_CoordTriggerFilter are now available on Project Information.");

            return Result.Succeeded;
        }


        private static void WriteProjectInfoString(Document doc, string paramName, string value)
        {
            Parameter parameter = doc.ProjectInformation?.LookupParameter(paramName);
            if (parameter == null || parameter.IsReadOnly || parameter.StorageType != StorageType.String)
            {
                return;
            }

            parameter.Set(value);
        }

        private static void RefreshUpdaterRegistration(Document doc)
        {
            AddInId addInId = App.AddInId;
            if (addInId == null)
            {
                return;
            }

            CoordinateUpdaterService.RegisterForDocument(doc, addInId);
        }

        /// <summary>
        /// Rolls back the active transaction without allowing rollback failures to hide the original registration error.
        /// Cleanup must be best-effort because the caller needs the root-cause message, not a secondary rollback exception.
        /// </summary>
        /// <param name="tx">Transaction that should be rolled back if still active.</param>
        private static void RollBackSafely(Transaction tx)
        {
            if (tx == null)
            {
                return;
            }

            try
            {
                tx.RollBack();
            }
            catch
            {
            }
        }
    }
}
