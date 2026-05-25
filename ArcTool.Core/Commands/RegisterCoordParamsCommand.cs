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

            // ── SAVE SETTINGS TO PROJECT INFORMATION ─────────────────────────────
            // Settings write requires a Transaction. Wrap in its own small transaction
            // before the shared parameter registration transaction below.
            // Saving settings and registering parameters are intentionally two separate
            // transactions so a settings-save failure does not roll back existing bindings.
            using (var txSettings = new Transaction(doc, "ArcTool: Save Coordinate Settings"))
            {
                txSettings.Start();
                try
                {
                    // Write Axis Mapping to Project Information
                    Parameter axisParam = doc.ProjectInformation
                        .LookupParameter("AT_CoordAxisMapping");
                    if (axisParam != null && !axisParam.IsReadOnly)
                    {
                        axisParam.Set(dialog.SelectedAxisMappingKey);
                    }

                    // Write Output Unit to Project Information
                    Parameter unitParam = doc.ProjectInformation
                        .LookupParameter("AT_CoordUnit");
                    if (unitParam != null && !unitParam.IsReadOnly)
                    {
                        unitParam.Set(dialog.SelectedOutputUnitKey);
                    }

                    // Write Trigger Filter to Project Information
                    Parameter triggerParam = doc.ProjectInformation
                        .LookupParameter("AT_CoordTriggerFilter");
                    if (triggerParam != null && !triggerParam.IsReadOnly)
                    {
                        triggerParam.Set(dialog.SelectedTriggerFilterKey);
                    }

                    txSettings.Commit();
                }
                catch (Exception ex)
                {
                    txSettings.RollBack();
                    // Non-fatal: log and continue with param registration.
                    // Settings can be edited again on the next run.
                    doc.Application.WriteJournalComment(
                        $"[ArcTool RegisterCoordParamsCommand] Settings save failed: {ex.Message}",
                        false);
                }
            }

            // Log the settings change.
            CoordinateLogService.LogSettingsChange(
                doc,
                currentAxis, dialog.SelectedAxisMappingKey,
                currentUnit, dialog.SelectedOutputUnitKey);
            // ── END PHASE E INSERTION ─────────────────────────────────────────────

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

            RefreshUpdaterRegistration(doc);

            RevitTaskDialog.Show(
                "ArcTool — Coordinate Parameters",
                "Element coordinate registration complete.\n" +
                $"AT_CoordX / AT_CoordY / AT_CoordZ are now available on {CoordV1Scope.GetElementTypeCategoryLabel()}.\n" +
                "AT_CoordAxisMapping / AT_CoordUnit / AT_CoordTriggerFilter are now available on Project Information.");

            return Result.Succeeded;
        }

        /// <summary>
        /// Checks whether a named shared parameter is already bound as an instance parameter to all supported coordinate categories.
        /// Using the forward-iterator pattern avoids unsupported assumptions about BindingMap enumeration behavior in Revit.
        /// </summary>
        /// <param name="doc">Active Revit document whose binding map is being inspected.</param>
        /// <param name="paramName">Shared-parameter definition name to look for.</param>
        /// <returns>true when the parameter is already bound to all supported coordinate categories as an instance parameter; otherwise false.</returns>
        private bool IsAlreadyBound(Document doc, string paramName)
        {
            DefinitionBindingMapIterator iterator = doc.ParameterBindings.ForwardIterator();
            iterator.Reset();

            while (iterator.MoveNext())
            {
                if (!string.Equals(iterator.Key?.Name, paramName, StringComparison.Ordinal))
                {
                    continue;
                }

                if (iterator.Current is not InstanceBinding instanceBinding)
                {
                    return false;
                }

                int foundCategoryCount = 0;
                foreach (Category category in instanceBinding.Categories)
                {
                    if (category == null)
                    {
                        continue;
                    }

                    if (CoordV1Scope.IsSupportedCategory(category.BuiltInCategory))
                    {
                        foundCategoryCount++;
                    }
                }

                return foundCategoryCount == CoordV1Scope.TargetCategories.Length;
            }

            return false;
        }

        /// <summary>
        /// Ensures a coordinate parameter definition exists and that its binding includes all supported coordinate categories.
        /// Merging categories through ReInsert prevents duplicate bindings while still repairing partial pre-existing bindings.
        /// </summary>
        /// <param name="doc">Active Revit document that will receive the parameter binding.</param>
        /// <param name="group">Shared-parameter definition group that stores ArcTool coordinate definitions.</param>
        /// <param name="paramName">Target coordinate parameter name.</param>
        private void EnsureParameterBinding(Document doc, DefinitionGroup group, string paramName)
        {
            Definition definition = group.Definitions.get_Item(paramName);
            if (definition == null)
            {
                var options = new ExternalDefinitionCreationOptions(paramName, SpecTypeId.Number)
                {
                    Visible = true
                };

                definition = group.Definitions.Create(options);
            }

            if (definition == null)
            {
                throw new InvalidOperationException($"Failed to create or retrieve shared parameter definition '{paramName}'.");
            }

            if (IsAlreadyBound(doc, paramName))
            {
                return;
            }

            if (TryReinsertMergedBinding(doc, definition, paramName))
            {
                return;
            }

            RegisterNewBinding(doc, definition, paramName);
        }

        /// <summary>
        /// Attempts to merge supported coordinate categories into an existing instance binding for the same parameter definition.
        /// ReInsert is required when the definition already exists in the binding map but a target category is missing.
        /// </summary>
        /// <param name="doc">Active Revit document whose binding map will be updated.</param>
        /// <param name="definition">Shared-parameter definition to repair if it already exists in the binding map.</param>
        /// <param name="paramName">Definition name used for name-based matching in the binding map.</param>
        /// <returns>true when the binding already existed or was successfully repaired; otherwise false.</returns>
        private static bool TryReinsertMergedBinding(Document doc, Definition definition, string paramName)
        {
            DefinitionBindingMapIterator iterator = doc.ParameterBindings.ForwardIterator();
            iterator.Reset();

            while (iterator.MoveNext())
            {
                if (!string.Equals(iterator.Key?.Name, paramName, StringComparison.Ordinal))
                {
                    continue;
                }

                if (iterator.Current is not InstanceBinding existingBinding)
                {
                    throw new InvalidOperationException($"Parameter '{paramName}' exists but is not an instance binding.");
                }

                CategorySet mergedCategories = doc.Application.Create.NewCategorySet();
                int foundCategoryCount = 0;

                foreach (Category category in existingBinding.Categories)
                {
                    if (category == null)
                    {
                        continue;
                    }

                    mergedCategories.Insert(category);
                    if (CoordV1Scope.IsSupportedCategory(category.BuiltInCategory))
                    {
                        foundCategoryCount++;
                    }
                }

                if (foundCategoryCount == CoordV1Scope.TargetCategories.Length)
                {
                    return true;
                }

                foreach (BuiltInCategory targetCategoryId in CoordV1Scope.TargetCategories)
                {
                    Category targetCategory = doc.Settings.Categories.get_Item(targetCategoryId);
                    if (targetCategory == null)
                    {
                        throw new InvalidOperationException($"Could not resolve supported coordinate category '{targetCategoryId}'.");
                    }

                    mergedCategories.Insert(targetCategory);
                }
                InstanceBinding mergedBinding = doc.Application.Create.NewInstanceBinding(mergedCategories);
                if (!doc.ParameterBindings.ReInsert(definition, mergedBinding))
                {
                    throw new InvalidOperationException($"Failed to update the existing binding for '{paramName}'.");
                }

                return true;
            }

            return false;
        }

        /// <summary>
        /// Registers a new instance binding for supported coordinate categories when no binding exists yet.
        /// Isolating first-time registration from repair logic keeps idempotence behavior explicit and easier to reason about.
        /// </summary>
        /// <param name="doc">Active Revit document that will receive the new binding.</param>
        /// <param name="definition">Shared-parameter definition to bind.</param>
        /// <param name="paramName">Parameter name used in failure messages.</param>
        private static void RegisterNewBinding(Document doc, Definition definition, string paramName)
        {
            CategorySet categorySet = doc.Application.Create.NewCategorySet();
            foreach (BuiltInCategory targetCategoryId in CoordV1Scope.TargetCategories)
            {
                Category targetCategory = doc.Settings.Categories.get_Item(targetCategoryId);
                if (targetCategory == null)
                {
                    throw new InvalidOperationException($"Could not resolve supported coordinate category '{targetCategoryId}'.");
                }

                categorySet.Insert(targetCategory);
            }

            InstanceBinding binding = doc.Application.Create.NewInstanceBinding(categorySet);
            if (!doc.ParameterBindings.Insert(definition, binding))
            {
                throw new InvalidOperationException($"Failed to bind '{paramName}' to {CoordV1Scope.GetSupportedCategoryLabel()}.");
            }
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
