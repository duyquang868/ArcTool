using System;
using System.IO;
using Autodesk.Revit.Attributes;
using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using ArcTool.Core.Models;
using ArcTool.Core.Services;
using RevitTaskDialog = Autodesk.Revit.UI.TaskDialog;

namespace ArcTool.Core.Commands
{
    /// <summary>
    /// Registers the ArcTool coordinate shared parameters for Structural Columns.
    /// This command exists to lock the project-side parameter contract before any coordinate extraction or updater workflow is introduced.
    /// </summary>
    [Transaction(TransactionMode.Manual)]
    public class RegisterCoordParamsCommand : IExternalCommand
    {
        /// <summary>
        /// Ensures the shared parameter definitions exist and are bound to Structural Columns.
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
                EnsureParameterBinding(doc, coordinateGroup, CoordParamNames.CoordX);
                EnsureParameterBinding(doc, coordinateGroup, CoordParamNames.CoordY);
                EnsureParameterBinding(doc, coordinateGroup, CoordParamNames.CoordZ);
                CoordinateProjectSettingsService.EnsureProjectInformationParameters(doc, coordinateGroup);
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

            RevitTaskDialog.Show(
                "ArcTool — Coordinate Parameters",
                "Coordinate parameters registered successfully.\n" +
                "AT_CoordX / AT_CoordY / AT_CoordZ are now available on Structural Columns.\n" +
                "AT_CoordAxisMapping / AT_CoordUnit are now available on Project Information.");

            return Result.Succeeded;
        }

        /// <summary>
        /// Checks whether a named shared parameter is already bound as an instance parameter to Structural Columns.
        /// Using the forward-iterator pattern avoids unsupported assumptions about BindingMap enumeration behavior in Revit.
        /// </summary>
        /// <param name="doc">Active Revit document whose binding map is being inspected.</param>
        /// <param name="paramName">Shared-parameter definition name to look for.</param>
        /// <returns>true when the parameter is already bound to Structural Columns as an instance parameter; otherwise false.</returns>
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

                foreach (Category category in instanceBinding.Categories)
                {
                    if (category == null)
                    {
                        continue;
                    }

                    if (category.BuiltInCategory == CoordV1Scope.TargetCategory)
                    {
                        return true;
                    }
                }

                return false;
            }

            return false;
        }

        /// <summary>
        /// Ensures a coordinate parameter definition exists and that its binding includes Structural Columns.
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
        /// Attempts to merge Structural Columns into an existing instance binding for the same parameter definition.
        /// ReInsert is required when the definition already exists in the binding map but the target category is missing.
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

                Category targetCategory = doc.Settings.Categories.get_Item(CoordV1Scope.TargetCategory);
                if (targetCategory == null)
                {
                    throw new InvalidOperationException("Could not resolve the Structural Columns category.");
                }

                CategorySet mergedCategories = doc.Application.Create.NewCategorySet();
                bool hasTargetCategory = false;

                foreach (Category category in existingBinding.Categories)
                {
                    if (category == null)
                    {
                        continue;
                    }

                    mergedCategories.Insert(category);
                    if (category.BuiltInCategory == CoordV1Scope.TargetCategory)
                    {
                        hasTargetCategory = true;
                    }
                }

                if (hasTargetCategory)
                {
                    return true;
                }

                mergedCategories.Insert(targetCategory);
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
        /// Registers a new instance binding for Structural Columns when no binding exists yet.
        /// Isolating first-time registration from repair logic keeps idempotence behavior explicit and easier to reason about.
        /// </summary>
        /// <param name="doc">Active Revit document that will receive the new binding.</param>
        /// <param name="definition">Shared-parameter definition to bind.</param>
        /// <param name="paramName">Parameter name used in failure messages.</param>
        private static void RegisterNewBinding(Document doc, Definition definition, string paramName)
        {
            Category targetCategory = doc.Settings.Categories.get_Item(CoordV1Scope.TargetCategory);
            if (targetCategory == null)
            {
                throw new InvalidOperationException("Could not resolve the Structural Columns category.");
            }

            CategorySet categorySet = doc.Application.Create.NewCategorySet();
            categorySet.Insert(targetCategory);

            InstanceBinding binding = doc.Application.Create.NewInstanceBinding(categorySet);
            if (!doc.ParameterBindings.Insert(definition, binding))
            {
                throw new InvalidOperationException($"Failed to bind '{paramName}' to Structural Columns.");
            }
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
