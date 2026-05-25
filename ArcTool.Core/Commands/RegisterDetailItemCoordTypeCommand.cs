#nullable enable
using System;
using System.IO;
using ArcTool.Core.Models;
using ArcTool.Core.Services;
using Autodesk.Revit.Attributes;
using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using Autodesk.Revit.UI.Selection;
using RevitTaskDialog = Autodesk.Revit.UI.TaskDialog;

namespace ArcTool.Core.Commands
{
    /// <summary>
    /// Registers one Detail Item type name for coordinate processing by asking the user to pick a representative instance.
    /// </summary>
    [Transaction(TransactionMode.Manual)]
    public class RegisterDetailItemCoordTypeCommand : IExternalCommand
    {
        /// <summary>
        /// Prompts the user to select a Detail Item instance and stores its type name in the RVT-adjacent JSON registry.
        /// </summary>
        /// <param name="commandData">Revit command context used to access selection and the active document.</param>
        /// <param name="message">Failure message returned to Revit when registration cannot complete.</param>
        /// <param name="elements">Unused element set required by the Revit external-command contract.</param>
        /// <returns>A Revit result describing whether registration succeeded, failed, or was cancelled.</returns>
        public Result Execute(ExternalCommandData commandData, ref string message, ElementSet elements)
        {
            UIDocument? uidoc = commandData.Application.ActiveUIDocument;
            if (uidoc == null)
            {
                RevitTaskDialog.Show("ArcTool", "No document is open.");
                return Result.Failed;
            }

            Document doc = uidoc.Document;
            if (string.IsNullOrWhiteSpace(doc.PathName))
            {
                RevitTaskDialog.Show(
                    "ArcTool — Detail Item Coordinates",
                    "Save the RVT file before registering Detail Item coordinate types.\n\n" +
                    "The registry JSON file is stored next to the RVT file.");
                return Result.Failed;
            }

            string? sharedParametersFilename = doc.Application.SharedParametersFilename;
            if (string.IsNullOrWhiteSpace(sharedParametersFilename) || !File.Exists(sharedParametersFilename))
            {
                RevitTaskDialog.Show(
                    "ArcTool — Detail Item Coordinates",
                    "No shared parameter file is set. Open Manage → Shared Parameters, create or select a .txt file, then run this command again.");
                message = "No shared parameter file is configured.";
                return Result.Failed;
            }

            try
            {
                Reference pickedReference = uidoc.Selection.PickObject(
                    ObjectType.Element,
                    new DetailItemSelectionFilter(),
                    "Select one Detail Item instance to register its type for ArcTool coordinate processing.");

                Element? pickedElement = doc.GetElement(pickedReference);
                if (pickedElement is not FamilyInstance detailItem)
                {
                    RevitTaskDialog.Show("ArcTool — Detail Item Coordinates", "Selected element is not a Detail Item family instance.");
                    return Result.Failed;
                }

                if (detailItem.Location is not LocationPoint)
                {
                    RevitTaskDialog.Show(
                        "ArcTool — Detail Item Coordinates",
                        "Selected Detail Item type is not supported because the selected instance does not use LocationPoint.");
                    return Result.Failed;
                }

                EnsureDetailItemCoordinateSetup(doc);
                string typeName = CoordinateDetailItemRegistryService.RegisterTypeFromInstance(doc, detailItem);
                RefreshUpdaterRegistration(doc);

                RevitTaskDialog.Show(
                    "ArcTool — Detail Item Coordinates",
                    $"Registered Detail Item type for coordinate processing:\n\n{typeName}\n\n" +
                    "AT_CoordX / AT_CoordY / AT_CoordZ are now available on Detail Items.\n" +
                    "Detail Items are now included in Write Coordinates and Auto Update.");

                return Result.Succeeded;
            }
            catch (Autodesk.Revit.Exceptions.OperationCanceledException)
            {
                return Result.Cancelled;
            }
            catch (Exception ex)
            {
                message = ex.Message;
                RevitTaskDialog.Show("ArcTool Error", $"Could not register Detail Item coordinate type.\n\n{ex.Message}");
                return Result.Failed;
            }
        }

        private static void EnsureDetailItemCoordinateSetup(Document doc)
        {
            using var tx = new Transaction(doc, "ArcTool: Register Detail Item Coordinates");
            tx.Start();

            try
            {
                DefinitionFile sharedParameterFile = doc.Application.OpenSharedParameterFile()
                    ?? throw new InvalidOperationException("Revit could not open the configured shared parameter file.");

                DefinitionGroup coordinateGroup = sharedParameterFile.Groups.get_Item(CoordGroupName.GroupName)
                    ?? sharedParameterFile.Groups.Create(CoordGroupName.GroupName);

                CoordinateParameterBindingService.EnsureCoordinateParameters(
                    doc,
                    coordinateGroup,
                    new[] { CoordV1Scope.DetailItemCategory },
                    "Detail Items");

                CoordinateProjectSettingsService.EnsureProjectInformationParameters(doc, coordinateGroup);
                WriteProjectInfoString(doc, CoordProjectSettingParamNames.TriggerFilter, CoordTriggerFilterKeys.DetailItems);

                tx.Commit();
            }
            catch
            {
                tx.RollBack();
                throw;
            }
        }

        private static void WriteProjectInfoString(Document doc, string paramName, string value)
        {
            Parameter? parameter = doc.ProjectInformation?.LookupParameter(paramName);
            if (parameter == null || parameter.IsReadOnly || parameter.StorageType != StorageType.String)
            {
                throw new InvalidOperationException($"Project Information parameter '{paramName}' is not available or is not writable.");
            }

            parameter.Set(value);
        }

        private static void RefreshUpdaterRegistration(Document doc)
        {
            AddInId? addInId = App.AddInId;
            if (addInId == null)
            {
                return;
            }

            CoordinateUpdaterService.UnregisterForDocument(doc, addInId);
            CoordinateUpdaterService.RegisterForDocument(doc, addInId);
        }

        private sealed class DetailItemSelectionFilter : ISelectionFilter
        {
            public bool AllowElement(Element elem)
            {
                return elem is FamilyInstance
                    && elem.Category?.Id.Value == (long)BuiltInCategory.OST_DetailComponents;
            }

            public bool AllowReference(Reference reference, XYZ position)
            {
                return false;
            }
        }
    }
}
