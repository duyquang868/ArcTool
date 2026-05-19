#nullable enable
using System;
using ArcTool.Core.Models;
using Autodesk.Revit.DB;

namespace ArcTool.Core.Services
{
    /// <summary>
    /// Immutable coordinate settings read from Project Information parameters.
    /// The settings travel with the RVT model and are parsed into runtime-safe enum values without changing Revit Project Units.
    /// </summary>
    /// <param name="AxisMapping">Project axis mapping used to route East/West and North/South into CoordX and CoordY.</param>
    /// <param name="ParameterUnit">Unit used when writing numeric values into AT_CoordX / AT_CoordY / AT_CoordZ.</param>
    public sealed record CoordinateProjectSettings(
        CoordAxisMapping AxisMapping,
        CoordParameterUnit ParameterUnit);

    /// <summary>
    /// Reads and registers ArcTool coordinate settings stored on the Revit Project Information element.
    /// This service avoids sidecar JSON settings so coordinate conventions stay inside the RVT model.
    /// </summary>
    public static class CoordinateProjectSettingsService
    {
        /// <summary>
        /// Default axis mapping for current ArcTool coordinate projects.
        /// </summary>
        public const CoordAxisMapping DefaultAxisMapping = CoordAxisMapping.VN2000;

        /// <summary>
        /// Default coordinate parameter unit for current ArcTool coordinate projects.
        /// </summary>
        public const CoordParameterUnit DefaultParameterUnit = CoordParameterUnit.Meters;

        /// <summary>
        /// Ensures AT_CoordAxisMapping and AT_CoordUnit exist as Project Information shared parameters and have default values.
        /// Must be called inside an active transaction by the registration command.
        /// </summary>
        /// <param name="doc">Active Revit document whose Project Information parameters are being registered.</param>
        /// <param name="group">Shared-parameter definition group used by ArcTool coordinate parameters.</param>
        public static void EnsureProjectInformationParameters(Document doc, DefinitionGroup group)
        {
            if (doc == null)
            {
                throw new ArgumentNullException(nameof(doc));
            }

            if (group == null)
            {
                throw new ArgumentNullException(nameof(group));
            }

            EnsureParameterBinding(doc, group, CoordProjectSettingParamNames.AxisMapping);
            EnsureParameterBinding(doc, group, CoordProjectSettingParamNames.Unit);
            EnsureDefaultSettingValues(doc);
        }

        /// <summary>
        /// Loads coordinate settings from Project Information, using safe defaults when parameters are blank or not yet present.
        /// This method only reads Project Information values and never modifies Revit Project Units or the setting parameters.
        /// </summary>
        /// <param name="doc">Active Revit document whose Project Information settings should be read.</param>
        /// <returns>Parsed coordinate settings for the current batch run.</returns>
        public static CoordinateProjectSettings LoadOrDefault(Document doc)
        {
            if (doc == null)
            {
                throw new ArgumentNullException(nameof(doc));
            }

            ProjectInfo projectInfo = doc.ProjectInformation
                ?? throw new InvalidOperationException("Document.ProjectInformation is not available.");

            string axisMappingKey = ReadString(projectInfo, CoordProjectSettingParamNames.AxisMapping);
            string unitKey = ReadString(projectInfo, CoordProjectSettingParamNames.Unit);

            return new CoordinateProjectSettings(
                ParseAxisMapping(axisMappingKey),
                ParseParameterUnit(unitKey));
        }

        /// <summary>
        /// Converts a runtime axis mapping enum into the stable key stored in Project Information.
        /// </summary>
        /// <param name="axisMapping">Axis mapping value to convert.</param>
        /// <returns>Stable Project Information key.</returns>
        public static string ToAxisMappingKey(CoordAxisMapping axisMapping)
        {
            return axisMapping switch
            {
                CoordAxisMapping.Standard => CoordAxisMappingKeys.Standard,
                CoordAxisMapping.VN2000 => CoordAxisMappingKeys.VN2000,
                _ => CoordAxisMappingKeys.VN2000
            };
        }

        /// <summary>
        /// Converts a runtime coordinate parameter unit enum into the stable key stored in Project Information.
        /// </summary>
        /// <param name="parameterUnit">Parameter unit value to convert.</param>
        /// <returns>Stable Project Information key.</returns>
        public static string ToParameterUnitKey(CoordParameterUnit parameterUnit)
        {
            return parameterUnit switch
            {
                CoordParameterUnit.Meters => CoordParameterUnitKeys.Meters,
                CoordParameterUnit.Millimeters => CoordParameterUnitKeys.Millimeters,
                CoordParameterUnit.Feet => CoordParameterUnitKeys.Feet,
                _ => CoordParameterUnitKeys.Meters
            };
        }

        /// <summary>
        /// Parses a Project Information axis mapping key into a runtime-safe enum value.
        /// Unknown or blank values fall back to the current ArcTool default.
        /// </summary>
        /// <param name="axisMappingKey">Stored Project Information key.</param>
        /// <returns>Parsed axis mapping value.</returns>
        public static CoordAxisMapping ParseAxisMapping(string? axisMappingKey)
        {
            if (string.Equals(axisMappingKey, CoordAxisMappingKeys.Standard, StringComparison.OrdinalIgnoreCase))
            {
                return CoordAxisMapping.Standard;
            }

            if (string.Equals(axisMappingKey, CoordAxisMappingKeys.VN2000, StringComparison.OrdinalIgnoreCase)
                || string.Equals(axisMappingKey, "VN2000", StringComparison.OrdinalIgnoreCase))
            {
                return CoordAxisMapping.VN2000;
            }

            return DefaultAxisMapping;
        }

        /// <summary>
        /// Parses a Project Information unit key into a runtime-safe enum value.
        /// Unknown or blank values fall back to the current ArcTool default.
        /// </summary>
        /// <param name="unitKey">Stored Project Information key.</param>
        /// <returns>Parsed coordinate parameter unit.</returns>
        public static CoordParameterUnit ParseParameterUnit(string? unitKey)
        {
            if (string.Equals(unitKey, CoordParameterUnitKeys.Meters, StringComparison.OrdinalIgnoreCase))
            {
                return CoordParameterUnit.Meters;
            }

            if (string.Equals(unitKey, CoordParameterUnitKeys.Millimeters, StringComparison.OrdinalIgnoreCase))
            {
                return CoordParameterUnit.Millimeters;
            }

            if (string.Equals(unitKey, CoordParameterUnitKeys.Feet, StringComparison.OrdinalIgnoreCase))
            {
                return CoordParameterUnit.Feet;
            }

            return DefaultParameterUnit;
        }

        private static void EnsureDefaultSettingValues(Document doc)
        {
            ProjectInfo projectInfo = doc.ProjectInformation
                ?? throw new InvalidOperationException("Document.ProjectInformation is not available.");

            WriteDefaultString(
                projectInfo,
                CoordProjectSettingParamNames.AxisMapping,
                ToAxisMappingKey(DefaultAxisMapping));

            WriteDefaultString(
                projectInfo,
                CoordProjectSettingParamNames.Unit,
                ToParameterUnitKey(DefaultParameterUnit));
        }

        private static string ReadString(Element element, string paramName)
        {
            Parameter? parameter = element.LookupParameter(paramName);
            if (parameter == null || parameter.StorageType != StorageType.String || !parameter.HasValue)
            {
                return string.Empty;
            }

            return parameter.AsString() ?? string.Empty;
        }

        private static void WriteDefaultString(Element element, string paramName, string defaultValue)
        {
            Parameter? parameter = element.LookupParameter(paramName);
            if (parameter == null || parameter.IsReadOnly || parameter.StorageType != StorageType.String)
            {
                return;
            }

            string currentValue = parameter.AsString() ?? string.Empty;
            if (!string.IsNullOrWhiteSpace(currentValue))
            {
                return;
            }

            parameter.Set(defaultValue);
        }

        private static void EnsureParameterBinding(Document doc, DefinitionGroup group, string paramName)
        {
            Definition definition = group.Definitions.get_Item(paramName);
            if (definition == null)
            {
                var options = new ExternalDefinitionCreationOptions(paramName, SpecTypeId.String.Text)
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

        private static bool IsAlreadyBound(Document doc, string paramName)
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
                    if (category?.BuiltInCategory == BuiltInCategory.OST_ProjectInformation)
                    {
                        return true;
                    }
                }

                return false;
            }

            return false;
        }

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

                Category targetCategory = GetProjectInformationCategory(doc);
                CategorySet mergedCategories = doc.Application.Create.NewCategorySet();
                bool hasTargetCategory = false;

                foreach (Category category in existingBinding.Categories)
                {
                    if (category == null)
                    {
                        continue;
                    }

                    mergedCategories.Insert(category);
                    if (category.BuiltInCategory == BuiltInCategory.OST_ProjectInformation)
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

        private static void RegisterNewBinding(Document doc, Definition definition, string paramName)
        {
            Category targetCategory = GetProjectInformationCategory(doc);
            CategorySet categorySet = doc.Application.Create.NewCategorySet();
            categorySet.Insert(targetCategory);

            InstanceBinding binding = doc.Application.Create.NewInstanceBinding(categorySet);
            if (!doc.ParameterBindings.Insert(definition, binding))
            {
                throw new InvalidOperationException($"Failed to bind '{paramName}' to Project Information.");
            }
        }

        private static Category GetProjectInformationCategory(Document doc)
        {
            return doc.Settings.Categories.get_Item(BuiltInCategory.OST_ProjectInformation)
                ?? throw new InvalidOperationException("Could not resolve the Project Information category.");
        }
    }
}
