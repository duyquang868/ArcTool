#nullable enable
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text.Json;
using Autodesk.Revit.DB;

namespace ArcTool.Core.Services
{
    /// <summary>
    /// Persists the Detail Item type-name allowlist used by the coordinate backend.
    /// The JSON file lives next to the RVT file and must travel with the model when the model is moved.
    /// </summary>
    public static class CoordinateDetailItemRegistryService
    {
        private const string RegistryFileName = "ArcTool_CoordDetailItemTypes.json";

        private static readonly JsonSerializerOptions JsonOptions = new JsonSerializerOptions
        {
            WriteIndented = true,
            PropertyNameCaseInsensitive = true
        };

        /// <summary>
        /// Registers the selected Detail Item instance type name into the RVT-adjacent JSON allowlist.
        /// </summary>
        /// <param name="doc">Active Revit document used to resolve the registry path.</param>
        /// <param name="instance">Selected Detail Item family instance.</param>
        /// <returns>The registered type name.</returns>
        public static string RegisterTypeFromInstance(Document doc, FamilyInstance instance)
        {
            if (doc == null)
            {
                throw new ArgumentNullException(nameof(doc));
            }

            if (instance == null)
            {
                throw new ArgumentNullException(nameof(instance));
            }

            if (!IsDetailItem(instance))
            {
                throw new InvalidOperationException("Selected element is not a Detail Item.");
            }

            string typeName = GetTypeName(instance);
            if (string.IsNullOrWhiteSpace(typeName))
            {
                throw new InvalidOperationException("Selected Detail Item type name is blank.");
            }

            DetailItemRegistry registry = LoadRegistry(doc);
            if (!registry.TypeNames.Any(t => string.Equals(t, typeName, StringComparison.OrdinalIgnoreCase)))
            {
                registry.TypeNames.Add(typeName);
                registry.TypeNames.Sort(StringComparer.OrdinalIgnoreCase);
                SaveRegistry(doc, registry);
            }

            return typeName;
        }

        /// <summary>
        /// Returns true when the Detail Item instance type name is listed in the RVT-adjacent JSON allowlist.
        /// </summary>
        /// <param name="doc">Active Revit document used to resolve the registry path.</param>
        /// <param name="instance">Detail Item family instance to test.</param>
        /// <returns>True if the instance type is registered; otherwise false.</returns>
        public static bool IsRegisteredType(Document doc, FamilyInstance instance)
        {
            try
            {
                if (doc == null || instance == null || !IsDetailItem(instance))
                {
                    return false;
                }

                string typeName = GetTypeName(instance);
                if (string.IsNullOrWhiteSpace(typeName))
                {
                    return false;
                }

                DetailItemRegistry registry = LoadRegistry(doc);
                return registry.TypeNames.Any(t => string.Equals(t, typeName, StringComparison.OrdinalIgnoreCase));
            }
            catch
            {
                return false;
            }
        }

        /// <summary>
        /// Loads the registered Detail Item type names from the RVT-adjacent JSON allowlist.
        /// </summary>
        /// <param name="doc">Active Revit document used to resolve the registry path.</param>
        /// <returns>Registered Detail Item type names. Empty when the registry file is missing.</returns>
        public static IReadOnlyList<string> LoadRegisteredTypeNames(Document doc)
        {
            try
            {
                return LoadRegistry(doc).TypeNames.AsReadOnly();
            }
            catch
            {
                return Array.Empty<string>();
            }
        }

        private static DetailItemRegistry LoadRegistry(Document doc)
        {
            string path = GetRegistryPath(doc);
            if (!File.Exists(path))
            {
                return new DetailItemRegistry();
            }

            try
            {
                string json = File.ReadAllText(path);
                return JsonSerializer.Deserialize<DetailItemRegistry>(json, JsonOptions) ?? new DetailItemRegistry();
            }
            catch
            {
                return new DetailItemRegistry();
            }
        }

        private static void SaveRegistry(Document doc, DetailItemRegistry registry)
        {
            string path = GetRegistryPath(doc);
            string tempPath = path + ".tmp";
            string json = JsonSerializer.Serialize(registry, JsonOptions);

            File.WriteAllText(tempPath, json);
            if (File.Exists(path))
            {
                File.Replace(tempPath, path, null);
            }
            else
            {
                File.Move(tempPath, path);
            }
        }

        private static string GetRegistryPath(Document doc)
        {
            if (doc == null)
            {
                throw new ArgumentNullException(nameof(doc));
            }

            if (string.IsNullOrWhiteSpace(doc.PathName))
            {
                throw new InvalidOperationException("Save the RVT file before registering Detail Item coordinate types.");
            }

            string? directory = Path.GetDirectoryName(doc.PathName);
            if (string.IsNullOrWhiteSpace(directory))
            {
                throw new InvalidOperationException("Could not resolve the RVT folder for Detail Item coordinate registry.");
            }

            return Path.Combine(directory, RegistryFileName);
        }

        private static bool IsDetailItem(FamilyInstance instance)
        {
            return instance.Category?.Id.Value == (long)BuiltInCategory.OST_DetailComponents;
        }

        private static string GetTypeName(FamilyInstance instance)
        {
            return instance.Symbol?.Name ?? string.Empty;
        }

        private sealed class DetailItemRegistry
        {
            public List<string> TypeNames { get; set; } = new List<string>();
        }
    }
}
