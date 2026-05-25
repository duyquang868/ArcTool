#nullable enable
using System;
using System.Collections.Generic;
using System.Linq;
using ArcTool.Core.Models;
using Autodesk.Revit.DB;

namespace ArcTool.Core.Services
{
    /// <summary>
    /// Extracts the representative point for supported coordinate categories without converting units or coordinate systems.
    /// Keeping extraction separate from conversion prevents category/geometry rules from being mixed with project-location policy.
    /// </summary>
    public static class CoordinateExtractionService
    {
        /// <summary>
        /// Classifies a supported family instance and extracts its representative point in Revit internal units.
        /// Must be called only inside a valid Revit API context because it reads live element geometry state.
        /// </summary>
        /// <param name="instance">Supported family instance to classify and extract.</param>
        /// <returns>A supported coordinate result for recognized placement types, otherwise an explicit unsupported result.</returns>
        public static CoordResult Extract(FamilyInstance instance)
        {
            if (instance == null)
            {
                throw new ArgumentNullException(nameof(instance));
            }

            if (!instance.IsValidObject)
            {
                return CoordResult.Unsupported(
                    instance.Id.Value,
                    "Element is not a valid Revit object (possibly deleted or invalidated).");
            }

            BuiltInCategory? builtInCategory = GetBuiltInCategory(instance);
            if (builtInCategory == null || !CoordV1Scope.IsSupportedCategory(builtInCategory.Value))
            {
                return CoordResult.Unsupported(
                    instance.Id.Value,
                    $"Element category is not supported. Current coordinate scope: {CoordV1Scope.GetSupportedCategoryLabel()}.");
            }

            if (builtInCategory.Value == CoordV1Scope.DetailItemCategory
                && !CoordinateDetailItemRegistryService.IsRegisteredType(instance.Document, instance))
            {
                return CoordResult.Unsupported(
                    instance.Id.Value,
                    "Detail Item type is not registered in ArcTool_CoordDetailItemTypes.json.");
            }

            try
            {
                Location location = instance.Location;
                CoordColumnType columnType = ClassifyLocation(location);

                if (builtInCategory.Value == CoordV1Scope.DetailItemCategory && columnType != CoordColumnType.Vertical)
                {
                    return CoordResult.Unsupported(
                        instance.Id.Value,
                        "Detail Item coordinate extraction supports LocationPoint only.");
                }

                switch (columnType)
                {
                    case CoordColumnType.Vertical:
                    {
                        var locationPoint = (LocationPoint)location;
                        XYZ point = locationPoint.Point;
                        return CoordResult.Success(instance.Id.Value, columnType, point.X, point.Y, point.Z);
                    }

                    case CoordColumnType.Slanted:
                    {
                        var locationCurve = (LocationCurve)location;
                        Curve curve = locationCurve.Curve;
                        XYZ point = curve.GetEndPoint(0);
                        return CoordResult.Success(instance.Id.Value, columnType, point.X, point.Y, point.Z);
                    }

                    default:
                        return CoordResult.Unsupported(
                            instance.Id.Value,
                            "Location is neither LocationPoint nor LocationCurve. Cannot determine representative coordinate point.");
                }
            }
            catch (Exception ex)
            {
                return CoordResult.Unsupported(instance.Id.Value, ex.Message);
            }
        }

        /// <summary>
        /// Extracts coordinate results for one supported coordinate category in the document using quick Revit filters first.
        /// Per-element failures are returned as diagnostics so one bad element cannot block the full batch.
        /// </summary>
        /// <param name="doc">Document containing supported coordinate elements to extract.</param>
        /// <param name="triggerFilter">Single trigger filter to extract.</param>
        /// <returns>Read-only extraction results in the same order as the collected family instances.</returns>
        public static IReadOnlyList<CoordResult> ExtractAll(Document doc, CoordTriggerFilter triggerFilter)
        {
            if (doc == null)
            {
                throw new ArgumentNullException(nameof(doc));
            }

            return ExtractInstances(CollectSupportedFamilyInstances(doc, triggerFilter));
        }

        /// <summary>
        /// Extracts coordinate results for every category that has already been registered in the active document.
        /// This is the operator-facing path for Write Coordinates: registration state, not a single UI trigger, defines scope.
        /// </summary>
        /// <param name="doc">Document containing registered coordinate elements to extract.</param>
        /// <returns>Read-only extraction results across all registered coordinate scopes.</returns>
        public static IReadOnlyList<CoordResult> ExtractAllRegistered(Document doc)
        {
            if (doc == null)
            {
                throw new ArgumentNullException(nameof(doc));
            }

            var instances = new List<FamilyInstance>();
            foreach (CoordTriggerFilter triggerFilter in GetRegisteredTriggerFilters(doc))
            {
                instances.AddRange(CollectSupportedFamilyInstances(doc, triggerFilter));
            }

            return ExtractInstances(instances);
        }

        private static IReadOnlyList<CoordResult> ExtractInstances(List<FamilyInstance> instances)
        {
            var results = new List<CoordResult>(instances.Count);

            foreach (FamilyInstance instance in instances)
            {
                try
                {
                    results.Add(Extract(instance));
                }
                catch (Exception ex)
                {
                    long elementId = instance?.Id.Value ?? 0L;
                    results.Add(CoordResult.Unsupported(elementId, ex.Message));
                }
            }

            return results.AsReadOnly();
        }

        public static IReadOnlyList<CoordTriggerFilter> GetRegisteredTriggerFilters(Document doc)
        {
            if (doc == null)
            {
                throw new ArgumentNullException(nameof(doc));
            }

            var triggerFilters = new List<CoordTriggerFilter>();

            if (CoordinateParameterBindingService.IsCoordinateCategoryRegistered(doc, CoordV1Scope.TargetCategory))
            {
                triggerFilters.Add(CoordTriggerFilter.StructuralColumns);
            }

            if (CoordinateParameterBindingService.IsCoordinateCategoryRegistered(doc, CoordV1Scope.FoundationCategory))
            {
                triggerFilters.Add(CoordTriggerFilter.StructuralFoundations);
            }

            if (CoordinateParameterBindingService.IsCoordinateCategoryRegistered(doc, CoordV1Scope.DetailItemCategory))
            {
                triggerFilters.Add(CoordTriggerFilter.DetailItems);
            }

            return triggerFilters.AsReadOnly();
        }

        private static List<FamilyInstance> CollectSupportedFamilyInstances(Document doc, CoordTriggerFilter triggerFilter)
        {
            BuiltInCategory category = CoordV1Scope.GetCategory(triggerFilter);
            List<FamilyInstance> instances = new FilteredElementCollector(doc)
                .OfClass(typeof(FamilyInstance))
                .OfCategory(category)
                .Cast<FamilyInstance>()
                .ToList();

            if (triggerFilter != CoordTriggerFilter.DetailItems)
            {
                return instances;
            }

            return instances
                .Where(instance => CoordinateDetailItemRegistryService.IsRegisteredType(doc, instance))
                .ToList();
        }

        private static BuiltInCategory? GetBuiltInCategory(Element element)
        {
            if (element?.Category == null)
            {
                return null;
            }

            long categoryValue = element.Category.Id.Value;
            foreach (BuiltInCategory targetCategory in CoordV1Scope.TargetCategories)
            {
                if (categoryValue == (long)targetCategory)
                {
                    return targetCategory;
                }
            }

            return null;
        }

        /// <summary>
        /// Maps Revit location runtime type to the locked V1 classification without attempting geometry inference.
        /// A narrow classifier keeps unsupported states explicit instead of silently guessing a point from unknown geometry.
        /// </summary>
        /// <param name="location">Live Revit location object read from a family instance.</param>
        /// <returns>The V1 classification represented by the location runtime type.</returns>
        private static CoordColumnType ClassifyLocation(Location location)
        {
            return location switch
            {
                LocationPoint => CoordColumnType.Vertical,
                LocationCurve => CoordColumnType.Slanted,
                _ => CoordColumnType.Unsupported
            };
        }
    }
}
