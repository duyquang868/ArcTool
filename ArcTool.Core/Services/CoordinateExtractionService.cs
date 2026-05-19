#nullable enable
using System;
using System.Collections.Generic;
using System.Linq;
using ArcTool.Core.Models;
using Autodesk.Revit.DB;

namespace ArcTool.Core.Services
{
    /// <summary>
    /// Extracts the Phase B representative point for Coordinate V1 without converting units or coordinate systems.
    /// Keeping extraction separate from conversion prevents category/geometry rules from being mixed with project-location policy.
    /// </summary>
    public static class CoordinateExtractionService
    {
        /// <summary>
        /// Classifies a structural column element and extracts its representative point in Revit internal units.
        /// Must be called only inside a valid Revit API context because it reads live element geometry state.
        /// </summary>
        /// <param name="column">Structural column family instance to classify and extract.</param>
        /// <returns>A supported coordinate result for recognized V1 placement types, otherwise an explicit unsupported result.</returns>
        public static CoordResult Extract(FamilyInstance column)
        {
            if (column == null)
            {
                throw new ArgumentNullException(nameof(column));
            }

            if (!column.IsValidObject)
            {
                return CoordResult.Unsupported(
                    column.Id.Value,
                    "Element is not a valid Revit object (possibly deleted or invalidated).");
            }

            if (column.Category?.Id.Value != (long)CoordV1Scope.TargetCategory)
            {
                return CoordResult.Unsupported(
                    column.Id.Value,
                    "Element category is not OST_StructuralColumns. V1 only supports structural columns.");
            }

            try
            {
                Location location = column.Location;
                CoordColumnType columnType = ClassifyLocation(location);

                switch (columnType)
                {
                    case CoordColumnType.Vertical:
                    {
                        var locationPoint = (LocationPoint)location;
                        XYZ point = locationPoint.Point;
                        return CoordResult.Success(column.Id.Value, columnType, point.X, point.Y, point.Z);
                    }

                    case CoordColumnType.Slanted:
                    {
                        var locationCurve = (LocationCurve)location;
                        Curve curve = locationCurve.Curve;
                        XYZ point = curve.GetEndPoint(0);
                        return CoordResult.Success(column.Id.Value, columnType, point.X, point.Y, point.Z);
                    }

                    default:
                        return CoordResult.Unsupported(
                            column.Id.Value,
                            "Location is neither LocationPoint nor LocationCurve. Cannot determine column base point.");
                }
            }
            catch (Exception ex)
            {
                return CoordResult.Unsupported(column.Id.Value, ex.Message);
            }
        }

        /// <summary>
        /// Extracts coordinate results for all V1 structural columns in the document using quick Revit filters first.
        /// Per-element failures are returned as diagnostics so one bad column cannot block the full batch.
        /// </summary>
        /// <param name="doc">Document containing structural columns to extract.</param>
        /// <returns>Read-only extraction results in the same order as the collected family instances.</returns>
        public static IReadOnlyList<CoordResult> ExtractAll(Document doc)
        {
            if (doc == null)
            {
                throw new ArgumentNullException(nameof(doc));
            }

            List<FamilyInstance> columns = new FilteredElementCollector(doc)
                .OfClass(typeof(FamilyInstance))
                .OfCategory(BuiltInCategory.OST_StructuralColumns)
                .Cast<FamilyInstance>()
                .ToList();

            var results = new List<CoordResult>(columns.Count);

            foreach (FamilyInstance column in columns)
            {
                try
                {
                    results.Add(Extract(column));
                }
                catch (Exception ex)
                {
                    long elementId = column?.Id.Value ?? 0L;
                    results.Add(CoordResult.Unsupported(elementId, ex.Message));
                }
            }

            return results.AsReadOnly();
        }

        /// <summary>
        /// Maps Revit location runtime type to the locked V1 column classification without attempting geometry inference.
        /// A narrow classifier keeps unsupported states explicit instead of silently guessing a point from unknown geometry.
        /// </summary>
        /// <param name="location">Live Revit location object read from a family instance.</param>
        /// <returns>The V1 column classification represented by the location runtime type.</returns>
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
