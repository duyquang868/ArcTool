#nullable enable
using System;
using ArcTool.Core.Models;
using Autodesk.Revit.DB;

namespace ArcTool.Core.Services
{
    /// <summary>
    /// Immutable result of a coordinate conversion after applying the locked storage precision.
    /// EastWest/NorthSouth/Elevation names are preserved before optional mapping so Phase C diagnostics can trace the source coordinate basis.
    /// </summary>
    /// <param name="EastWestMm">East/West coordinate value in the current mapping context, stored in millimeters.</param>
    /// <param name="NorthSouthMm">North/South coordinate value in the current mapping context, stored in millimeters.</param>
    /// <param name="ElevationMm">Elevation coordinate value in the current mapping context, stored in millimeters.</param>
    public sealed record ConvertedCoordinate(
        double EastWestMm,
        double NorthSouthMm,
        double ElevationMm);

    /// <summary>
    /// Named axis mapping rules keep project coordinate conventions explicit instead of burying an X/Y swap in conversion math.
    /// Add a new enum value only after a real internationally recognized project convention is confirmed.
    /// </summary>
    public enum CoordAxisMapping
    {
        /// <summary>
        /// Standard mapping: East/West becomes X, North/South becomes Y, and Elevation becomes Z.
        /// </summary>
        Standard,

        /// <summary>
        /// VN-2000 mapping: North/South becomes X, East/West becomes Y, and Elevation remains Z.
        /// The persisted Project Information key is "VN-2000" because C# enum identifiers cannot contain hyphens.
        /// </summary>
        VN2000
    }

    /// <summary>
    /// Converts extracted Revit model-space points into storage-ready coordinate values without reading elements or writing parameters.
    /// Keeping this service stateless lets Phase C choose batch error policy without coupling it to conversion mechanics.
    /// </summary>
    public static class CoordinateConversionService
    {
        /// <summary>
        /// Converts a model-space point through the active project location so shared-coordinate policy is centralized in one place.
        /// The result is immediately normalized to the locked storage unit and rounding policy to prevent future updater chatter.
        /// </summary>
        /// <param name="doc">Document whose active project location defines the shared coordinate context.</param>
        /// <param name="internalFeet">Model-space point in Revit internal units.</param>
        /// <returns>Shared-coordinate values converted to millimeters and rounded for storage.</returns>
        public static ConvertedCoordinate ToSharedMm(Document doc, XYZ internalFeet)
        {
            if (doc == null)
            {
                throw new ArgumentNullException(nameof(doc));
            }

            if (internalFeet == null)
            {
                throw new ArgumentNullException(nameof(internalFeet));
            }

            ProjectLocation location = doc.ActiveProjectLocation;
            if (location == null)
            {
                throw new InvalidOperationException(
                    "doc.ActiveProjectLocation is null. The document may not be fully loaded.");
            }

            ProjectPosition pos = location.GetProjectPosition(internalFeet);

            double ewMm = UnitUtils.ConvertFromInternalUnits(pos.EastWest, UnitTypeId.Millimeters);
            double nsMm = UnitUtils.ConvertFromInternalUnits(pos.NorthSouth, UnitTypeId.Millimeters);
            double elMm = UnitUtils.ConvertFromInternalUnits(pos.Elevation, UnitTypeId.Millimeters);

            ewMm = CoordStoragePolicy.RoundForStorage(ewMm);
            nsMm = CoordStoragePolicy.RoundForStorage(nsMm);
            elMm = CoordStoragePolicy.RoundForStorage(elMm);

            return new ConvertedCoordinate(ewMm, nsMm, elMm);
        }

        /// <summary>
        /// Applies a named axis rule after shared-coordinate conversion so regional conventions remain visible at the call site.
        /// This is the intended extension point for VN-2000/local-grid projects rather than changing core conversion math.
        /// </summary>
        /// <param name="raw">Converted coordinate before project-specific axis remapping.</param>
        /// <param name="mapping">Named mapping rule to apply.</param>
        /// <returns>A converted coordinate with values arranged according to the selected mapping rule.</returns>
        public static ConvertedCoordinate ApplyAxisMapping(ConvertedCoordinate raw, CoordAxisMapping mapping)
        {
            if (raw == null)
            {
                throw new ArgumentNullException(nameof(raw));
            }

            return mapping switch
            {
                CoordAxisMapping.Standard => new ConvertedCoordinate(
                    raw.EastWestMm,
                    raw.NorthSouthMm,
                    raw.ElevationMm),

                CoordAxisMapping.VN2000 => new ConvertedCoordinate(
                    raw.NorthSouthMm,
                    raw.EastWestMm,
                    raw.ElevationMm),

                _ => throw new ArgumentOutOfRangeException(
                    nameof(mapping),
                    mapping,
                    "Unrecognized CoordAxisMapping value.")
            };
        }

        /// <summary>
        /// Converts a supported extraction result to storage-ready millimeter values while preserving unsupported results as null.
        /// Returning null keeps expected unsupported geometry separate from project-location failures, which should still propagate.
        /// </summary>
        /// <param name="doc">Document whose active project location defines the shared coordinate context.</param>
        /// <param name="coordResult">Raw extraction result produced by CoordinateExtractionService.</param>
        /// <param name="axisMapping">Explicit project axis convention to apply after shared-coordinate conversion.</param>
        /// <returns>Storage-ready coordinate values, or null when the extraction result is unsupported.</returns>
        public static ConvertedCoordinate? ToStorageMm(
            Document doc,
            CoordResult coordResult,
            CoordAxisMapping axisMapping = CoordAxisMapping.Standard)
        {
            if (doc == null)
            {
                throw new ArgumentNullException(nameof(doc));
            }

            if (coordResult == null)
            {
                throw new ArgumentNullException(nameof(coordResult));
            }

            if (!coordResult.IsSupported)
            {
                return null;
            }

            var point = new XYZ(coordResult.X, coordResult.Y, coordResult.Z);
            ConvertedCoordinate raw = ToSharedMm(doc, point);
            return ApplyAxisMapping(raw, axisMapping);
        }
    }
}
