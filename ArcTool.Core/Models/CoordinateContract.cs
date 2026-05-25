#nullable enable
using System;

namespace ArcTool.Core.Models
{
    /// <summary>
    /// Centralizes the shared parameter names so later phases cannot drift on spelling.
    /// A single source of truth prevents silent binding mismatches across command, updater, UI, and schedule workflows.
    /// </summary>
    public static class CoordParamNames
    {
        /// <summary>
        /// Stores the normalized X coordinate as a raw numeric value.
        /// Keeping the persisted value unformatted preserves deterministic comparisons in later updater phases.
        /// </summary>
        public const string CoordX = "AT_CoordX";

        /// <summary>
        /// Stores the normalized Y coordinate as a raw numeric value.
        /// Keeping the persisted value unformatted preserves deterministic comparisons in later updater phases.
        /// </summary>
        public const string CoordY = "AT_CoordY";

        /// <summary>
        /// Stores the normalized Z coordinate as a raw numeric value.
        /// Keeping the persisted value unformatted preserves deterministic comparisons in later updater phases.
        /// </summary>
        public const string CoordZ = "AT_CoordZ";
    }

    /// <summary>
    /// Centralizes Project Information parameter names for project-owned coordinate settings.
    /// Storing these settings inside the RVT keeps coordinate convention data with the model instead of an external sidecar file.
    /// </summary>
    public static class CoordProjectSettingParamNames
    {
        /// <summary>
        /// Stores the project coordinate axis mapping key, such as Standard or VN-2000.
        /// </summary>
        public const string AxisMapping = "AT_CoordAxisMapping";

        /// <summary>
        /// Stores the coordinate parameter output unit key, such as Meters, Millimeters, or Feet.
        /// </summary>
        public const string Unit = "AT_CoordUnit";

        /// <summary>
        /// Stores the supported coordinate trigger filter key, such as StructuralColumns or StructuralFoundations.
        /// </summary>
        public const string TriggerFilter = "AT_CoordTriggerFilter";
    }

    /// <summary>
    /// Stable string keys persisted in AT_CoordAxisMapping.
    /// C# enum members use valid identifiers, while Project Information stores internationally recognizable display keys.
    /// </summary>
    public static class CoordAxisMappingKeys
    {
        /// <summary>
        /// Standard international axis convention: CoordX receives East/West and CoordY receives North/South.
        /// </summary>
        public const string Standard = "Standard";

        /// <summary>
        /// VN-2000 axis convention: CoordX receives North/South and CoordY receives East/West.
        /// </summary>
        public const string VN2000 = "VN-2000";
    }

    /// <summary>
    /// Stable string keys persisted in AT_CoordUnit.
    /// These keys describe the numeric unit used when writing AT_CoordX / AT_CoordY / AT_CoordZ.
    /// </summary>
    public static class CoordParameterUnitKeys
    {
        /// <summary>
        /// Coordinate parameter values are written in meters.
        /// </summary>
        public const string Meters = "Meters";

        /// <summary>
        /// Coordinate parameter values are written in millimeters.
        /// </summary>
        public const string Millimeters = "Millimeters";

        /// <summary>
        /// Coordinate parameter values are written in feet.
        /// </summary>
        public const string Feet = "Feet";
    }

    /// <summary>
    /// Stable string keys persisted in AT_CoordTriggerFilter.
    /// These keys define which supported category the coordinate batch and updater should process.
    /// </summary>
    public static class CoordTriggerFilterKeys
    {
        /// <summary>
        /// Process Structural Columns only.
        /// </summary>
        public const string StructuralColumns = "StructuralColumns";

        /// <summary>
        /// Process Structural Foundations only.
        /// </summary>
        public const string StructuralFoundations = "StructuralFoundations";

        /// <summary>
        /// Process registered Detail Item types only.
        /// </summary>
        public const string DetailItems = "DetailItems";
    }

    /// <summary>
    /// Runtime-safe trigger filter choice parsed from AT_CoordTriggerFilter.
    /// </summary>
    public enum CoordTriggerFilter
    {
        /// <summary>
        /// Process Structural Columns only.
        /// </summary>
        StructuralColumns = 0,

        /// <summary>
        /// Process Structural Foundations only.
        /// </summary>
        StructuralFoundations = 1,

        /// <summary>
        /// Process registered Detail Item types only.
        /// </summary>
        DetailItems = 2
    }

    /// <summary>
    /// Runtime-safe unit choice parsed from AT_CoordUnit.
    /// This enum is used only by ArcTool code and does not modify Revit Project Units.
    /// </summary>
    public enum CoordParameterUnit
    {
        /// <summary>
        /// Write coordinate parameter values in meters.
        /// </summary>
        Meters = 0,

        /// <summary>
        /// Write coordinate parameter values in millimeters.
        /// </summary>
        Millimeters = 1,

        /// <summary>
        /// Write coordinate parameter values in feet.
        /// </summary>
        Feet = 2
    }

    /// <summary>
    /// Centralizes the shared-parameter definition group name so the registration command and future tooling always target the same shared-parameter bucket.
    /// </summary>
    public static class CoordGroupName
    {
        /// <summary>
        /// Shared-parameter definition group used for all coordinate parameters in Phase A and later phases.
        /// A stable group name reduces support ambiguity when inspecting the shared parameter file manually.
        /// </summary>
        public const string GroupName = "ArcTool_Coordinates";
    }

    /// <summary>
    /// Classifies structural-column placement into the explicit V1 extraction rules.
    /// A dedicated domain-prefixed enum avoids namespace collisions and prevents hidden rule changes through magic strings.
    /// </summary>
    public enum CoordColumnType
    {
        /// <summary>
        /// Uses LocationPoint-based extraction.
        /// This isolates the stable vertical-column rule so later services do not infer it implicitly from geometry heuristics.
        /// </summary>
        Vertical = 0,

        /// <summary>
        /// Uses the start point of LocationCurve as the V1 slanted-column rule.
        /// Locking the rule to index 0 prevents midpoint or average-point drift across future implementations.
        /// </summary>
        Slanted = 1,

        /// <summary>
        /// Indicates that the element is in-scope by category but does not expose a recognized supported location pattern.
        /// Returning an explicit unsupported state is safer than guessing a coordinate and propagating bad data.
        /// </summary>
        Unsupported = 2
    }

    /// <summary>
    /// Defines how numeric coordinate values are normalized before persistence.
    /// Keeping this policy centralized ensures Phase C writes and Phase D updater comparisons use the same storage contract.
    /// </summary>
    public static class CoordStoragePolicy
    {
        /// <summary>
        /// Shared storage unit for persisted coordinate values.
        /// A single locked unit prevents mixed-unit parameter data across projects and commands.
        /// </summary>
        public static Autodesk.Revit.DB.ForgeTypeId StorageUnit { get; } = Autodesk.Revit.DB.UnitTypeId.Millimeters;

        /// <summary>
        /// Decimal precision applied before writing values to shared parameters.
        /// Fixed storage precision avoids noisy diffs caused by insignificant floating-point variance.
        /// </summary>
        public const int RoundingDecimalPlaces = 3;

        /// <summary>
        /// Rounds a millimeter value to the persisted storage precision.
        /// Rounding at storage time creates one canonical value so future IUpdater comparisons do not chatter on sub-threshold floating-point noise.
        /// </summary>
        /// <param name="valueInMm">Coordinate value already converted into the locked storage unit of millimeters.</param>
        /// <returns>The normalized value that should be persisted to the shared parameter.</returns>
        public static double RoundForStorage(double valueInMm)
        {
            return Math.Round(valueInMm, RoundingDecimalPlaces, MidpointRounding.AwayFromZero);
        }
    }

    /// <summary>
    /// Locks the functional boundary of Coordinate Feature V1.
    /// A centralized scope contract prevents accidental category creep beyond validated coordinate rules.
    /// </summary>
    public static class CoordV1Scope
    {
        /// <summary>
        /// Original V1 category supported by the coordinate feature.
        /// Kept for compatibility with existing code paths that still refer to the column target explicitly.
        /// </summary>
        public const Autodesk.Revit.DB.BuiltInCategory TargetCategory = Autodesk.Revit.DB.BuiltInCategory.OST_StructuralColumns;

        /// <summary>
        /// Structural Foundation category added after the column workflow was validated.
        /// The extraction rule intentionally matches the column V1 rule: LocationPoint, LocationCurve start point, otherwise unsupported.
        /// </summary>
        public const Autodesk.Revit.DB.BuiltInCategory FoundationCategory = Autodesk.Revit.DB.BuiltInCategory.OST_StructuralFoundation;

        /// <summary>
        /// Detail Item category supported only through registered type names and LocationPoint extraction.
        /// </summary>
        public const Autodesk.Revit.DB.BuiltInCategory DetailItemCategory = Autodesk.Revit.DB.BuiltInCategory.OST_DetailComponents;

        /// <summary>
        /// 3D element categories registered by the Register Element Type command.
        /// Detail Items intentionally use a separate registration pipeline.
        /// </summary>
        public static readonly Autodesk.Revit.DB.BuiltInCategory[] ElementTypeCategories =
        {
            TargetCategory,
            FoundationCategory
        };

        /// <summary>
        /// All categories currently supported by the coordinate extraction, batch write, and updater workflow.
        /// </summary>
        public static readonly Autodesk.Revit.DB.BuiltInCategory[] TargetCategories =
        {
            TargetCategory,
            FoundationCategory,
            DetailItemCategory
        };

        /// <summary>
        /// Gets the single category represented by the selected trigger filter.
        /// </summary>
        /// <param name="triggerFilter">Selected trigger filter.</param>
        /// <returns>The Revit built-in category processed by the selected filter.</returns>
        public static Autodesk.Revit.DB.BuiltInCategory GetCategory(CoordTriggerFilter triggerFilter)
        {
            return triggerFilter switch
            {
                CoordTriggerFilter.StructuralFoundations => FoundationCategory,
                CoordTriggerFilter.DetailItems => DetailItemCategory,
                _ => TargetCategory
            };
        }

        /// <summary>
        /// Explicit contract version for diagnostics and dossier tracking.
        /// Naming the version in code makes support discussions less ambiguous than relying on session notes alone.
        /// </summary>
        public const string FeatureVersion = "1.1";

        /// <summary>
        /// Diagnostic-only location subtype names recognized by V1 extraction rules.
        /// This list is intentionally not used for runtime switching so extraction logic remains type-safe in later service implementations.
        /// </summary>
        public static readonly string[] SupportedLocationTypes =
        {
            "LocationPoint",
            "LocationCurve"
        };

        /// <summary>
        /// Returns true when the category is part of the validated coordinate workflow scope.
        /// </summary>
        /// <param name="category">Built-in category to test.</param>
        /// <returns>True when the category is currently supported; otherwise false.</returns>
        public static bool IsSupportedCategory(Autodesk.Revit.DB.BuiltInCategory category)
        {
            foreach (Autodesk.Revit.DB.BuiltInCategory targetCategory in TargetCategories)
            {
                if (targetCategory == category)
                {
                    return true;
                }
            }

            return false;
        }

        /// <summary>
        /// Returns a stable support label for diagnostics and user-facing messages.
        /// </summary>
        /// <returns>Supported category names.</returns>
        public static string GetSupportedCategoryLabel()
        {
            return "Structural Columns / Structural Foundations / registered Detail Items";
        }

        /// <summary>
        /// Returns the stable label for categories handled by the Register Element Type command.
        /// </summary>
        /// <returns>Supported 3D element category names.</returns>
        public static string GetElementTypeCategoryLabel()
        {
            return "Structural Columns / Structural Foundations";
        }

        /// <summary>
        /// Returns a stable support label for the selected trigger filter.
        /// </summary>
        /// <param name="triggerFilter">Selected trigger filter.</param>
        /// <returns>Selected category name.</returns>
        public static string GetCategoryLabel(CoordTriggerFilter triggerFilter)
        {
            return triggerFilter switch
            {
                CoordTriggerFilter.StructuralFoundations => "Structural Foundations",
                CoordTriggerFilter.DetailItems => "Detail Items",
                _ => "Structural Columns"
            };
        }
    }

    /// <summary>
    /// Represents the result of a single coordinate-extraction attempt before any storage-unit conversion occurs.
    /// Keeping this model free of Revit element objects avoids leaking transient API state across service boundaries.
    /// </summary>
    public sealed class CoordResult
    {
        private CoordResult(
            long elementId,
            CoordColumnType columnType,
            bool isSupported,
            double x,
            double y,
            double z,
            string? diagnosticMessage)
        {
            ElementId = elementId;
            ColumnType = columnType;
            IsSupported = isSupported;
            X = x;
            Y = y;
            Z = z;
            DiagnosticMessage = diagnosticMessage;
        }

        /// <summary>
        /// Gets the ElementId.Value of the processed element.
        /// The contract uses long because Revit ElementId values are long and narrowing would reintroduce overflow risk.
        /// </summary>
        public long ElementId { get; }

        /// <summary>
        /// Gets the extraction rule classification that produced this result.
        /// Carrying the rule forward aids support diagnostics without forcing callers to re-evaluate geometry.
        /// </summary>
        public CoordColumnType ColumnType { get; }

        /// <summary>
        /// Gets a value indicating whether the element matched a supported V1 extraction rule.
        /// An explicit support flag makes failure handling deterministic and safer than inferring support from coordinate defaults.
        /// </summary>
        public bool IsSupported { get; }

        /// <summary>
        /// Gets the X coordinate in Revit internal units (feet).
        /// Internal-unit storage in this result preserves the extraction/conversion boundary so Phase B can own all unit conversion logic.
        /// </summary>
        public double X { get; }

        /// <summary>
        /// Gets the Y coordinate in Revit internal units (feet).
        /// Internal-unit storage in this result preserves the extraction/conversion boundary so Phase B can own all unit conversion logic.
        /// </summary>
        public double Y { get; }

        /// <summary>
        /// Gets the Z coordinate in Revit internal units (feet).
        /// Internal-unit storage in this result preserves the extraction/conversion boundary so Phase B can own all unit conversion logic.
        /// </summary>
        public double Z { get; }

        /// <summary>
        /// Gets the diagnostic reason when the element is unsupported.
        /// Carrying a human-readable reason reduces guesswork during QA and support triage.
        /// </summary>
        public string? DiagnosticMessage { get; }

        /// <summary>
        /// Creates a successful extraction result using raw Revit internal units.
        /// Conversion to storage units is intentionally deferred so Phase B can centralize all coordinate-conversion policy.
        /// </summary>
        /// <param name="elementId">ElementId.Value of the processed structural column.</param>
        /// <param name="type">Recognized V1 column type used to extract the point.</param>
        /// <param name="xFt">X coordinate in Revit internal units (feet).</param>
        /// <param name="yFt">Y coordinate in Revit internal units (feet).</param>
        /// <param name="zFt">Z coordinate in Revit internal units (feet).</param>
        /// <returns>A supported extraction result.</returns>
        public static CoordResult Success(long elementId, CoordColumnType type, double xFt, double yFt, double zFt)
        {
            if (type == CoordColumnType.Unsupported)
            {
                throw new ArgumentException("Success results require a supported column type.", nameof(type));
            }

            return new CoordResult(elementId, type, true, xFt, yFt, zFt, null);
        }

        /// <summary>
        /// Creates an unsupported extraction result with a diagnostic reason.
        /// Returning a typed unsupported result is safer than throwing for expected geometry mismatches in batch workflows.
        /// </summary>
        /// <param name="elementId">ElementId.Value of the processed structural column.</param>
        /// <param name="reason">Reason the element could not be matched to a supported V1 rule.</param>
        /// <returns>An unsupported extraction result.</returns>
        public static CoordResult Unsupported(long elementId, string reason)
        {
            if (string.IsNullOrWhiteSpace(reason))
            {
                throw new ArgumentException("Unsupported results require a diagnostic reason.", nameof(reason));
            }

            return new CoordResult(elementId, CoordColumnType.Unsupported, false, 0.0, 0.0, 0.0, reason);
        }
    }
}
