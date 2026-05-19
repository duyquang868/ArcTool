#nullable enable
using System;
using System.Collections.Generic;
using ArcTool.Core.Models;
using Autodesk.Revit.DB;

namespace ArcTool.Core.Services
{
    /// <summary>
    /// Outcome of a single element's coordinate write attempt.
    /// </summary>
    public enum CoordWriteOutcome
    {
        /// <summary>
        /// New or changed values were written successfully.
        /// </summary>
        Written,

        /// <summary>
        /// Values match existing stored values; no write was needed.
        /// </summary>
        Skipped,

        /// <summary>
        /// Element geometry could not be classified by the Coordinate V1 extraction contract.
        /// </summary>
        Unsupported,

        /// <summary>
        /// Write was attempted but at least one parameter was missing, read-only, or rejected the value.
        /// </summary>
        Failed
    }

    /// <summary>
    /// Per-element result produced by <see cref="CoordinateBatchService"/>.
    /// Immutable command and UI layers can read this result without mutating batch state.
    /// </summary>
    /// <param name="ElementId">The processed Revit ElementId.Value.</param>
    /// <param name="Outcome">The coordinate write outcome for this element.</param>
    /// <param name="DiagnosticMessage">Optional diagnostic text for unsupported or failed elements.</param>
    public sealed record CoordWriteResult(
        long ElementId,
        CoordWriteOutcome Outcome,
        string? DiagnosticMessage);

    /// <summary>
    /// Aggregate result of a full coordinate batch run.
    /// </summary>
    /// <param name="TotalCollected">Total number of structural columns returned by the extraction service.</param>
    /// <param name="WrittenCount">Number of elements whose coordinate parameters were written.</param>
    /// <param name="SkippedCount">Number of elements skipped because stored values were already current.</param>
    /// <param name="UnsupportedCount">Number of elements skipped because their geometry is unsupported.</param>
    /// <param name="FailedCount">Number of elements that could not be written due to parameter or API failures.</param>
    /// <param name="Details">Per-element detail rows for the batch run.</param>
    public sealed record CoordBatchSummary(
        int TotalCollected,
        int WrittenCount,
        int SkippedCount,
        int UnsupportedCount,
        int FailedCount,
        IReadOnlyList<CoordWriteResult> Details);

    /// <summary>
    /// Runs the Coordinate V1 batch write pipeline against structural columns.
    /// This service owns extraction, conversion, skip checking, writeback, and reporting, but not transaction or UI boundaries.
    /// </summary>
    public static class CoordinateBatchService
    {
        private const double MillimetersPerMeter = 1000.0;
        private const int ParameterDecimalPlaces = 4;

        /// <summary>
        /// Runs the full batch pipeline on all Structural Columns in the document.
        /// The caller must open the transaction before calling this method.
        /// Per-element failures are recorded in the returned summary so the batch can continue.
        /// </summary>
        /// <param name="doc">Active Revit document to process.</param>
        /// <param name="axisMapping">Explicit coordinate axis mapping selected by the command layer.</param>
        /// <param name="parameterUnit">Unit used when writing AT_CoordX / AT_CoordY / AT_CoordZ.</param>
        /// <returns>Aggregate write summary with per-element details.</returns>
        public static CoordBatchSummary RunBatch(
            Document doc,
            CoordAxisMapping axisMapping,
            CoordParameterUnit parameterUnit)
        {
            if (doc == null)
            {
                throw new ArgumentNullException(nameof(doc));
            }

            IReadOnlyList<CoordResult> extractionResults = CoordinateExtractionService.ExtractAll(doc);
            var details = new List<CoordWriteResult>(extractionResults.Count);

            foreach (CoordResult result in extractionResults)
            {
                long elementId = result.ElementId;

                try
                {
                    if (!result.IsSupported)
                    {
                        details.Add(new CoordWriteResult(
                            elementId,
                            CoordWriteOutcome.Unsupported,
                            result.DiagnosticMessage ?? "Unsupported coordinate extraction result."));
                        continue;
                    }

                    ConvertedCoordinate? converted = CoordinateConversionService.ToStorageMm(doc, result, axisMapping);
                    if (converted == null)
                    {
                        details.Add(new CoordWriteResult(
                            elementId,
                            CoordWriteOutcome.Unsupported,
                            "Coordinate conversion returned null for a supported extraction result."));
                        continue;
                    }

                    Element? element = doc.GetElement(new ElementId(elementId));
                    if (element == null || !element.IsValidObject)
                    {
                        details.Add(new CoordWriteResult(
                            elementId,
                            CoordWriteOutcome.Failed,
                            $"[{elementId}] Element could not be found or is not valid."));
                        continue;
                    }

                    double coordX = ConvertFromStorageMillimeters(converted.EastWestMm, parameterUnit);
                    double coordY = ConvertFromStorageMillimeters(converted.NorthSouthMm, parameterUnit);
                    double coordZ = ConvertFromStorageMillimeters(converted.ElevationMm, parameterUnit);

                    if (ValuesAreUnchanged(
                        element,
                        coordX,
                        coordY,
                        coordZ,
                        parameterUnit))
                    {
                        details.Add(new CoordWriteResult(elementId, CoordWriteOutcome.Skipped, null));
                        continue;
                    }

                    var failedParams = new List<string>(3);

                    if (!WriteParamValue(element, CoordParamNames.CoordX, coordX))
                    {
                        failedParams.Add(CoordParamNames.CoordX);
                    }

                    if (!WriteParamValue(element, CoordParamNames.CoordY, coordY))
                    {
                        failedParams.Add(CoordParamNames.CoordY);
                    }

                    if (!WriteParamValue(element, CoordParamNames.CoordZ, coordZ))
                    {
                        failedParams.Add(CoordParamNames.CoordZ);
                    }

                    if (failedParams.Count > 0)
                    {
                        details.Add(new CoordWriteResult(
                            elementId,
                            CoordWriteOutcome.Failed,
                            $"[{elementId}] Failed to write parameter(s): {string.Join(", ", failedParams)}."));
                        continue;
                    }

                    details.Add(new CoordWriteResult(elementId, CoordWriteOutcome.Written, null));
                }
                catch (Exception ex)
                {
                    details.Add(new CoordWriteResult(
                        elementId,
                        CoordWriteOutcome.Failed,
                        $"[{elementId}] {ex.Message}"));
                }
            }

            int writtenCount = 0;
            int skippedCount = 0;
            int unsupportedCount = 0;
            int failedCount = 0;

            foreach (CoordWriteResult detail in details)
            {
                switch (detail.Outcome)
                {
                    case CoordWriteOutcome.Written:
                        writtenCount++;
                        break;
                    case CoordWriteOutcome.Skipped:
                        skippedCount++;
                        break;
                    case CoordWriteOutcome.Unsupported:
                        unsupportedCount++;
                        break;
                    case CoordWriteOutcome.Failed:
                        failedCount++;
                        break;
                }
            }

            return new CoordBatchSummary(
                extractionResults.Count,
                writtenCount,
                skippedCount,
                unsupportedCount,
                failedCount,
                details.AsReadOnly());
        }

        /// <summary>
        /// Reads the current stored double value of a named parameter on an element.
        /// Returns <see cref="double.NaN"/> when the parameter is missing, read-only, non-double, or has no value.
        /// </summary>
        /// <param name="element">Element whose parameter should be read.</param>
        /// <param name="paramName">Parameter name to read.</param>
        /// <returns>Stored double value, or <see cref="double.NaN"/> when the value cannot be read safely.</returns>
        private static double ReadParamValue(Element element, string paramName)
        {
            try
            {
                Parameter? parameter = element.LookupParameter(paramName);
                if (parameter == null || parameter.IsReadOnly || parameter.StorageType != StorageType.Double || !parameter.HasValue)
                {
                    return double.NaN;
                }

                return parameter.AsDouble();
            }
            catch
            {
                return double.NaN;
            }
        }

        /// <summary>
        /// Converts a storage-ready millimeter value to the selected coordinate parameter unit and rounds it for parameter writeback.
        /// </summary>
        /// <param name="valueInMillimeters">Coordinate value already rounded in millimeters by the conversion service.</param>
        /// <param name="parameterUnit">Target parameter unit selected from Project Information.</param>
        /// <returns>The coordinate value expressed in the selected parameter unit and rounded to four decimal places.</returns>
        private static double ConvertFromStorageMillimeters(double valueInMillimeters, CoordParameterUnit parameterUnit)
        {
            double convertedValue = parameterUnit switch
            {
                CoordParameterUnit.Meters => valueInMillimeters / MillimetersPerMeter,
                CoordParameterUnit.Millimeters => valueInMillimeters,
                CoordParameterUnit.Feet => UnitUtils.ConvertToInternalUnits(valueInMillimeters, UnitTypeId.Millimeters),
                _ => valueInMillimeters / MillimetersPerMeter
            };

            return RoundForParameterWrite(convertedValue);
        }

        /// <summary>
        /// Rounds the final parameter value to the operator-facing coordinate precision.
        /// </summary>
        /// <param name="value">Value already converted to the selected parameter unit.</param>
        /// <returns>The value rounded to four decimal places.</returns>
        private static double RoundForParameterWrite(double value)
        {
            return Math.Round(value, ParameterDecimalPlaces, MidpointRounding.AwayFromZero);
        }

        /// <summary>
        /// Normalizes a stored parameter value through the locked millimeter rounding policy before comparison.
        /// </summary>
        /// <param name="storedValue">Stored coordinate value in the selected parameter unit.</param>
        /// <param name="parameterUnit">Unit used by the stored parameter value.</param>
        /// <returns>The comparable value expressed in the selected parameter unit.</returns>
        private static double NormalizeStoredValue(double storedValue, CoordParameterUnit parameterUnit)
        {
            double storageMillimeters = parameterUnit switch
            {
                CoordParameterUnit.Meters => storedValue * MillimetersPerMeter,
                CoordParameterUnit.Millimeters => storedValue,
                CoordParameterUnit.Feet => UnitUtils.ConvertFromInternalUnits(storedValue, UnitTypeId.Millimeters),
                _ => storedValue * MillimetersPerMeter
            };

            double roundedMillimeters = CoordStoragePolicy.RoundForStorage(storageMillimeters);
            return ConvertFromStorageMillimeters(roundedMillimeters, parameterUnit);
        }

        /// <summary>
        /// Writes a double value to a named parameter on an element.
        /// Returns true only when the parameter exists, is writable, stores a double, and accepts the new value.
        /// </summary>
        /// <param name="element">Element whose parameter should be written.</param>
        /// <param name="paramName">Parameter name to write.</param>
        /// <param name="value">Raw numeric value to write directly into the dimensionless parameter.</param>
        /// <returns>True when the value was accepted; otherwise false.</returns>
        private static bool WriteParamValue(Element element, string paramName, double value)
        {
            try
            {
                Parameter? parameter = element.LookupParameter(paramName);
                if (parameter == null || parameter.IsReadOnly || parameter.StorageType != StorageType.Double)
                {
                    return false;
                }

                return parameter.Set(value);
            }
            catch
            {
                return false;
            }
        }

        /// <summary>
        /// Determines whether an element's current stored coordinate values already match the new computed values.
        /// All three coordinate parameters use the unit selected by AT_CoordUnit.
        /// </summary>
        /// <param name="element">Element whose coordinate parameters should be checked.</param>
        /// <param name="newX">X value expressed in the selected parameter unit.</param>
        /// <param name="newY">Y value expressed in the selected parameter unit.</param>
        /// <param name="newZ">Z value expressed in the selected parameter unit.</param>
        /// <param name="parameterUnit">Unit used for all coordinate parameter values.</param>
        /// <returns>True only when all three normalized stored values exactly match the new values.</returns>
        private static bool ValuesAreUnchanged(
            Element element,
            double newX,
            double newY,
            double newZ,
            CoordParameterUnit parameterUnit)
        {
            double storedX = ReadParamValue(element, CoordParamNames.CoordX);
            double storedY = ReadParamValue(element, CoordParamNames.CoordY);
            double storedZ = ReadParamValue(element, CoordParamNames.CoordZ);

            if (double.IsNaN(storedX) || double.IsNaN(storedY) || double.IsNaN(storedZ))
            {
                return false;
            }

            double normalizedX = NormalizeStoredValue(storedX, parameterUnit);
            double normalizedY = NormalizeStoredValue(storedY, parameterUnit);
            double normalizedZ = NormalizeStoredValue(storedZ, parameterUnit);

            return normalizedX == newX && normalizedY == newY && normalizedZ == newZ;
        }
    }
}
