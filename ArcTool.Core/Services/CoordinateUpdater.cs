#nullable enable
using System;
using System.Collections.Generic;
using ArcTool.Core.Models;
using Autodesk.Revit.DB;

namespace ArcTool.Core.Services
{
    /// <summary>
    /// Dynamic updater for Coordinate Feature V1.
    /// Processes only modified Structural Column instances and writes coordinate parameters inside Revit's active updater transaction.
    /// </summary>
    public sealed class CoordinateUpdater : IUpdater
    {
        private const double MillimetersPerMeter = 1000.0;
        private const int ParameterDecimalPlaces = 4;

        // Reentrance guard — static so it spans all Execute() calls across the lifetime of the application, not just one instance.
        private static bool _isUpdating = false;

        // UpdaterId is created once and reused. The GUID comes from CoordinateUpdaterService.UpdaterGuid to keep it in one place.
        private readonly UpdaterId _updaterId;

        public CoordinateUpdater(AddInId addInId)
        {
            _updaterId = new UpdaterId(addInId, CoordinateUpdaterService.UpdaterGuid);
        }

        public void Execute(UpdaterData data)
        {
            if (_isUpdating)
            {
                return;
            }

            _isUpdating = true;
            try
            {
                ExecuteCore(data);
            }
            finally
            {
                _isUpdating = false;
            }
        }

        public UpdaterId GetUpdaterId()
        {
            return _updaterId;
        }

        public string GetUpdaterName()
        {
            return "ArcTool Coordinate Updater";
        }

        public string GetAdditionalInformation()
        {
            return "Writes AT_CoordX / AT_CoordY / AT_CoordZ into Structural Columns " +
                   "when their location changes. Part of ArcTool V1 Coordinate Feature.";
        }

        public ChangePriority GetChangePriority()
        {
            // FreeStandingComponents (9) runs after structural walls/floors are resolved,
            // so the column location is stable before coordinate writeback.
            return ChangePriority.FreeStandingComponents;
        }

        private void ExecuteCore(UpdaterData data)
        {
            Document doc = data.GetDocument();
            if (doc == null)
            {
                return;
            }

            CoordinateProjectSettings settings = CoordinateProjectSettingsService.LoadOrDefault(doc);

            ICollection<ElementId> modifiedIds = data.GetModifiedElementIds();
            if (modifiedIds == null || modifiedIds.Count == 0)
            {
                return;
            }

            foreach (ElementId id in modifiedIds)
            {
                try
                {
                    ProcessOneElement(doc, id, settings.AxisMapping, settings.ParameterUnit);
                }
                catch (Exception ex)
                {
                    doc.Application.WriteJournalComment(
                        $"[ArcTool CoordinateUpdater] Element {id.Value}: {ex.Message}",
                        false);
                }
            }
        }

        private static void ProcessOneElement(
            Document doc,
            ElementId id,
            CoordAxisMapping axisMapping,
            CoordParameterUnit parameterUnit)
        {
            Element elem = doc.GetElement(id);
            if (elem == null || !elem.IsValidObject)
            {
                return;
            }

            if (elem is not FamilyInstance column)
            {
                return;
            }

            if (column.Category?.Id.Value != (long)CoordV1Scope.TargetCategory)
            {
                return;
            }

            CoordResult result = CoordinateExtractionService.Extract(column);
            if (!result.IsSupported)
            {
                return;
            }

            ConvertedCoordinate? converted = CoordinateConversionService.ToStorageMm(doc, result, axisMapping);
            if (converted == null)
            {
                return;
            }

            double newX = ConvertFromStorageMillimeters(converted.EastWestMm, parameterUnit);
            double newY = ConvertFromStorageMillimeters(converted.NorthSouthMm, parameterUnit);
            double newZ = ConvertFromStorageMillimeters(converted.ElevationMm, parameterUnit);

            double storedX = ReadParam(column, CoordParamNames.CoordX);
            double storedY = ReadParam(column, CoordParamNames.CoordY);
            double storedZ = ReadParam(column, CoordParamNames.CoordZ);

            if (!double.IsNaN(storedX) && !double.IsNaN(storedY) && !double.IsNaN(storedZ))
            {
                double normalizedX = NormalizeStoredValue(storedX, parameterUnit);
                double normalizedY = NormalizeStoredValue(storedY, parameterUnit);
                double normalizedZ = NormalizeStoredValue(storedZ, parameterUnit);

                if (normalizedX == newX && normalizedY == newY && normalizedZ == newZ)
                {
                    return;
                }
            }

            // Execute() runs inside Revit's already-open user transaction, so opening another Transaction here would be invalid.
            // Undo reverts both the column edit and these parameter writes as one action; this is intentional behavior for Phase D.
            WriteParam(column, CoordParamNames.CoordX, newX);
            WriteParam(column, CoordParamNames.CoordY, newY);
            WriteParam(column, CoordParamNames.CoordZ, newZ);
        }

        private static double ReadParam(Element element, string paramName)
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

        private static bool WriteParam(Element element, string paramName, double value)
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

        private static double RoundForParameterWrite(double value)
        {
            return Math.Round(value, ParameterDecimalPlaces, MidpointRounding.AwayFromZero);
        }

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
    }
}
