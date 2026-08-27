using System;
using System.Globalization;

namespace HybridSurvey
{
    internal static class HybridValuePolicy
    {
        internal const int StorageDecimals = 3;
        internal const int DisplayDecimals = 2;

        internal static double StoreCoordinate(double value)
        {
            if (double.IsNaN(value) || double.IsInfinity(value))
                throw new ArgumentOutOfRangeException(nameof(value), "Coordinate must be finite.");

            return Math.Round(value, StorageDecimals, MidpointRounding.AwayFromZero);
        }

        internal static string DisplayCoordinate(double value)
        {
            return StoreCoordinate(value).ToString(
                "F" + DisplayDecimals,
                CultureInfo.InvariantCulture);
        }

        internal static string StorageCoordinateText(double value)
        {
            return StoreCoordinate(value).ToString(
                "F" + StorageDecimals,
                CultureInfo.InvariantCulture);
        }

        internal static string StoredCoordinateKey(double easting, double northing)
        {
            return StorageCoordinateText(easting) + "|" + StorageCoordinateText(northing);
        }

        internal static bool TryParseCoordinate(object value, out double coordinate)
        {
            coordinate = 0.0;
            if (value == null) return false;

            if (value is double typed)
            {
                if (double.IsNaN(typed) || double.IsInfinity(typed)) return false;
                coordinate = StoreCoordinate(typed);
                return true;
            }

            if (!double.TryParse(
                    Convert.ToString(value, CultureInfo.InvariantCulture),
                    NumberStyles.Float,
                    CultureInfo.InvariantCulture,
                    out var parsed) ||
                double.IsNaN(parsed) ||
                double.IsInfinity(parsed))
                return false;

            coordinate = StoreCoordinate(parsed);
            return true;
        }

        internal static bool TryResolveEditedCoordinate(
            object editedValue,
            double originalCoordinate,
            string originalDisplay,
            out double coordinate)
        {
            string entered = Convert.ToString(
                editedValue,
                CultureInfo.InvariantCulture)?.Trim() ?? string.Empty;
            if (string.Equals(
                    entered,
                    (originalDisplay ?? string.Empty).Trim(),
                    StringComparison.Ordinal))
            {
                coordinate = StoreCoordinate(originalCoordinate);
                return true;
            }

            return TryParseCoordinate(editedValue, out coordinate);
        }
    }
}
