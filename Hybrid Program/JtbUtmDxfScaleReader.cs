using System;
using System.Globalization;
using System.IO;

namespace HybridSurvey
{
    internal static class JtbUtmDxfScaleReader
    {
        private const int EntityTypeCode = 0;
        private const int NameCode = 2;
        private const int UcsElevationCode = 146;
        private const double MinimumScaleFactor = 1e-12;

        internal static bool TryReadRawScale(string dxfPath, out double rawScale)
        {
            rawScale = 0.0;
            var inUcsTable = false;
            var waitingForTableName = false;
            var waitingForUcsName = false;
            var inUtmRecord = false;

            using (var reader = new StreamReader(dxfPath))
            {
                string codeLine;
                while ((codeLine = reader.ReadLine()) != null)
                {
                    var valueLine = reader.ReadLine();
                    if (valueLine == null) break;

                    if (!int.TryParse(
                            codeLine.Trim(),
                            NumberStyles.Integer,
                            CultureInfo.InvariantCulture,
                            out var code))
                        continue;

                    var value = valueLine.Trim();
                    if (code == EntityTypeCode)
                    {
                        if (EqualsDxf(value, "TABLE"))
                        {
                            waitingForTableName = true;
                            waitingForUcsName = false;
                            inUtmRecord = false;
                            continue;
                        }
                        if (EqualsDxf(value, "ENDTAB"))
                        {
                            inUcsTable = false;
                            waitingForTableName = false;
                            waitingForUcsName = false;
                            inUtmRecord = false;
                            continue;
                        }
                        if (inUcsTable && EqualsDxf(value, "UCS"))
                        {
                            waitingForUcsName = true;
                            inUtmRecord = false;
                            continue;
                        }
                        if (inUcsTable)
                        {
                            waitingForUcsName = false;
                            inUtmRecord = false;
                        }
                        continue;
                    }

                    if (waitingForTableName && code == NameCode)
                    {
                        inUcsTable = EqualsDxf(value, "UCS");
                        waitingForTableName = false;
                        continue;
                    }
                    if (waitingForUcsName && code == NameCode)
                    {
                        inUtmRecord = EqualsDxf(value, "utm");
                        waitingForUcsName = false;
                        continue;
                    }
                    if (inUtmRecord && code == UcsElevationCode &&
                        double.TryParse(
                            value,
                            NumberStyles.Float,
                            CultureInfo.InvariantCulture,
                            out var parsed) &&
                        IsFinite(parsed))
                    {
                        rawScale = parsed;
                        return true;
                    }
                }
            }

            return false;
        }

        internal static bool IsFinite(double value)
        {
            return !double.IsNaN(value) && !double.IsInfinity(value);
        }

        internal static bool IsValidUtmRawScale(double rawScale)
        {
            return IsFinite(rawScale) &&
                   rawScale < 0.0 &&
                   Math.Abs(rawScale) > MinimumScaleFactor;
        }

        private static bool EqualsDxf(string left, string right)
        {
            return string.Equals(left, right, StringComparison.OrdinalIgnoreCase);
        }
    }
}
