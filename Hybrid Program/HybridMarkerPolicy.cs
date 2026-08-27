using System;

namespace HybridSurvey
{
    internal static class HybridMarkerPolicy
    {
        internal static string DesiredMarkerType(string pointType, bool markerUpdatesEnabled)
        {
            if (!markerUpdatesEnabled) return string.Empty;

            string normalized = (pointType ?? string.Empty).Trim().ToUpperInvariant();
            return normalized == "XC" || normalized == "EC"
                ? normalized
                : string.Empty;
        }
    }
}
