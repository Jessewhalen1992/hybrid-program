using Autodesk.AutoCAD.DatabaseServices;
using System;
using System.IO;

namespace HybridSurvey
{
    /// <summary>
    /// Reads the JTB UTM mode marker from a DXF snapshot of the active drawing.
    /// JTB stores UTM as a finite negative value on group 146 of the UCS record
    /// named "utm"; Ground mode uses the positive value.
    /// </summary>
    internal static class JtbUtmModeGuard
    {
        public static bool TryRequireUtmMode(
            Database database,
            string actionName,
            out string detail)
        {
            detail = string.Empty;
            string action = string.IsNullOrWhiteSpace(actionName)
                ? "this action"
                : actionName.Trim();
            string temporaryPath = Path.Combine(
                Path.GetTempPath(),
                "HybridSurvey-JtbUtm-" + Guid.NewGuid().ToString("N") + ".dxf");

            try
            {
                if (database == null)
                {
                    detail = "The active drawing could not be read. Open a drawing, then run " + action + " again.";
                    return false;
                }

                database.DxfOut(temporaryPath, 16, false);
                if (!JtbUtmDxfScaleReader.TryReadRawScale(temporaryPath, out var rawScale))
                {
                    detail = "JTB UTM settings were not found in this drawing. Run UTMSETTINGS, switch to UTM, then run " + action + " again.";
                    return false;
                }

                if (!JtbUtmDxfScaleReader.IsValidUtmRawScale(rawScale))
                {
                    detail = action + " requires the drawing to be in JTB UTM mode. Run UTMSETTINGS, switch to UTM, then run " + action + " again.";
                    return false;
                }

                return true;
            }
            catch (Exception ex)
            {
                string reason = string.IsNullOrWhiteSpace(ex.Message)
                    ? string.Empty
                    : " " + ex.Message.Trim();
                detail = "Could not read the JTB UTM settings for " + action + ". Run UTMSETTINGS, switch to UTM, then try again." + reason;
                return false;
            }
            finally
            {
                try
                {
                    if (File.Exists(temporaryPath)) File.Delete(temporaryPath);
                }
                catch (Exception)
                {
                    // Best-effort cleanup for a temporary DXF snapshot.
                }
            }
        }
    }
}
