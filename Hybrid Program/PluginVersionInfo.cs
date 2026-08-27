using System;
using Autodesk.AutoCAD.ApplicationServices;
using Autodesk.AutoCAD.EditorInput;
using Autodesk.AutoCAD.Runtime;

[assembly: CommandClass(typeof(HybridSurvey.PluginVersionCommands))]

namespace HybridSurvey
{
    public static class PluginVersionInfo
    {
        public const string ProductName = "Hybrid Manager";
        public const string VersionNumber = "1.1.1";
        public const string VersionDate = "2026-08-27";
        public const string DisplayVersion = "Version " + VersionNumber;
        public const string DisplayVersionDate = "Version date: " + VersionDate;
        public const string DisplayVersionAndDate = DisplayVersion + " | " + DisplayVersionDate;

        public static string[] BuildReportLines()
        {
            return new[]
            {
                ProductName + " " + DisplayVersionAndDate,
                "HybridProgram_One.dll " + DisplayVersionAndDate
            };
        }
    }

    public sealed class PluginVersionCommands
    {
        [CommandMethod("HYBRIDVERSION", CommandFlags.Modal | CommandFlags.Session)]
        public void ShowHybridVersion()
        {
            Editor editor = Application.DocumentManager.MdiActiveDocument == null
                ? null
                : Application.DocumentManager.MdiActiveDocument.Editor;

            if (editor == null)
            {
                Application.ShowAlertDialog(string.Join(Environment.NewLine, PluginVersionInfo.BuildReportLines()));
                return;
            }

            foreach (string line in PluginVersionInfo.BuildReportLines())
                editor.WriteMessage(Environment.NewLine + line);
        }
    }
}
