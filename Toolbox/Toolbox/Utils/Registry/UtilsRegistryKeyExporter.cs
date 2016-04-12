using System;
using System.Diagnostics;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace NetOffice.DeveloperToolbox.Utils.Registry
{
    internal static class UtilsRegistryKeyExporter
    {
        internal static void Export(string exportPath, string registryPath)
        {
            string path = "\"" + exportPath + "\"";
            string key = "\"" + registryPath + "\"";
            Process proc = new Process();

            try
            {
                proc.StartInfo.FileName = "regedit.exe";
                proc.StartInfo.UseShellExecute = false;

                proc = Process.Start("regedit.exe", "/e " + path + " " + key);
                proc.WaitForExit();
            }
            catch (Exception)
            {
                try
                {
                    proc.Dispose();
                }
                catch
                {
                    ;
                }
            }
        }
    }
}
