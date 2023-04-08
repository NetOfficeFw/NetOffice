using System.Diagnostics;
using System.Runtime.InteropServices;
using NetOffice.Tools;
using NetOffice.WordApi.Tools;

namespace WordAddinCore
{
    [ComVisible(true)]
    [Guid("44408BE8-1C1A-4E80-93A5-3FE1B54B4384")]
    [ProgId("NetOfficeSamples.WordAddinCore")]
    public class Addin : COMAddin
    {
        public Addin()
        {
            OnStartupComplete += Addin_OnStartupComplete;
            OnDisconnection += Addin_OnDisconnection;
        }

        private void Addin_OnStartupComplete(ref Array custom)
        {
            var appName = this.Application.Name;
            var appVersion = this.Application.Version;
            var appBuild = this.Application.Build;
            Trace.WriteLine($"Addin connected to {appName} version {appVersion} ({appBuild}).");
        }

        private void Addin_OnDisconnection(ext_DisconnectMode removeMode, ref Array custom)
        {
            Trace.WriteLine($"Addin will shutdown.");
        }
    }
}