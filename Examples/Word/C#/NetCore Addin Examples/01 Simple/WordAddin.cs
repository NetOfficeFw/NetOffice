using System;
using System.Diagnostics;
using System.Runtime.InteropServices;
using NetOffice.Tools;
using NetOffice.Tools.Native;
using NetOffice.WordApi;
using WordApplication = NetOffice.WordApi.Application;

namespace NetOffice.Samples.SimpleNetCoreAddin
{
    [ComVisible(true)]
    [Guid("57F2E15C-F391-446C-9B37-857EABDF1F08")]
    [ProgId("NetOfficeSimpleNetCoreAddin.WordAddin")]
    public class WordAddin : IDTExtensibility2
    {
        private WordApplication _wordApplication;

        public void OnConnection(object Application, ext_ConnectMode ConnectMode, object AddInInst, ref Array custom)
        {
            try
            {
                _wordApplication = new Application(null, Application);
            }
            catch (Exception exception)
            {
                Trace.TraceError($"Failed to create Word Application object. {exception.Message}");
            }
        }

        public void OnDisconnection(ext_DisconnectMode RemoveMode, ref Array custom)
        {
            try
            {
                _wordApplication?.Dispose();
            }
            catch (Exception exception)
            {
                Trace.TraceError($"Failed to dispose Word Application object. {exception.Message}");
            }
        }

        public void OnAddInsUpdate(ref Array custom)
        {
        }

        public void OnStartupComplete(ref Array custom)
        {
        }

        public void OnBeginShutdown(ref Array custom)
        {
        }
    }
}
