
using NetOffice.PowerPointApi.Tools;
using System.Diagnostics;
using System.Runtime.InteropServices;

[ComVisible(true)]
[Guid("AB6D3A7D-33CF-4197-91D9-9D6B984DDDB1")]
[ProgId("NetOffice.DevToolsAddin")]
public class Addin : COMAddin
{
    public Addin()
    {
        this.OnConnection += Addin_OnConnection;
        this.OnStartupComplete += Addin_OnStartupComplete;
    }

    private void Addin_OnConnection(object application, NetOffice.Tools.ext_ConnectMode connectMode, object addInInst, ref Array custom)
    {
        Trace.WriteLine($"Addin connected to application. Mode: {connectMode}");
    }

    private void Addin_OnStartupComplete(ref Array custom)
    {
        Trace.WriteLine($"Addin startup completed.");
    }
}
