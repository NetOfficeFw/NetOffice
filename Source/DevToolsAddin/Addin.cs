
using Microsoft.AspNetCore.Builder;
using Microsoft.AspNetCore.Components;
using Microsoft.AspNetCore.Hosting;
using Microsoft.AspNetCore.Http;
using Microsoft.AspNetCore.Mvc;
using NetOffice.PowerPointApi;
using NetOffice.PowerPointApi.Events;
using NetOffice.PowerPointApi.Tools;
using System.Diagnostics;
using System.Runtime.InteropServices;
using System.Windows.Threading;

[ComVisible(true)]
[Guid("AB6D3A7D-33CF-4197-91D9-9D6B984DDDB1")]
[ProgId("NetOffice.DevToolsAddin")]
public class Addin : COMAddin
{
    private Task webTask;
    private WebApplication webApplication;

    public Addin()
    {
        this.OnConnection += Addin_OnConnection;
        this.OnStartupComplete += Addin_OnStartupComplete;
        this.OnDisconnection += Addin_OnDisconnection;
    }

    private void Addin_OnConnection(object application, NetOffice.Tools.ext_ConnectMode connectMode, object addInInst, ref Array custom)
    {
        Trace.WriteLine($"Addin connected to application. Mode: {connectMode}");
    }

    private void Addin_OnStartupComplete(ref Array custom)
    {
        Trace.WriteLine($"Addin startup completed.");

        var sync = System.Windows.Threading.Dispatcher.CurrentDispatcher;

        var options = new WebApplicationOptions();
        var builder = WebApplication.CreateBuilder(options);
        //builder.WebHost.UseUrls("http://localhost:53080");
        var app = builder.Build();

        app.MapGet("/", () => "Hello World!");

        app.MapGet("/json/version", () =>
        {
            var metadata = new BrowserVersionMetadata
            {
                Browser = "PowerPoint/ 16.0.18330",
                ProtocolVersion = "1.3",
                UserAgent = "Microsoft Office/16.0 (Windows NT 10.0; Microsoft PowerPoint 16.0.18330; Pro)",
                V8Version = "0.0",
                WebKitVersion = "0.0",
                WebSocketDebuggerUrl = "ws://localhost:53080/devtools/browser/abcd1234"
            };

            return Results.Ok(metadata);
        });

        app.MapGet("/json/activate/{id}", (string id) =>
        {
            if (id != "abcd1234")
            {
                return Results.NotFound($"No such target id: {id}");
            }

            sync.BeginInvoke(() =>
            {
                var window = this.Application.Windows.FirstOrDefault() as DocumentWindow;
                if (window != null)
                {
                    window.Activate();
                }
            }, null);
            return Results.Ok();
        });

        this.webApplication = app;
        Task.Run(async () =>
        {
            await app.RunAsync("http://localhost:53080");
            Trace.WriteLine($"Web server started. Visit http://localhost:53080");
        });
    }

    private void Addin_OnDisconnection(NetOffice.Tools.ext_DisconnectMode removeMode, ref Array custom)
    {
        this.webApplication?.StopAsync().Wait();
        Trace.WriteLine("Addin shutdown.");
    }
}
