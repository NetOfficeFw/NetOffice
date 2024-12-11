
using Microsoft.AspNetCore.Builder;
using Microsoft.AspNetCore.Components;
using Microsoft.AspNetCore.Hosting;
using Microsoft.AspNetCore.Http;
using Microsoft.AspNetCore.Http.HttpResults;
using Microsoft.AspNetCore.Mvc;
using NetOffice.DevToolsAddin.Protocol;
using NetOffice.PowerPointApi;
using NetOffice.PowerPointApi.Events;
using NetOffice.PowerPointApi.Tools;
using System.Diagnostics;
using System.Net.WebSockets;
using System.Runtime.InteropServices;
using System.Text;
using System.Text.Json;
using System.Text.Json.Nodes;
using System.Text.Json.Serialization;
using System.Windows.Threading;

[ComVisible(true)]
[Guid("AB6D3A7D-33CF-4197-91D9-9D6B984DDDB1")]
[ProgId("NetOffice.DevToolsAddin")]
public class Addin : COMAddin
{
    private Task? webTask;
    private WebApplication? webApplication;

    public Addin()
    {
        this.OnConnection += Addin_OnConnection;
        this.OnStartupComplete += Addin_OnStartupComplete;
        this.OnDisconnection += Addin_OnDisconnection;
    }

    private void Addin_OnConnection(object application, NetOffice.Tools.ext_ConnectMode connectMode, object addInInst, ref Array custom)
    {
        Trace.WriteLine($"Addin connected to application. Mode: {connectMode}");

        var remotePort = Environment.GetEnvironmentVariable("PW_REMOTE_PORT");
        var userDataFolder = Environment.GetEnvironmentVariable("PW_USER_DATA_FOLDER");

        Console.WriteLine($"Playwright remote debugger port {remotePort}");
        Console.WriteLine($"Playwright user data directory {userDataFolder}");
    }

    private void Addin_OnStartupComplete(ref Array custom)
    {
        Trace.WriteLine($"Addin startup completed.");

        var powerpointPid = Process.GetCurrentProcess().Id;

        var sync = System.Windows.Threading.Dispatcher.CurrentDispatcher;

        var builder = WebApplication.CreateBuilder();
        //builder.WebHost.UseUrls("http://localhost:53080");
        var app = builder.Build();

        app.UseWebSockets();

        app.Use((context, next) =>
        {
            context.Response.Headers.Append("Content-Security-Policy", "frame-ancestors 'none'");
            return next();
        });

        // app.MapGet("/", () => "Hello World!");

        app.MapGet("/json/version", () =>
        {
            var metadata = new BrowserVersionMetadata
            {
                Browser = "PowerPoint/ 16.0.18330",
                ProtocolVersion = "1.3",
                UserAgent = "Microsoft Office/16.0 (Windows NT 10.0; Microsoft PowerPoint 16.0.18330; Pro)",
                V8Version = "0.0",
                WebKitVersion = "0.0",
                WebSocketDebuggerUrl = $"ws://localhost:53080/devtools/browser/powerpoint-pid-{powerpointPid}"
            };

            return Results.Ok(metadata);
        });

        app.MapGet("/json/activate/{id}", Results<NotFound<string>, Ok> (string id) =>
        {
            if (id != "abcd1234")
            {
                return TypedResults.NotFound($"No such target id: {id}");
            }

            sync.BeginInvoke(() =>
            {
                var window = this.Application.Windows.FirstOrDefault() as DocumentWindow;
                if (window != null)
                {
                    window.Activate();
                }
            }, null);
            return TypedResults.Ok();
        });

        app.Map("/devtools/browser/{id}", async (string id, HttpContext context) =>
        {
            if (!context.WebSockets.IsWebSocketRequest)
            {
                return Results.BadRequest();
            }

            var jsonOptions = new JsonSerializerOptions
            {
                PropertyNamingPolicy = JsonNamingPolicy.CamelCase,
                DefaultIgnoreCondition = JsonIgnoreCondition.WhenWritingNull
            };

            var sessionId = Guid.NewGuid().ToString("N").ToUpperInvariant();

            using var ws = await context.WebSockets.AcceptWebSocketAsync();
            while(true)
            {
                var receiveBuffer = WebSocket.CreateClientBuffer(4096, 4096);
                while (ws.State == WebSocketState.Open)
                {
                    var result = await ws.ReceiveAsync(receiveBuffer, CancellationToken.None);
                    if (result.MessageType == WebSocketMessageType.Text)
                    {
                        var text = Encoding.UTF8.GetString(receiveBuffer.AsSpan(0, result.Count));
                        Trace.WriteLine($"Received {result.Count} bytes: {text}");

                        var receivedMessage = JsonSerializer.Deserialize<RequestMessage>(text, jsonOptions);

                        if (receivedMessage.Method == "Browser.getVersion")
                        {
                            var responseBrowserGetVersion = new BrowserGetVersion
                            {
                                ProtocolVersion = "1.3",
                                Product = "PowerPoint/ 16.0.18330",
                                Revision = "16.0.18330",
                                UserAgent = "Microsoft Office/16.0 (Windows NT 10.0; Microsoft PowerPoint 16.0.18330; Pro)",
                                JsVersion = "0.0",
                            };

                            var responseMessage1 = ResponseMessage<BrowserGetVersion>.Create(receivedMessage.Id, responseBrowserGetVersion);
                            var responseBytes1 = JsonSerializer.SerializeToUtf8Bytes(responseMessage1, jsonOptions);
                            await ws.SendAsync(responseBytes1, WebSocketMessageType.Text, true, CancellationToken.None);
                        }
                        // HACK
                        else if (receivedMessage.Method == "Target.setAutoAttach")
                        {
                            if (receivedMessage.SessionId != null)
                            {
                                var responseMessage1 = ResponseMessage<object>.Create(receivedMessage.Id, receivedMessage.SessionId!);
                                var responseBytes1 = JsonSerializer.SerializeToUtf8Bytes(responseMessage1, jsonOptions);
                                await ws.SendAsync(responseBytes1, WebSocketMessageType.Text, true, CancellationToken.None);
                                continue;
                            }

                            var attachedToTarget = new TargetAttachedToTargetEventParams
                            {
                                SessionId = sessionId,
                                WaitingForDebugger = false,
                                TargetInfo = new TargetTargetInfo
                                {
                                    TargetId = "Presentation1",
                                    Type = "page",
                                    Title = "Presentation",
                                    Url = "about:blank",
                                    Attached = true,
                                    CanAccessOpener = false,
                                    BrowserContextId = id
                                }
                            };

                            var pushMessage1 = new RequestMessage
                            {
                                Id = default,
                                Method = "Target.attachedToTarget",
                                Params = JsonValue.Create(attachedToTarget)
                            };

                            var pushRequestBytes1 = JsonSerializer.SerializeToUtf8Bytes(pushMessage1, jsonOptions);
                            await ws.SendAsync(pushRequestBytes1, WebSocketMessageType.Text, true, CancellationToken.None);
                        }
                        else if (receivedMessage.Method == "Target.getTargetInfo")
                        {
                            var attachedToTarget = new TargetGetTargetInfoResponse
                            {
                                TargetInfo = new TargetTargetInfo
                                {
                                    TargetId = Guid.NewGuid().ToString("D"),
                                    Type = "browser",
                                    Title = "",
                                    Url = "",
                                    Attached = true,
                                    CanAccessOpener = false,
                                }
                            };

                            var responseMessage1 = ResponseMessage<TargetGetTargetInfoResponse>.Create(receivedMessage.Id, attachedToTarget);

                            var pushRequestBytes1 = JsonSerializer.SerializeToUtf8Bytes(responseMessage1, jsonOptions);
                            await ws.SendAsync(pushRequestBytes1, WebSocketMessageType.Text, true, CancellationToken.None);
                        }
                        else if (
                            receivedMessage.Method == "Page.enable" ||
                            receivedMessage.Method == "Log.enable" ||
                            receivedMessage.Method == "Page.setLifecycleEventsEnabled"
                            )
                        {
                            var responseMessage1 = ResponseMessage<object>.Create(receivedMessage.Id, receivedMessage.SessionId!);
                            var responseBytes1 = JsonSerializer.SerializeToUtf8Bytes(responseMessage1, jsonOptions);
                            await ws.SendAsync(responseBytes1, WebSocketMessageType.Text, true, CancellationToken.None);
                        }
                        else
                        {
                            // default empty response
                            var responseMessage = ResponseMessage<object>.Create(receivedMessage.Id, new object());
                            if (receivedMessage.SessionId != null)
                            {
                                responseMessage = responseMessage with { SessionId = receivedMessage.SessionId };
                            }

                            var responseBytes = JsonSerializer.SerializeToUtf8Bytes(responseMessage, jsonOptions);
                            await ws.SendAsync(responseBytes, WebSocketMessageType.Text, true, CancellationToken.None);
                        }
                    }
                    else if (result.MessageType == WebSocketMessageType.Close)
                    {
                        return Results.Empty;
                    }
                }

                if (ws.State == System.Net.WebSockets.WebSocketState.Closed || ws.State == System.Net.WebSockets.WebSocketState.Aborted)
                {
                    break;
                }

                await Task.Delay(200);
            }

            return Results.Ok();
        });

        this.webApplication = app;
        Task.Run(async () =>
        {
            Console.WriteLine("NetOffice DevTools server running");
            Console.OpenStandardOutput().Flush();
            Trace.WriteLine("NetOffice DevTools server running");
            Trace.WriteLine($"Web server started. Visit http://localhost:53080");
            await app.RunAsync("http://localhost:53080");
        });
    }

    private void Addin_OnDisconnection(NetOffice.Tools.ext_DisconnectMode removeMode, ref Array custom)
    {
        this.webApplication?.StopAsync().Wait();
        Trace.WriteLine("Addin shutdown.");
    }
}
