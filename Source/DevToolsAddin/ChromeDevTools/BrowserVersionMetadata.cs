using System.Text.Json.Serialization;

public class BrowserVersionMetadata
{
    [JsonPropertyName("Browser")]
    public required string Browser { get; init; }

    [JsonPropertyName("Protocol-Version")]
    public required string ProtocolVersion { get; init; }

    [JsonPropertyName("User-Agent")]
    public required string UserAgent { get; init; }

    [JsonPropertyName("V8-Version")]
    public required string V8Version { get; init; }

    [JsonPropertyName("WebKit-Version")]
    public required string WebKitVersion { get; init; }

    [JsonPropertyName("webSocketDebuggerUrl")]
    public required string WebSocketDebuggerUrl { get; init; }
}

public class PowerPointAppVersionMetadata
{
    public required string AppType { get; init; }

    public required string Version { get; init; }

    public required int ProcessId { get; init; }

    public required string GrpcDebuggerUrl { get; init; }
}