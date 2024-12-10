using System;

namespace NetOffice.DevToolsAddin.Protocol;

internal class Browser
{
}

public readonly record struct BrowserGetVersion
{
    public required string ProtocolVersion { get; init; }

    public required string Product { get; init; }

    public required string Revision { get; init; }

    public required string UserAgent { get; init; }

    public required string JsVersion { get; init; }
}
