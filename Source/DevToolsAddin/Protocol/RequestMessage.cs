using System;
using System.Text.Json.Nodes;

namespace NetOffice.DevToolsAddin.Protocol;

/// <summary>
/// A rpc call is represented by sending a Request object to a Server.
/// </summary>
/// <remarks>
/// JSON-RPC 2.0, see https://www.jsonrpc.org/specification
/// </remarks>
public struct RequestMessage
{
    public required int Id { get; init; }

    public required string Method { get; init; }

    public JsonValue? Params { get; init; }
}
