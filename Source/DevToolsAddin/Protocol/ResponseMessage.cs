using System;
using System.Text.Json.Nodes;

namespace NetOffice.DevToolsAddin.Protocol;

/// <summary>
/// A rpc call is represented by sending a Request object to a Server.
/// </summary>
/// <remarks>
/// JSON-RPC 2.0, see https://www.jsonrpc.org/specification
/// </remarks>
public struct ResponseMessage<TResult>
{
    public required int Id { get; init; }

    public TResult? Result { get; init; }

    public JsonRpcError? Error { get; init; }

    public static ResponseMessage<TResult> Create(int id, TResult result)
    {
        var response = new ResponseMessage<TResult>
        { Id = id, Result = result, Error = default };

        return response;
    }

    public static ResponseMessage<object> Create(int id, int code, string message, object data)
    {
        var response = new ResponseMessage<object>
        {
            Id = id,
            Result = default,
            Error = new JsonRpcError
            {
                Code = code,
                Message = message,
                Data = JsonValue.Create(data)
            }
        };

        return response;
    }
}

public struct JsonRpcError
{
    public required int Code { get; init; }

    public required string Message { get; init; }

    public JsonValue? Data { get; init; }
}
