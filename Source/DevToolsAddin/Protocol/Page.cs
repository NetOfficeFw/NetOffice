using FrameId = System.String;
using NetworkLoaderId = System.String;

namespace NetOffice.DevToolsAddin.Protocol;

public readonly record struct PageGetFrameTreeResponse
{
    public required PageFrameTree FrameTree { get; init; }
}

public readonly record struct PageFrameTree
{
    public required PageFrame Frame { get; init; }

    public PageFrameTree[]? ChildFrames { get; init; }
}


public readonly record struct PageFrame
{
    public required FrameId Id { get; init; }

    public FrameId? ParentId { get; init; }

    public NetworkLoaderId LoaderId { get; init; }

    public string? Name { get; init; }

    public required string Url { get; init; }

    public required string DomainAndRegistry { get; init; }

    public required string SecurityOrigin { get; init; }

    public required string MimeType { get; init; }

    public required string SecureContextType { get; init; }
    public required string CrossOriginIsolatedContextType { get; init; }
    public required string[] GatedAPIFeatures { get; init; }
}

public readonly record struct PageAddScriptToEvaluateOnNewDocumentResponse
{
    public string Identifier { get; init; }
}
