using System;
using TargetID = System.String;

namespace NetOffice.DevToolsAddin.Protocol;

/// <summary>
/// Issued when attached to target because of auto-attach or `attachToTarget` command.
/// </summary>
public struct TargetAttachedToTargetEvent
{
    public required TargetAttachedToTargetEventParams TargetInfo { get; init; }
}

public readonly record struct TargetAttachedToTargetEventParams
{
    public required string SessionId { get; init; }
    public required TargetTargetInfo TargetInfo { get; init; }
    public required bool WaitingForDebugger { get; init; }

}

public readonly record struct TargetGetTargetInfoResponse
{
    public required TargetTargetInfo TargetInfo { get; init; }
}

public readonly record struct TargetTargetInfo
{
    public required TargetID TargetId { get; init; }

    public required string Type { get; init; }

    public required string Title { get; init; }

    public required string Url { get; init; }

    public required bool Attached { get; init; }

    public TargetID? OpenerId { get; init; }

    public required bool CanAccessOpener { get; init; }

    public TargetID? OpenerFrameId { get; init; }

    public TargetID? BrowserContextId { get; init; }

    public string Subtype { get; init; }
}
