using System.Collections.Immutable;

namespace NetOffice.CodeGen.TypeLib;

public enum TypeLibProvenanceIssueKind
{
    MissingCurrentChannel,
    WrongChannel,
    MissingField,
    WrongField,
    BinaryDigestMismatch,
    InvalidSchema
}

public sealed record TypeLibProvenanceIssue(TypeLibProvenanceIssueKind Kind, string Field, string Message);

/// <summary>Validation policy for imported typelib provenance.</summary>
public sealed record TypeLibProvenancePolicy(
    bool RequireCurrentChannel,
    CurrentChannelMetadata? ExpectedCurrentChannel,
    string? ExpectedBinarySha256)
{
    public static TypeLibProvenancePolicy CurrentChannel(string? expectedBinarySha256 = null) => new(
        true,
        null,
        expectedBinarySha256);
}

/// <summary>Machine-readable provenance validation outcome.</summary>
public sealed class TypeLibProvenanceReport
{
    internal TypeLibProvenanceReport(ImmutableArray<TypeLibProvenanceIssue> issues) => Issues = issues;
    public ImmutableArray<TypeLibProvenanceIssue> Issues { get; }
    public bool IsValid => Issues.IsEmpty;
}

public static class TypeLibProvenanceValidator
{
    public static TypeLibProvenanceReport Validate(TypeLibObservation observation, TypeLibProvenancePolicy policy)
    {
        ArgumentNullException.ThrowIfNull(observation);
        ArgumentNullException.ThrowIfNull(policy);
        var issues = ImmutableArray.CreateBuilder<TypeLibProvenanceIssue>();
        if (!string.Equals(observation.SchemaVersion, TypeLibObservationSchema.Version, StringComparison.Ordinal))
            issues.Add(new(TypeLibProvenanceIssueKind.InvalidSchema, "schemaVersion", $"Unsupported observation schema '{observation.SchemaVersion}'."));

        var channel = observation.Provenance.CurrentChannel;
        if (policy.RequireCurrentChannel && channel is null)
            issues.Add(new(TypeLibProvenanceIssueKind.MissingCurrentChannel, "currentChannel", "Current Channel metadata is required."));
        if (channel is not null)
        {
            if (!Guid.TryParse(channel.ChannelGuid, out var channelGuid) || channelGuid != Guid.Parse(CurrentChannelMetadata.RequiredChannelGuid))
                issues.Add(new(TypeLibProvenanceIssueKind.WrongChannel, "channelGuid", $"Expected Current Channel GUID '{CurrentChannelMetadata.RequiredChannelGuid}'."));
            ValidateNonEmpty(issues, "build", channel.Build);
            ValidateNonEmpty(issues, "sku", channel.Sku);
            ValidateNonEmpty(issues, "locale", channel.Locale);
            ValidateNonEmpty(issues, "architecture", channel.Architecture);
            if (policy.ExpectedCurrentChannel is not null)
            {
                var expected = policy.ExpectedCurrentChannel;
                Compare(issues, "channelGuid", expected.ChannelGuid, channel.ChannelGuid);
                Compare(issues, "build", expected.Build, channel.Build);
                Compare(issues, "sku", expected.Sku, channel.Sku);
                Compare(issues, "locale", expected.Locale, channel.Locale);
                Compare(issues, "architecture", expected.Architecture, channel.Architecture);
            }
        }

        if (policy.ExpectedBinarySha256 is not null && !string.Equals(policy.ExpectedBinarySha256, observation.Provenance.BinarySha256, StringComparison.OrdinalIgnoreCase))
            issues.Add(new(TypeLibProvenanceIssueKind.BinaryDigestMismatch, "binarySha256", $"Expected '{policy.ExpectedBinarySha256}' but found '{observation.Provenance.BinarySha256}'."));
        return new TypeLibProvenanceReport(issues.ToImmutable());
    }

    private static void ValidateNonEmpty(ImmutableArray<TypeLibProvenanceIssue>.Builder issues, string field, string value)
    {
        if (value.Length == 0)
            issues.Add(new(TypeLibProvenanceIssueKind.MissingField, field, $"Current Channel field '{field}' is empty."));
    }

    private static void Compare(ImmutableArray<TypeLibProvenanceIssue>.Builder issues, string field, string expected, string actual)
    {
        if (!string.Equals(expected, actual, StringComparison.OrdinalIgnoreCase))
            issues.Add(new(TypeLibProvenanceIssueKind.WrongField, field, $"Expected '{expected}' but found '{actual}'."));
    }
}
