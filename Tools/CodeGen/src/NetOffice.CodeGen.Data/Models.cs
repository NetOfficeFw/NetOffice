using System.Collections.ObjectModel;

namespace NetOffice.CodeGen.Data;

public static class DataSchema
{
    public const string Version = "2.0";
    public const string Serialization = "netoffice-data-v2-json";
    public const string DigestAlgorithm = "SHA-256";
}

public sealed record DataSource
{
    public string Kind { get; init; } = "canonical-input";
    public string Path { get; init; } = "";
    public string Sha256 { get; init; } = "";
    public string? Revision { get; init; }
}

public sealed record Provenance
{
    public string SourcePath { get; init; } = "";
    public string SourceSha256 { get; init; } = "";
    public string? SourceRevision { get; init; }
    public string? Location { get; init; }
}

public sealed record DataLibrary
{
    public string LogicalId { get; init; } = "";
    public string Name { get; init; } = "";
    public string Guid { get; init; } = "";
    public string Version { get; init; } = "";
    public Provenance Provenance { get; init; } = new();
}

public sealed record DataType
{
    public string LogicalId { get; init; } = "";
    public string LibraryId { get; init; } = "";
    public string Name { get; init; } = "";
    public string Kind { get; init; } = "";
    public string SourceKey { get; init; } = "";
    public IReadOnlyList<string> BaseTypeIds { get; init; } = Array.Empty<string>();
    public Provenance Provenance { get; init; } = new();
}

public sealed record DataMember
{
    public string LogicalId { get; init; } = "";
    public string TypeId { get; init; } = "";
    public string Name { get; init; } = "";
    public string Kind { get; init; } = "";
    public string SourceKey { get; init; } = "";
    public int? DispId { get; init; }
    public string? ReturnType { get; init; }
    public IReadOnlyList<string> ParameterTypes { get; init; } = Array.Empty<string>();
    public string? AccessorGroupId { get; init; }
    public Provenance Provenance { get; init; } = new();
}

public sealed record AccessorGroup
{
    public string LogicalId { get; init; } = "";
    public string TypeId { get; init; } = "";
    public string Name { get; init; } = "";
    public IReadOnlyList<string> MemberIds { get; init; } = Array.Empty<string>();
}

public sealed record AmbiguityRecord
{
    public string LogicalId { get; init; } = "";
    public string Kind { get; init; } = "";
    public string Reason { get; init; } = "";
    public IReadOnlyList<string> Candidates { get; init; } = Array.Empty<string>();
    public Provenance Provenance { get; init; } = new();
    public string? Resolution { get; init; }
}

public sealed record DataGraph
{
    public string SchemaVersion { get; init; } = DataSchema.Version;
    public string Serialization { get; init; } = DataSchema.Serialization;
    public DataSource Source { get; init; } = new();
    public IReadOnlyList<DataLibrary> Libraries { get; init; } = Array.Empty<DataLibrary>();
    public IReadOnlyList<DataType> Types { get; init; } = Array.Empty<DataType>();
    public IReadOnlyList<DataMember> Members { get; init; } = Array.Empty<DataMember>();
    public IReadOnlyList<AccessorGroup> AccessorGroups { get; init; } = Array.Empty<AccessorGroup>();
    public IReadOnlyList<AmbiguityRecord> Ambiguities { get; init; } = Array.Empty<AmbiguityRecord>();
    public string Digest { get; init; } = "";
}

public sealed record ValidationIssue(string Code, string Message, string? Path = null)
{
    public override string ToString() => Path is null ? $"{Code}: {Message}" : $"{Code} at {Path}: {Message}";
}

public sealed record ValidationResult(IReadOnlyList<ValidationIssue> Issues)
{
    public bool IsValid => Issues.Count == 0;

    public void ThrowIfInvalid()
    {
        if (!IsValid)
            throw new DataValidationException(this);
    }
}

public sealed class DataValidationException : Exception
{
    public DataValidationException(ValidationResult result)
        : base(string.Join(Environment.NewLine, result.Issues.Select(static issue => issue.ToString())))
    {
        Result = result;
    }

    public ValidationResult Result { get; }
}

public sealed record TreeHashResult(string Algorithm, string Digest, IReadOnlyList<string> Files);
