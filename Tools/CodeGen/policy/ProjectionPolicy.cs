using System.Security.Cryptography;
using System.Text;
using System.Text.Json;
using System.Text.Json.Serialization;

namespace NetOffice.CodeGen.Projection;

/// <summary>Versioned, product-neutral rules consumed by the pure projection stage.</summary>
public sealed record ProjectionPolicy
{
    public const string CurrentSchema = "projection-policy/v1";
    public string SchemaVersion { get; init; } = CurrentSchema;
    public string DataSchemaVersion { get; init; } = "2.0";
    public string ContractSchemaVersion { get; init; } = "1.0";
    public string GeneratedRoot { get; init; } = "Generated";
    public string FilePartitionStrategy { get; init; } = "type";
    public string FileExtension { get; init; } = ".cs";
    public string DocsKeyPrefix { get; init; } = "vba";
    public ProjectionNameRules Names { get; init; } = new();
    public IReadOnlyList<ProjectionOverride> Overrides { get; init; } = Array.Empty<ProjectionOverride>();
    public IReadOnlyDictionary<string, string> RuntimeCapabilities { get; init; } = new Dictionary<string, string>(StringComparer.Ordinal);
    public string Digest { get; init; } = "";

    public static ProjectionPolicy Default => WithDigest(new ProjectionPolicy());

    public static ProjectionPolicy Parse(string json)
    {
        ArgumentNullException.ThrowIfNull(json);
        return JsonSerializer.Deserialize<ProjectionPolicy>(json, JsonOptions)
            ?? throw new ProjectionValidationException(new[] { new ProjectionIssue("policy.empty", "Policy JSON is empty.", "") });
    }

    public static ProjectionPolicy Read(string path) => Parse(File.ReadAllText(path));

    public static ProjectionPolicy WithDigest(ProjectionPolicy policy)
        => policy with { Digest = ComputeDigest(policy with { Digest = "" }) };

    public static string ComputeDigest(ProjectionPolicy policy)
    {
        ArgumentNullException.ThrowIfNull(policy);
        var bytes = Encoding.UTF8.GetBytes(JsonSerializer.Serialize(policy with { Digest = "" }, JsonOptions));
        return Convert.ToHexString(SHA256.HashData(bytes)).ToLowerInvariant();
    }

    public ValidationResult Validate(string? expectedDigest = null)
    {
        var issues = new List<ProjectionIssue>();
        if (!string.Equals(SchemaVersion, CurrentSchema, StringComparison.Ordinal))
            issues.Add(new("policy.schema", $"Expected {CurrentSchema}, got {SchemaVersion}.", "schemaVersion"));
        if (!string.Equals(DataSchemaVersion, NetOffice.CodeGen.Data.DataSchema.Version, StringComparison.Ordinal))
            issues.Add(new("policy.data-schema", $"Expected {NetOffice.CodeGen.Data.DataSchema.Version}, got {DataSchemaVersion}.", "dataSchemaVersion"));
        if (!string.Equals(ContractSchemaVersion, "1.0", StringComparison.Ordinal))
            issues.Add(new("policy.contract-schema", $"Expected 1.0, got {ContractSchemaVersion}.", "contractSchemaVersion"));
        if (string.IsNullOrWhiteSpace(GeneratedRoot) || Path.IsPathRooted(GeneratedRoot) || GeneratedRoot.Contains("..", StringComparison.Ordinal))
            issues.Add(new("policy.root", "GeneratedRoot must be a relative safe path.", "generatedRoot"));
        if (!string.Equals(FilePartitionStrategy, "type", StringComparison.Ordinal) && !string.Equals(FilePartitionStrategy, "source", StringComparison.Ordinal))
            issues.Add(new("policy.partition", "FilePartitionStrategy must be type or source.", "filePartitionStrategy"));
        if (!string.Equals(FileExtension, ".cs", StringComparison.Ordinal))
            issues.Add(new("policy.extension", "Only .cs is supported by the C# projection.", "fileExtension"));
        if (Names is null)
            issues.Add(new("policy.names", "Name rules are required.", "names"));
        else if (Names.ReservedWords is null)
            issues.Add(new("policy.names.reserved", "ReservedWords cannot be null.", "names.reservedWords"));
        var seenOverrides = new HashSet<string>(StringComparer.Ordinal);
        var allowedProperties = new HashSet<string>(new[] { "csharpName", "namespace", "signature", "docsBindingKey", "invocationOperation", "runtimeCapability" }, StringComparer.Ordinal);
        foreach (var item in Overrides ?? Array.Empty<ProjectionOverride>())
        {
            if (string.IsNullOrWhiteSpace(item.Property) || string.IsNullOrWhiteSpace(item.Value))
                issues.Add(new("policy.override.empty", "An override requires property and value.", item.Id ?? "overrides"));
            if (!allowedProperties.Contains(item.Property ?? ""))
                issues.Add(new("policy.override.property", $"Unsupported override property {item.Property ?? ""}.", item.Id ?? item.Property ?? "overrides"));
            if (item.TypeLogicalId is null && item.MemberLogicalId is null && item.MemberName is null)
                issues.Add(new("policy.override.selector", "An override requires a type or member selector.", item.Id ?? item.Property ?? "overrides"));
            var key = string.Join("|", item.TypeLogicalId ?? "", item.MemberLogicalId ?? "", item.MemberName ?? "", item.Property ?? "");
            if (!seenOverrides.Add(key))
                issues.Add(new("policy.override.conflict", $"Conflicting duplicate override selector {key}.", item.Id ?? key));
            if (item.ExpectedMatches < 1)
                issues.Add(new("policy.override.expected", "ExpectedMatches must be positive.", item.Id ?? "overrides"));
        }
        var actual = ComputeDigest(this);
        if (!string.Equals(Digest, actual, StringComparison.OrdinalIgnoreCase))
            issues.Add(new("policy.digest", $"Expected {actual}, got {Digest}.", "digest"));
        if (expectedDigest is not null && !string.Equals(expectedDigest, actual, StringComparison.OrdinalIgnoreCase))
            issues.Add(new("policy.expected-digest", $"Expected pinned digest {expectedDigest}, got {actual}.", "digest"));
        return new ValidationResult(issues);
    }

    public void ValidateOrThrow(string? expectedDigest = null) => Validate(expectedDigest).ThrowIfInvalid();

    internal static readonly JsonSerializerOptions JsonOptions = new()
    {
        PropertyNamingPolicy = JsonNamingPolicy.CamelCase,
        PropertyNameCaseInsensitive = true,
        DictionaryKeyPolicy = JsonNamingPolicy.CamelCase,
        DefaultIgnoreCondition = JsonIgnoreCondition.WhenWritingNull,
        WriteIndented = false
    };
}

public sealed record ProjectionNameRules
{
    public IReadOnlyList<string> ReservedWords { get; init; } = new[] { "class", "event", "interface", "namespace", "object", "string", "base", "this", "params", "ref", "out", "in", "void", "new", "public", "private", "protected", "internal", "static", "readonly", "get", "set", "value" };
    public string InvalidCharacterReplacement { get; init; } = "_";
}

public sealed record ProjectionOverride
{
    public string? Id { get; init; }
    public string? TypeLogicalId { get; init; }
    public string? MemberLogicalId { get; init; }
    public string? MemberName { get; init; }
    public string Property { get; init; } = "";
    public string Value { get; init; } = "";
    public int ExpectedMatches { get; init; } = 1;
}

public sealed record ProjectionIssue(string Code, string Message, string Path)
{
    public override string ToString() => string.IsNullOrEmpty(Path) ? $"{Code}: {Message}" : $"{Code} at {Path}: {Message}";
}

public sealed record ValidationResult(IReadOnlyList<ProjectionIssue> Issues)
{
    public bool IsValid => Issues.Count == 0;
    public void ThrowIfInvalid()
    {
        if (!IsValid) throw new ProjectionValidationException(Issues);
    }
}

public sealed class ProjectionValidationException : Exception
{
    public ProjectionValidationException(IReadOnlyList<ProjectionIssue> issues)
        : base(string.Join(Environment.NewLine, issues.Select(static x => x.ToString()))) => Issues = issues;
    public IReadOnlyList<ProjectionIssue> Issues { get; }
}
