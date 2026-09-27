using System.Text.Json;
using System.Text.Json.Serialization;

namespace NetOffice.CodeGen.Projection;

public sealed record ProjectionOptions
{
    public string? ExpectedDataDigest { get; init; }
    public string? ExpectedContractDigest { get; init; }
    public string? ExpectedPolicyDigest { get; init; }
    public bool RejectUnresolvedAmbiguities { get; init; } = true;
}

public sealed record WrapperContract
{
    public string SchemaVersion { get; init; } = "";
    public string ContractKind { get; init; } = "";
    public ContractGenerator Generator { get; init; } = new();
    public ContractSource Source { get; init; } = new();
    public IReadOnlyList<ContractType> Types { get; init; } = Array.Empty<ContractType>();
    public IReadOnlyList<ContractUnknown> Unknowns { get; init; } = Array.Empty<ContractUnknown>();
    public IReadOnlyList<ContractAmbiguity> Ambiguities { get; init; } = Array.Empty<ContractAmbiguity>();

    public static WrapperContract Parse(string json)
        => JsonSerializer.Deserialize<WrapperContract>(json, ProjectionPolicy.JsonOptions)
           ?? throw new ProjectionValidationException(new[] { new ProjectionIssue("contract.empty", "Contract JSON is empty.", "") });
    public static WrapperContract Read(string path) => Parse(File.ReadAllText(path));
}

public sealed record ContractGenerator { public string Name { get; init; } = ""; public string Version { get; init; } = ""; }
public sealed record ContractSource { public string Root { get; init; } = ""; public string Api { get; init; } = ""; public IReadOnlyList<string> SourceRoots { get; init; } = Array.Empty<string>(); public IReadOnlyList<ContractFile> Files { get; init; } = Array.Empty<ContractFile>(); }
public sealed record ContractFile { public string Path { get; init; } = ""; public string Sha256 { get; init; } = ""; public int LineCount { get; init; } }
public sealed record ContractType
{
    public string LogicalId { get; init; } = "";
    public string Namespace { get; init; } = "";
    public string Name { get; init; } = "";
    public string Kind { get; init; } = "class";
    public string Accessibility { get; init; } = "public";
    public IReadOnlyList<string> Modifiers { get; init; } = Array.Empty<string>();
    public string? BaseType { get; init; }
    public IReadOnlyList<string> Interfaces { get; init; } = Array.Empty<string>();
    public IReadOnlyList<string> Attributes { get; init; } = Array.Empty<string>();
    public string Source { get; init; } = "";
    public int Line { get; init; }
    public string Signature { get; init; } = "";
    public IReadOnlyList<ContractMember> Members { get; init; } = Array.Empty<ContractMember>();
}
public sealed record ContractMember
{
    public string Name { get; init; } = "";
    public string Kind { get; init; } = "";
    public string Accessibility { get; init; } = "public";
    public IReadOnlyList<string> Modifiers { get; init; } = Array.Empty<string>();
    public string? ReturnType { get; init; }
    public string? Parameters { get; init; }
    public string Signature { get; init; } = "";
    public IReadOnlyList<string> Attributes { get; init; } = Array.Empty<string>();
    public IReadOnlyList<string> InvocationText { get; init; } = Array.Empty<string>();
    public string Source { get; init; } = "";
    public int Line { get; init; }
}
public sealed record ContractUnknown { public string Code { get; init; } = ""; public string Message { get; init; } = ""; public string Path { get; init; } = ""; public int? Line { get; init; } public string? LogicalId { get; init; } }
public sealed record ContractAmbiguity { public string Code { get; init; } = ""; public string Message { get; init; } = ""; public IReadOnlyList<string> Candidates { get; init; } = Array.Empty<string>(); }

public sealed record ProjectionResult
{
    public string SchemaVersion { get; init; } = "projection/v1";
    public string DataSchemaVersion { get; init; } = "";
    public string DataDigest { get; init; } = "";
    public string ContractSchemaVersion { get; init; } = "";
    public string ContractDigest { get; init; } = "";
    public string PolicyDigest { get; init; } = "";
    public IReadOnlyList<WrapperFile> Files { get; init; } = Array.Empty<WrapperFile>();
    public IReadOnlyList<ProjectionTrace> Explain { get; init; } = Array.Empty<ProjectionTrace>();

    public string ToJson(bool indented = false) => JsonSerializer.Serialize(this, new JsonSerializerOptions(ProjectionPolicy.JsonOptions) { WriteIndented = indented });
}

public sealed record WrapperFile
{
    public string Path { get; init; } = "";
    public string Namespace { get; init; } = "";
    public IReadOnlyList<ProjectedType> Types { get; init; } = Array.Empty<ProjectedType>();
}
public sealed record ProjectedType
{
    public string LogicalId { get; init; } = "";
    public string DataLogicalId { get; init; } = "";
    public string Namespace { get; init; } = "";
    public string Name { get; init; } = "";
    public string CSharpName { get; init; } = "";
    public string Kind { get; init; } = "";
    public string? BaseType { get; init; }
    public IReadOnlyList<string> Interfaces { get; init; } = Array.Empty<string>();
    public IReadOnlyList<string> EffectiveBaseTypes { get; init; } = Array.Empty<string>();
    public IReadOnlyList<string> DuplicateGroups { get; init; } = Array.Empty<string>();
    public string Signature { get; init; } = "";
    public IReadOnlyList<VersionSupport> SupportVersions { get; init; } = Array.Empty<VersionSupport>();
    public IReadOnlyList<string> Capabilities { get; init; } = Array.Empty<string>();
    public IReadOnlyList<string> RuntimeRequirements { get; init; } = Array.Empty<string>();
    public string DocsBindingKey { get; init; } = "";
    public string Source { get; init; } = "";
    public int Line { get; init; }
    public IReadOnlyList<ProjectedMember> Members { get; init; } = Array.Empty<ProjectedMember>();
}
public sealed record ProjectedMember
{
    public string LogicalId { get; init; } = "";
    public string DataLogicalId { get; init; } = "";
    public string Name { get; init; } = "";
    public string CSharpName { get; init; } = "";
    public string Kind { get; init; } = "";
    public string Accessibility { get; init; } = "";
    public IReadOnlyList<string> Modifiers { get; init; } = Array.Empty<string>();
    public string? ReturnType { get; init; }
    public string? Parameters { get; init; }
    public string Signature { get; init; } = "";
    public string OverloadGroup { get; init; } = "";
    public string? DuplicateOf { get; init; }
    public InvocationPlan Invocation { get; init; } = new();
    public IReadOnlyList<VersionSupport> SupportVersions { get; init; } = Array.Empty<VersionSupport>();
    public IReadOnlyList<string> Capabilities { get; init; } = Array.Empty<string>();
    public IReadOnlyList<string> RuntimeRequirements { get; init; } = Array.Empty<string>();
    public string DocsBindingKey { get; init; } = "";
    public IReadOnlyList<string> Attributes { get; init; } = Array.Empty<string>();
    public string Source { get; init; } = "";
    public int Line { get; init; }
}
public sealed record VersionSupport { public string Product { get; init; } = ""; public IReadOnlyList<string> Versions { get; init; } = Array.Empty<string>(); }
public sealed record InvocationPlan
{
    public string Operation { get; init; } = "method";
    public string DispatchName { get; init; } = "";
    public int? DispId { get; init; }
    public int ArgumentCount { get; init; }
    public string? ResultType { get; init; }
    public bool RequiresProxy { get; init; }
    public IReadOnlyList<string> Text { get; init; } = Array.Empty<string>();
}
public sealed record ProjectionTrace
{
    public string LogicalId { get; init; } = "";
    public IReadOnlyList<ProjectionTraceStep> Steps { get; init; } = Array.Empty<ProjectionTraceStep>();
}
public sealed record ProjectionTraceStep { public string Stage { get; init; } = ""; public string Rule { get; init; } = ""; public string Result { get; init; } = ""; }
