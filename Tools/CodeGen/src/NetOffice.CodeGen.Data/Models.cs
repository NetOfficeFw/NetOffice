
namespace NetOffice.CodeGen.Data;

public static class DataSchema
{
    public const string Version = "2.0";
    public const string ApiVersion = "2.2.0";
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
public sealed record DataProject
{
    public string LogicalId { get; init; } = "";
    public string Name { get; init; } = "";
    public string Namespace { get; init; } = "";
    public string SourceKey { get; init; } = "";
    public string Version { get; init; } = "";
    public string FileVersion { get; init; } = "";
    public bool Ignore { get; init; }
    public IReadOnlyList<string> SourceCategories { get; init; } = Array.Empty<string>();
    public IReadOnlyList<string> LibraryIds { get; init; } = Array.Empty<string>();
    public IReadOnlyList<string> ReferenceProjectIds { get; init; } = Array.Empty<string>();
    public Provenance Provenance { get; init; } = new();
}

public sealed record DataLibraryDependency
{
    public string Name { get; init; } = "";
    public string Guid { get; init; } = "";
    public string Major { get; init; } = "";
    public string Minor { get; init; } = "";
    public string? Description { get; init; }
    public IReadOnlyList<string> TargetLibraryIds { get; init; } = Array.Empty<string>();
    public Provenance Provenance { get; init; } = new();
}


public sealed record DataLibrary
{
    public string LogicalId { get; init; } = "";
    public string Name { get; init; } = "";
    public string Guid { get; init; } = "";
    public string SourceKey { get; init; } = "";
    public string Version { get; init; } = "";
    public string Major { get; init; } = "";
    public string Minor { get; init; } = "";
    public string? Description { get; init; }
    public IReadOnlyList<DataLibraryDependency> Dependencies { get; init; } = Array.Empty<DataLibraryDependency>();
    public Provenance Provenance { get; init; } = new();
}
public sealed record DataTypeReference
{
    public string Name { get; init; } = "";
    public string? TypeKind { get; init; }
    public string? VarType { get; init; }
    public string? MarshalAs { get; init; }
    public string? TypeKey { get; init; }
    public string? TargetTypeId { get; init; }
    public string? ProjectKey { get; init; }
    public string? LibraryKey { get; init; }
    public bool IsComProxy { get; init; }
    public bool IsExternal { get; init; }
    public bool IsEnum { get; init; }
    public bool IsArray { get; init; }
    public bool IsNative { get; init; }
}

public sealed record DataIdentifierObservation
{
    public string Value { get; init; } = "";
    public IReadOnlyList<string> LibraryIds { get; init; } = Array.Empty<string>();
    public Provenance Provenance { get; init; } = new();
}

public sealed record DataTypeLinkObservation
{
    public string Kind { get; init; } = "";
    public string TargetTypeId { get; init; } = "";
    public IReadOnlyList<string> LibraryIds { get; init; } = Array.Empty<string>();
    public Provenance Provenance { get; init; } = new();
}


public sealed record DataParameter
{
    public string Name { get; init; } = "";
    public string Type { get; init; } = "";
    public string RefKind { get; init; } = "value";
    public bool IsOptional { get; init; }
    public bool HasDefaultValue { get; init; }
    public string? DefaultValue { get; init; }
    public DataTypeReference? TypeReference { get; init; }
    public string? ParamFlags { get; init; }
    public Provenance Provenance { get; init; } = new();
}

public sealed record DataValue
{
    public string LogicalId { get; init; } = "";
    public string TypeId { get; init; } = "";
    public string Name { get; init; } = "";
    public string Kind { get; init; } = "";
    public string Value { get; init; } = "";
    public string? ValueType { get; init; }
    public IReadOnlyList<string> SupportObservationIds { get; init; } = Array.Empty<string>();
    public Provenance Provenance { get; init; } = new();
}

public sealed record DataType
{
    public string LogicalId { get; init; } = "";
    public string LibraryId { get; init; } = "";
    public string ProjectId { get; init; } = "";
    public string Namespace { get; init; } = "";
    public string SourceCategory { get; init; } = "";
    public string Name { get; init; } = "";
    public string Kind { get; init; } = "";
    public string SourceKey { get; init; } = "";
    public string? AliasTarget { get; init; }
    public string? DeclaredGuid { get; init; }
    public int? TypeLibType { get; init; }
    public bool? IsEventInterface { get; init; }
    public bool? IsEarlyBind { get; init; }
    public bool? IsHidden { get; init; }
    public bool? AutomaticQuit { get; init; }
    public bool? IsApplicationObject { get; init; }
    public IReadOnlyList<DataIdentifierObservation> GuidObservations { get; init; } = Array.Empty<DataIdentifierObservation>();
    public IReadOnlyList<DataTypeLinkObservation> ReferenceObservations { get; init; } = Array.Empty<DataTypeLinkObservation>();
    public IReadOnlyList<string> BaseTypeIds { get; init; } = Array.Empty<string>();
    public IReadOnlyList<string> DefaultInterfaceIds { get; init; } = Array.Empty<string>();
    public IReadOnlyList<string> EventInterfaceIds { get; init; } = Array.Empty<string>();
    public IReadOnlyList<string> SupportObservationIds { get; init; } = Array.Empty<string>();
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
    public IReadOnlyList<DataIdentifierObservation> DispIdObservations { get; init; } = Array.Empty<DataIdentifierObservation>();
    public string? ReturnType { get; init; }
    public DataTypeReference? ReturnTypeReference { get; init; }
    public bool IsHidden { get; init; }
    public bool AnalyzeReturn { get; init; }
    public bool IsComProxy { get; init; }
    // ParameterTypes remains as a compact compatibility view for existing consumers.
    // Parameters is the emit-ready signature and is authoritative for new producers.
    public IReadOnlyList<string> ParameterTypes { get; init; } = Array.Empty<string>();
    public IReadOnlyList<DataParameter> Parameters { get; init; } = Array.Empty<DataParameter>();
    public string? AccessorGroupId { get; init; }
    public string? AccessorKind { get; init; }
    public string? Value { get; init; }
    public string? ValueType { get; init; }
    public IReadOnlyList<string> SupportObservationIds { get; init; } = Array.Empty<string>();
    public IReadOnlyList<string> SignatureLibraryIds { get; init; } = Array.Empty<string>();
    public string? InvocationEvidenceId { get; init; }
    public Provenance Provenance { get; init; } = new();
}

public sealed record AccessorGroup
{
    public string LogicalId { get; init; } = "";
    public string TypeId { get; init; } = "";
    public string Name { get; init; } = "";
    public string Kind { get; init; } = "property";
    public IReadOnlyList<string> MemberIds { get; init; } = Array.Empty<string>();
    public string? EvidenceId { get; init; }
    public Provenance Provenance { get; init; } = new();
}

public sealed record SupportObservation
{
    public string LogicalId { get; init; } = "";
    public string TargetId { get; init; } = "";
    public string Product { get; init; } = "";
    public IReadOnlyList<string> Versions { get; init; } = Array.Empty<string>();
    public IReadOnlyList<string> LibraryIds { get; init; } = Array.Empty<string>();
    public bool Present { get; init; } = true;
    public Provenance Provenance { get; init; } = new();
}

public sealed record AliasRecord
{
    public string LogicalId { get; init; } = "";
    public string Alias { get; init; } = "";
    public string TargetId { get; init; } = "";
    public string Kind { get; init; } = "";
    public Provenance Provenance { get; init; } = new();
}

public sealed record UnificationRecord
{
    public string LogicalId { get; init; } = "";
    public string CanonicalId { get; init; } = "";
    public IReadOnlyList<string> EquivalentIds { get; init; } = Array.Empty<string>();
    public string Reason { get; init; } = "";
    public Provenance Provenance { get; init; } = new();
}

public sealed record InvocationEvidence
{
    public string LogicalId { get; init; } = "";
    public string MemberId { get; init; } = "";
    public string Operation { get; init; } = "";
    public string DispatchName { get; init; } = "";
    public int? DispId { get; init; }
    public int ArgumentCount { get; init; }
    public string? ResultType { get; init; }
    public bool RequiresProxy { get; init; }
    public Provenance Provenance { get; init; } = new();
}

public sealed record AccessorEvidence
{
    public string LogicalId { get; init; } = "";
    public string AccessorGroupId { get; init; } = "";
    public string Kind { get; init; } = "";
    public IReadOnlyList<string> MemberIds { get; init; } = Array.Empty<string>();
    public Provenance Provenance { get; init; } = new();
}

public sealed record UnknownRecord
{
    public string LogicalId { get; init; } = "";
    public string Kind { get; init; } = "";
    public string Reason { get; init; } = "";
    public Provenance Provenance { get; init; } = new();
}

public sealed record StaleRecord
{
    public string LogicalId { get; init; } = "";
    public string TargetId { get; init; } = "";
    public string ExpectedDigest { get; init; } = "";
    public string ActualDigest { get; init; } = "";
    public string Reason { get; init; } = "";
    public Provenance Provenance { get; init; } = new();
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
public sealed record AbsentFactRecord
{
    public string LogicalId { get; init; } = "";
    public string? TargetId { get; init; }
    public string Kind { get; init; } = "";
    public string Reason { get; init; } = "";
    public Provenance Provenance { get; init; } = new();
}


public sealed record DataGraph
{
    public string SchemaVersion { get; init; } = DataSchema.Version;
    public string Serialization { get; init; } = DataSchema.Serialization;
    public DataSource Source { get; init; } = new();
    public IReadOnlyList<DataProject> Projects { get; init; } = Array.Empty<DataProject>();
    public IReadOnlyList<DataLibrary> Libraries { get; init; } = Array.Empty<DataLibrary>();
    public IReadOnlyList<DataType> Types { get; init; } = Array.Empty<DataType>();
    public IReadOnlyList<DataMember> Members { get; init; } = Array.Empty<DataMember>();
    public IReadOnlyList<DataValue> Values { get; init; } = Array.Empty<DataValue>();
    public IReadOnlyList<AccessorGroup> AccessorGroups { get; init; } = Array.Empty<AccessorGroup>();
    public IReadOnlyList<SupportObservation> SupportObservations { get; init; } = Array.Empty<SupportObservation>();
    public IReadOnlyList<AliasRecord> Aliases { get; init; } = Array.Empty<AliasRecord>();
    public IReadOnlyList<UnificationRecord> Unifications { get; init; } = Array.Empty<UnificationRecord>();
    public IReadOnlyList<InvocationEvidence> InvocationEvidence { get; init; } = Array.Empty<InvocationEvidence>();
    public IReadOnlyList<AccessorEvidence> AccessorEvidence { get; init; } = Array.Empty<AccessorEvidence>();
    public IReadOnlyList<AbsentFactRecord> AbsentFacts { get; init; } = Array.Empty<AbsentFactRecord>();
    public IReadOnlyList<UnknownRecord> Unknowns { get; init; } = Array.Empty<UnknownRecord>();
    public IReadOnlyList<StaleRecord> StaleRecords { get; init; } = Array.Empty<StaleRecord>();
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
