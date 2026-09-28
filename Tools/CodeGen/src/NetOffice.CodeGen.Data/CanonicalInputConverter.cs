using System.Text.Json;
using System.Text.Json.Serialization;

namespace NetOffice.CodeGen.Data;

public sealed record CanonicalInput
{
    public string Format { get; init; } = "netoffice-data-input-v1";
    public DataSource Source { get; init; } = new();
    public IReadOnlyList<CanonicalInputLibrary> Libraries { get; init; } = Array.Empty<CanonicalInputLibrary>();
    public IReadOnlyList<CanonicalInputAlias> Aliases { get; init; } = Array.Empty<CanonicalInputAlias>();
    public IReadOnlyList<CanonicalInputUnification> Unifications { get; init; } = Array.Empty<CanonicalInputUnification>();
}

public sealed record CanonicalInputParameter
{
    public string Name { get; init; } = "";
    public string Type { get; init; } = "";
    public string RefKind { get; init; } = "value";
    public bool IsOptional { get; init; }
    public string? DefaultValue { get; init; }
}

public sealed record CanonicalInputAlias
{
    public string Alias { get; init; } = "";
    public string TargetKey { get; init; } = "";
    public string Kind { get; init; } = "";
}

public sealed record CanonicalInputUnification
{
    public string CanonicalKey { get; init; } = "";
    public IReadOnlyList<string> EquivalentKeys { get; init; } = Array.Empty<string>();
    public string Reason { get; init; } = "";
}

public sealed record CanonicalInputLibrary
{
    public string Name { get; init; } = "";
    public string Guid { get; init; } = "";
    public string Version { get; init; } = "";
    public string? Key { get; init; }
    public IReadOnlyList<CanonicalInputType> Types { get; init; } = Array.Empty<CanonicalInputType>();
}

public sealed record CanonicalInputType
{
    public string Name { get; init; } = "";
    public string Kind { get; init; } = "";
    public string? Key { get; init; }
    public IReadOnlyList<string> BaseTypeKeys { get; init; } = Array.Empty<string>();
    public IReadOnlyList<CanonicalInputValue> Values { get; init; } = Array.Empty<CanonicalInputValue>();
    public IReadOnlyList<CanonicalInputMember> Members { get; init; } = Array.Empty<CanonicalInputMember>();
}

public sealed record CanonicalInputValue
{
    public string Name { get; init; } = "";
    public string Kind { get; init; } = "";
    public string Value { get; init; } = "";
    public string? ValueType { get; init; }
}

public sealed record CanonicalInputMember
{
    public string Name { get; init; } = "";
    public string Kind { get; init; } = "";
    public string? Key { get; init; }
    public int? DispId { get; init; }
    public string? ReturnType { get; init; }
    public IReadOnlyList<string> ParameterTypes { get; init; } = Array.Empty<string>();
    public IReadOnlyList<CanonicalInputParameter> Parameters { get; init; } = Array.Empty<CanonicalInputParameter>();
    public string? AccessorGroup { get; init; }
    public string? Value { get; init; }
    public string? AccessorKind { get; init; }
    public string? ValueType { get; init; }
}

public sealed record ConversionResult(DataGraph Graph, IReadOnlyList<ValidationIssue> Issues)
{
    public bool IsValid => Issues.Count == 0;
}

public static class CanonicalInputConverter
{
    private static readonly JsonSerializerOptions Options = new()
    {
        PropertyNamingPolicy = JsonNamingPolicy.CamelCase,
        DefaultIgnoreCondition = JsonIgnoreCondition.WhenWritingNull,
        UnmappedMemberHandling = JsonUnmappedMemberHandling.Disallow
    };

    public static CanonicalInput Parse(string json)
    {
        ArgumentNullException.ThrowIfNull(json);
        return JsonSerializer.Deserialize<CanonicalInput>(json, Options)
            ?? throw new InvalidDataException("Canonical input is empty.");
    }

    public static CanonicalInput Read(string path)
        => Parse(File.ReadAllText(path));

    public static ConversionResult Convert(CanonicalInput input)
    {
        ArgumentNullException.ThrowIfNull(input);
        var issues = new List<ValidationIssue>();
        if (!string.Equals(input.Format, "netoffice-data-input-v1", StringComparison.Ordinal))
            issues.Add(new("input.format", "Expected netoffice-data-input-v1.", "format"));

        var source = input.Source ?? new DataSource();
        if (string.IsNullOrWhiteSpace(source.Path))
            source = source with { Path = "canonical-input" };
        var sourceSha = source.Sha256;
        if (string.IsNullOrWhiteSpace(sourceSha))
            issues.Add(new("input.source.sha256", "Canonical input source.sha256 is required.", "source.sha256"));

        var libraries = new List<DataLibrary>();
        var types = new List<DataType>();
        var members = new List<DataMember>();
        var values = new List<DataValue>();
        var groups = new Dictionary<string, AccessorGroup>(StringComparer.Ordinal);
        var ambiguities = new List<AmbiguityRecord>();

        foreach (var inputLibrary in input.Libraries.OrderBy(static item => item.Key ?? item.Guid, StringComparer.Ordinal).ThenBy(static item => item.Name, StringComparer.Ordinal))
        {
            var libraryKey = inputLibrary.Key ?? inputLibrary.Guid;
            if (string.IsNullOrWhiteSpace(libraryKey))
            {
                issues.Add(new("input.library.key", "Library key or GUID is required.", inputLibrary.Name));
                continue;
            }

            var libraryId = LogicalIds.Library(inputLibrary.Guid, libraryKey);
            var libraryProvenance = ProvenanceFor(source, $"libraries/{libraryKey}");
            if (libraries.Any(item => item.LogicalId == libraryId))
            {
                AddAmbiguity(ambiguities, "duplicate-library", $"{inputLibrary.Guid}:{libraryKey}", [libraryId], libraryProvenance, "Duplicate library key.");
                continue;
            }
            libraries.Add(new DataLibrary
            {
                LogicalId = libraryId,
                Name = inputLibrary.Name,
                Guid = inputLibrary.Guid,
                Version = inputLibrary.Version,
                Provenance = libraryProvenance
            });

            var typeEntries = inputLibrary.Types
                .OrderBy(static item => item.Key ?? item.Name, StringComparer.Ordinal)
                .ThenBy(static item => item.Name, StringComparer.Ordinal)
                .ToArray();
            var typeIdsByKey = new Dictionary<string, string>(StringComparer.Ordinal);
            foreach (var inputType in typeEntries)
            {
                var typeKey = inputType.Key ?? inputType.Name;
                if (string.IsNullOrWhiteSpace(typeKey))
                {
                    issues.Add(new("input.type.key", "Type key or name is required.", libraryKey));
                    continue;
                }
                var typeId = LogicalIds.Type(libraryId, typeKey);
                if (!typeIdsByKey.TryAdd(typeKey, typeId))
                {
                    AddAmbiguity(ambiguities, "duplicate-type", $"{libraryId}:{typeKey}", [typeId], ProvenanceFor(source, $"libraries/{libraryKey}/types/{typeKey}"), "Duplicate type key.");
                    continue;
                }
                types.Add(new DataType
                {
                    LogicalId = typeId,
                    LibraryId = libraryId,
                    Name = inputType.Name,
                    Kind = inputType.Kind,
                    SourceKey = typeKey,
                    Provenance = ProvenanceFor(source, $"libraries/{libraryKey}/types/{typeKey}")
                });
            }

            foreach (var inputType in typeEntries)
            {
                var typeKey = inputType.Key ?? inputType.Name;
                if (!typeIdsByKey.TryGetValue(typeKey, out var typeId))
                    continue;
                var type = types.First(item => item.LogicalId == typeId);
                var baseIds = new List<string>();
                foreach (var baseKey in inputType.BaseTypeKeys.OrderBy(static key => key, StringComparer.Ordinal))
                {
                    if (typeIdsByKey.TryGetValue(baseKey, out var baseId))
                        baseIds.Add(baseId);
                    else
                        AddAmbiguity(ambiguities, "missing-base-type", $"{typeId}:{baseKey}", [], ProvenanceFor(source, $"libraries/{libraryKey}/types/{typeKey}"), $"Base type {baseKey} was not found.");
                }
                types[types.FindIndex(item => item.LogicalId == typeId)] = type with { BaseTypeIds = baseIds };
                foreach (var inputValue in inputType.Values.OrderBy(static item => item.Name, StringComparer.Ordinal))
                {
                    var valueKey = $"value:{inputValue.Name}";
                    var valueProvenance = ProvenanceFor(source, $"libraries/{libraryKey}/types/{typeKey}/values/{inputValue.Name}");
                    values.Add(new DataValue
                    {
                        LogicalId = LogicalIds.Member(typeId, valueKey),
                        TypeId = typeId,
                        Name = inputValue.Name,
                        Kind = inputValue.Kind,
                        Value = inputValue.Value,
                        ValueType = inputValue.ValueType,
                        Provenance = valueProvenance
                    });
                }

                var memberEntries = inputType.Members
                    .OrderBy(static item => item.Key ?? item.Name, StringComparer.Ordinal)
                    .ThenBy(static item => item.Name, StringComparer.Ordinal)
                    .ThenBy(static item => item.Kind, StringComparer.Ordinal)
                    .ToArray();
                var memberOrdinal = new Dictionary<string, int>(StringComparer.Ordinal);
                foreach (var inputMember in memberEntries)
                {
                    var memberKey = inputMember.Key ?? inputMember.Name;
                    if (string.IsNullOrWhiteSpace(memberKey))
                    {
                        issues.Add(new("input.member.key", "Member key or name is required.", typeKey));
                        continue;
                    }
                    memberOrdinal.TryGetValue(memberKey, out var ordinal);
                    memberOrdinal[memberKey] = ordinal + 1;
                    var identityKey = ordinal == 0 ? memberKey : $"{memberKey}#{ordinal + 1}";
                    var memberId = LogicalIds.Member(typeId, identityKey);
                    var accessorId = string.IsNullOrWhiteSpace(inputMember.AccessorGroup)
                        ? null
                        : LogicalIds.AccessorGroup(typeId, inputMember.AccessorGroup!);
                    var provenance = ProvenanceFor(source, $"libraries/{libraryKey}/types/{typeKey}/members/{memberKey}");
                    var parameters = inputMember.Parameters.Select((parameter, index) => new DataParameter
                    {
                        Name = parameter.Name,
                        Type = parameter.Type,
                        RefKind = parameter.RefKind,
                        IsOptional = parameter.IsOptional,
                        HasDefaultValue = parameter.DefaultValue is not null,
                        DefaultValue = parameter.DefaultValue,
                        Provenance = ProvenanceFor(source, $"{provenance.Location}/parameters/{index}")
                    }).ToArray();
                    if (inputMember.ParameterTypes.Count != 0 && parameters.Length == 0)
                        issues.Add(new("input.signature.parameters", "Parameter names/ref/default facts are required; parameterTypes alone cannot be emitted.", provenance.Location));
                    members.Add(new DataMember
                    {
                        LogicalId = memberId,
                        TypeId = typeId,
                        Name = inputMember.Name,
                        Kind = inputMember.Kind,
                        SourceKey = memberKey,
                        DispId = inputMember.DispId,
                        ReturnType = inputMember.ReturnType,
                        ParameterTypes = inputMember.ParameterTypes.Count == 0 ? parameters.Select(static item => item.Type).ToArray() : inputMember.ParameterTypes,
                        Parameters = parameters,
                        Value = inputMember.Value,
                        ValueType = inputMember.ValueType,
                        AccessorGroupId = accessorId,
                        AccessorKind = inputMember.AccessorKind,
                        Provenance = provenance
                    });
                    if (ordinal > 0)
                        AddAmbiguity(ambiguities, "duplicate-member", $"{typeId}:{memberKey}", members.Where(item => item.TypeId == typeId && item.SourceKey == memberKey).Select(static item => item.LogicalId).ToArray(), provenance, "Duplicate member key retained with deterministic ordinal.");
                    if (accessorId is not null)
                    {
                        if (!groups.TryGetValue(accessorId, out var group))
                            group = new AccessorGroup { LogicalId = accessorId, TypeId = typeId, Name = inputMember.AccessorGroup!, Kind = inputMember.Kind, Provenance = provenance };
                        groups[accessorId] = group with { MemberIds = group.MemberIds.Concat([memberId]).ToArray() };
                    }
                }
            }
        }
        var aliases = new List<AliasRecord>();
        var unifications = new List<UnificationRecord>();
        var knownEntityIds = libraries.Select(static item => item.LogicalId)
            .Concat(types.Select(static item => item.LogicalId))
            .Concat(members.Select(static item => item.LogicalId))
            .Concat(values.Select(static item => item.LogicalId))
            .ToHashSet(StringComparer.Ordinal);
        foreach (var inputAlias in input.Aliases.OrderBy(static item => item.Alias, StringComparer.Ordinal))
        {
            var targetId = knownEntityIds.Contains(inputAlias.TargetKey)
                ? inputAlias.TargetKey
                : types.FirstOrDefault(item => item.SourceKey == inputAlias.TargetKey)?.LogicalId
                    ?? members.FirstOrDefault(item => item.SourceKey == inputAlias.TargetKey)?.LogicalId
                    ?? libraries.FirstOrDefault(item => item.Name == inputAlias.TargetKey)?.LogicalId;
            var provenance = ProvenanceFor(source, $"aliases/{inputAlias.Alias}");
            if (targetId is null)
            {
                AddAmbiguity(ambiguities, "missing-alias-target", inputAlias.Alias, [], provenance, $"Alias target {inputAlias.TargetKey} was not found.");
                continue;
            }
            aliases.Add(new AliasRecord { LogicalId = LogicalIds.Alias(inputAlias.Alias), Alias = inputAlias.Alias, TargetId = targetId, Kind = inputAlias.Kind, Provenance = provenance });
        }
        foreach (var inputUnification in input.Unifications.OrderBy(static item => item.CanonicalKey, StringComparer.Ordinal))
        {
            var canonicalId = knownEntityIds.Contains(inputUnification.CanonicalKey)
                ? inputUnification.CanonicalKey
                : types.FirstOrDefault(item => item.SourceKey == inputUnification.CanonicalKey)?.LogicalId
                    ?? members.FirstOrDefault(item => item.SourceKey == inputUnification.CanonicalKey)?.LogicalId;
            var provenance = ProvenanceFor(source, $"unifications/{inputUnification.CanonicalKey}");
            if (canonicalId is null)
            {
                AddAmbiguity(ambiguities, "missing-unification-target", inputUnification.CanonicalKey, [], provenance, "Unification canonical identity was not found.");
                continue;
            }
            var equivalentIds = inputUnification.EquivalentKeys.Select(key => knownEntityIds.Contains(key) ? key : members.FirstOrDefault(item => item.SourceKey == key)?.LogicalId ?? types.FirstOrDefault(item => item.SourceKey == key)?.LogicalId).Where(static id => id is not null).Cast<string>().ToArray();
            unifications.Add(new UnificationRecord { LogicalId = LogicalIds.Unification(inputUnification.CanonicalKey), CanonicalId = canonicalId, EquivalentIds = equivalentIds, Reason = inputUnification.Reason, Provenance = provenance });
        }

        var graph = new DataGraph
        {
            Source = source,
            Libraries = libraries,
            Types = types,
            Members = members,
            Values = values,
            AccessorGroups = groups.Values.ToArray(),
            Aliases = aliases,
            Unifications = unifications,
            Ambiguities = ambiguities
        };
        graph = graph with { Digest = CanonicalJson.ComputeDigest(graph) };
        return new ConversionResult(graph, issues);
    }

    public static ConversionResult ConvertJson(string json) => Convert(Parse(json));

    private static Provenance ProvenanceFor(DataSource source, string location)
        => new() { SourcePath = source.Path, SourceSha256 = source.Sha256, SourceRevision = source.Revision, Location = location };

    private static void AddAmbiguity(ICollection<AmbiguityRecord> records, string kind, string key, IReadOnlyList<string> candidates, Provenance provenance, string reason)
        => records.Add(new AmbiguityRecord
        {
            LogicalId = LogicalIds.Ambiguity(kind, key),
            Kind = kind,
            Reason = reason,
            Candidates = candidates,
            Provenance = provenance
        });
}
