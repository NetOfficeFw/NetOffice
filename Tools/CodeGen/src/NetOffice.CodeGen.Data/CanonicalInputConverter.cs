using System.Text.Json;
using System.Text.Json.Serialization;

namespace NetOffice.CodeGen.Data;

public sealed record CanonicalInput
{
    public string Format { get; init; } = "netoffice-data-input-v1";
    public DataSource Source { get; init; } = new();
    public IReadOnlyList<CanonicalInputLibrary> Libraries { get; init; } = Array.Empty<CanonicalInputLibrary>();
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
    public IReadOnlyList<CanonicalInputMember> Members { get; init; } = Array.Empty<CanonicalInputMember>();
}

public sealed record CanonicalInputMember
{
    public string Name { get; init; } = "";
    public string Kind { get; init; } = "";
    public string? Key { get; init; }
    public int? DispId { get; init; }
    public string? ReturnType { get; init; }
    public IReadOnlyList<string> ParameterTypes { get; init; } = Array.Empty<string>();
    public string? AccessorGroup { get; init; }
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
        DefaultIgnoreCondition = JsonIgnoreCondition.WhenWritingNull
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
                    members.Add(new DataMember
                    {
                        LogicalId = memberId,
                        TypeId = typeId,
                        Name = inputMember.Name,
                        Kind = inputMember.Kind,
                        SourceKey = memberKey,
                        DispId = inputMember.DispId,
                        ReturnType = inputMember.ReturnType,
                        ParameterTypes = inputMember.ParameterTypes,
                        AccessorGroupId = accessorId,
                        Provenance = provenance
                    });
                    if (ordinal > 0)
                        AddAmbiguity(ambiguities, "duplicate-member", $"{typeId}:{memberKey}", members.Where(item => item.TypeId == typeId && item.SourceKey == memberKey).Select(static item => item.LogicalId).ToArray(), provenance, "Duplicate member key retained with deterministic ordinal.");
                    if (accessorId is not null)
                    {
                        if (!groups.TryGetValue(accessorId, out var group))
                            group = new AccessorGroup { LogicalId = accessorId, TypeId = typeId, Name = inputMember.AccessorGroup! };
                        groups[accessorId] = group with { MemberIds = group.MemberIds.Concat([memberId]).ToArray() };
                    }
                }
            }
        }

        var graph = new DataGraph
        {
            Source = source,
            Libraries = libraries,
            Types = types,
            Members = members,
            AccessorGroups = groups.Values.ToArray(),
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
