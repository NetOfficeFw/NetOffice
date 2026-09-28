using System.Security.Cryptography;
using System.Text;
using System.Text.Json;
using System.Text.Json.Serialization;

namespace NetOffice.CodeGen.Data;

public static class CanonicalJson
{
    private static readonly JsonSerializerOptions Options = new()
    {
        Encoder = System.Text.Encodings.Web.JavaScriptEncoder.Default,
        PropertyNamingPolicy = JsonNamingPolicy.CamelCase,
        DefaultIgnoreCondition = JsonIgnoreCondition.WhenWritingNull,
        UnmappedMemberHandling = JsonUnmappedMemberHandling.Disallow,
        WriteIndented = false
    };

    public static string Serialize(DataGraph graph)
    {
        ArgumentNullException.ThrowIfNull(graph);
        var normalized = Normalize(graph);
        var digest = ComputeDigestNormalized(normalized);
        return JsonSerializer.Serialize(ToDocument(normalized with { Digest = digest }), Options);
    }

    public static byte[] SerializeUtf8(DataGraph graph)
    {
        ArgumentNullException.ThrowIfNull(graph);
        var normalized = Normalize(graph);
        var digest = ComputeDigestNormalized(normalized);
        return JsonSerializer.SerializeToUtf8Bytes(ToDocument(normalized with { Digest = digest }), Options);
    }

    public static DataGraph Parse(string json)
    {
        ArgumentNullException.ThrowIfNull(json);
        var graph = JsonSerializer.Deserialize<DataGraph>(json, Options)
            ?? throw new InvalidDataException("The data graph JSON is empty.");
        return Normalize(graph);
    }

    public static DataGraph Parse(ReadOnlySpan<byte> utf8Json)
    {
        var graph = JsonSerializer.Deserialize<DataGraph>(utf8Json, Options)
            ?? throw new InvalidDataException("The data graph JSON is empty.");
        return Normalize(graph);
    }

    public static DataGraph Read(Stream stream)
    {
        ArgumentNullException.ThrowIfNull(stream);
        var graph = JsonSerializer.Deserialize<DataGraph>(stream, Options)
            ?? throw new InvalidDataException("The data graph JSON is empty.");
        return Normalize(graph);
    }

    public static DataGraph Read(string path)
    {
        ArgumentException.ThrowIfNullOrWhiteSpace(path);
        using var stream = new FileStream(Path.GetFullPath(path), FileMode.Open, FileAccess.Read, FileShare.Read, 1024 * 1024, FileOptions.SequentialScan);
        return Read(stream);
    }

    public static void Write(string path, DataGraph graph)
    {
        ArgumentException.ThrowIfNullOrWhiteSpace(path);
        ArgumentNullException.ThrowIfNull(graph);
        var normalized = Normalize(graph);
        var digest = ComputeDigestNormalized(normalized);
        WriteDocument(path, ToDocument(normalized with { Digest = digest }));
    }

    public static void WritePrecomputed(string path, DataGraph graph)
    {
        ArgumentException.ThrowIfNullOrWhiteSpace(path);
        ArgumentNullException.ThrowIfNull(graph);
        if (graph.Digest.Length != 64 || graph.Digest.Any(static character => !Uri.IsHexDigit(character)))
            throw new InvalidDataException("A precomputed SHA-256 graph digest is required.");
        WriteDocument(path, ToDocument(Normalize(graph)));
    }

    public static string ComputeDigest(DataGraph graph)
    {
        ArgumentNullException.ThrowIfNull(graph);
        return ComputeDigestNormalized(Normalize(graph));
    }

    private static string ComputeDigestNormalized(DataGraph normalized)
    {
        using var algorithm = SHA256.Create();
        using var stream = new CryptoStream(Stream.Null, algorithm, CryptoStreamMode.Write);
        JsonSerializer.Serialize(stream, ToPayload(normalized), Options);
        stream.FlushFinalBlock();
        return Convert.ToHexString(algorithm.Hash!).ToLowerInvariant();
    }

    private static void WriteDocument(string path, GraphDocument document)
    {
        var fullPath = Path.GetFullPath(path);
        Directory.CreateDirectory(Path.GetDirectoryName(fullPath)!);
        using var stream = new FileStream(fullPath, FileMode.Create, FileAccess.Write, FileShare.None, 1024 * 1024, FileOptions.SequentialScan);
        JsonSerializer.Serialize(stream, document, Options);
    }

    public static DataGraph Normalize(DataGraph graph)
    {
        ArgumentNullException.ThrowIfNull(graph);
        return graph with
        {
            Projects = graph.Projects.Select(static item => item with
            {
                SourceCategories = item.SourceCategories.OrderBy(static value => value, StringComparer.Ordinal).ToArray(),
                LibraryIds = item.LibraryIds.OrderBy(static id => id, StringComparer.Ordinal).ToArray(),
                ReferenceProjectIds = item.ReferenceProjectIds.OrderBy(static id => id, StringComparer.Ordinal).ToArray()
            }).OrderBy(static item => item.LogicalId, StringComparer.Ordinal).ToArray(),
            Libraries = graph.Libraries.Select(static item => item with
            {
                Dependencies = item.Dependencies.Select(static dependency => dependency with
                {
                    TargetLibraryIds = dependency.TargetLibraryIds.OrderBy(static id => id, StringComparer.Ordinal).ToArray()
                }).OrderBy(static dependency => dependency.Guid, StringComparer.Ordinal)
                  .ThenBy(static dependency => dependency.Major, StringComparer.Ordinal)
                  .ThenBy(static dependency => dependency.Minor, StringComparer.Ordinal)
                  .ThenBy(static dependency => dependency.Name, StringComparer.Ordinal)
                  .ToArray()
            }).OrderBy(static item => item.LogicalId, StringComparer.Ordinal).ToArray(),
            Types = graph.Types.Select(static item => item with
            {
                GuidObservations = NormalizeIdentifiers(item.GuidObservations),
                ReferenceObservations = item.ReferenceObservations.Select(static observation => observation with
                {
                    LibraryIds = observation.LibraryIds.OrderBy(static id => id, StringComparer.Ordinal).ToArray()
                }).OrderBy(static observation => observation.Kind, StringComparer.Ordinal)
                  .ThenBy(static observation => observation.TargetTypeId, StringComparer.Ordinal)
                  .ThenBy(static observation => string.Join("\0", observation.LibraryIds), StringComparer.Ordinal)
                  .ToArray(),
                BaseTypeIds = item.BaseTypeIds.OrderBy(static id => id, StringComparer.Ordinal).ToArray(),
                DefaultInterfaceIds = item.DefaultInterfaceIds.OrderBy(static id => id, StringComparer.Ordinal).ToArray(),
                EventInterfaceIds = item.EventInterfaceIds.OrderBy(static id => id, StringComparer.Ordinal).ToArray(),
                SupportObservationIds = item.SupportObservationIds.OrderBy(static id => id, StringComparer.Ordinal).ToArray()
            }).OrderBy(static item => item.LogicalId, StringComparer.Ordinal).ToArray(),
            Members = graph.Members.Select(static item => item with
            {
                DispIdObservations = NormalizeIdentifiers(item.DispIdObservations),
                ParameterTypes = item.ParameterTypes.ToArray(),
                Parameters = item.Parameters.ToArray(),
                SupportObservationIds = item.SupportObservationIds.OrderBy(static id => id, StringComparer.Ordinal).ToArray(),
                SignatureLibraryIds = item.SignatureLibraryIds.OrderBy(static id => id, StringComparer.Ordinal).ToArray()
            }).OrderBy(static item => item.LogicalId, StringComparer.Ordinal).ToArray(),
            Values = graph.Values.Select(static item => item with
            {
                SupportObservationIds = item.SupportObservationIds.OrderBy(static id => id, StringComparer.Ordinal).ToArray()
            }).OrderBy(static item => item.LogicalId, StringComparer.Ordinal).ToArray(),
            AccessorGroups = graph.AccessorGroups.Select(static item => item with
            {
                MemberIds = item.MemberIds.OrderBy(static id => id, StringComparer.Ordinal).ToArray()
            }).OrderBy(static item => item.LogicalId, StringComparer.Ordinal).ToArray(),
            SupportObservations = graph.SupportObservations.Select(static item => item with
            {
                Versions = item.Versions.OrderBy(static version => version, StringComparer.Ordinal).ToArray(),
                LibraryIds = item.LibraryIds.OrderBy(static id => id, StringComparer.Ordinal).ToArray()
            }).OrderBy(static item => item.LogicalId, StringComparer.Ordinal).ToArray(),
            Aliases = graph.Aliases.OrderBy(static item => item.LogicalId, StringComparer.Ordinal).ToArray(),
            Unifications = graph.Unifications.Select(static item => item with
            {
                EquivalentIds = item.EquivalentIds.OrderBy(static id => id, StringComparer.Ordinal).ToArray()
            }).OrderBy(static item => item.LogicalId, StringComparer.Ordinal).ToArray(),
            InvocationEvidence = graph.InvocationEvidence.OrderBy(static item => item.LogicalId, StringComparer.Ordinal).ToArray(),
            AccessorEvidence = graph.AccessorEvidence.Select(static item => item with
            {
                MemberIds = item.MemberIds.OrderBy(static id => id, StringComparer.Ordinal).ToArray()
            }).OrderBy(static item => item.LogicalId, StringComparer.Ordinal).ToArray(),
            AbsentFacts = graph.AbsentFacts.OrderBy(static item => item.LogicalId, StringComparer.Ordinal).ToArray(),
            Unknowns = graph.Unknowns.OrderBy(static item => item.LogicalId, StringComparer.Ordinal).ToArray(),
            StaleRecords = graph.StaleRecords.OrderBy(static item => item.LogicalId, StringComparer.Ordinal).ToArray(),
            Ambiguities = graph.Ambiguities.Select(static item => item with
            {
                Candidates = item.Candidates.OrderBy(static id => id, StringComparer.Ordinal).ToArray()
            }).OrderBy(static item => item.LogicalId, StringComparer.Ordinal).ToArray()
        };
    }
    private static IReadOnlyList<DataIdentifierObservation> NormalizeIdentifiers(IEnumerable<DataIdentifierObservation> observations)
        => observations.Select(static observation => observation with
        {
            LibraryIds = observation.LibraryIds.OrderBy(static id => id, StringComparer.Ordinal).ToArray()
        }).OrderBy(static observation => observation.Value, StringComparer.Ordinal)
          .ThenBy(static observation => string.Join("\0", observation.LibraryIds), StringComparer.Ordinal)
          .ToArray();


    private static GraphDocument ToDocument(DataGraph graph) => new()
    {
        SchemaVersion = graph.SchemaVersion,
        Serialization = graph.Serialization,
        Source = graph.Source,
        Projects = graph.Projects,
        Libraries = graph.Libraries,
        Types = graph.Types,
        Members = graph.Members,
        Values = graph.Values,
        AccessorGroups = graph.AccessorGroups,
        SupportObservations = graph.SupportObservations,
        Aliases = graph.Aliases,
        Unifications = graph.Unifications,
        InvocationEvidence = graph.InvocationEvidence,
        AccessorEvidence = graph.AccessorEvidence,
        Unknowns = graph.Unknowns,
        AbsentFacts = graph.AbsentFacts,
        StaleRecords = graph.StaleRecords,
        Ambiguities = graph.Ambiguities,
        Digest = graph.Digest
    };

    private static GraphPayload ToPayload(DataGraph graph) => new()
    {
        SchemaVersion = graph.SchemaVersion,
        Serialization = graph.Serialization,
        Source = graph.Source,
        Projects = graph.Projects,
        Libraries = graph.Libraries,
        Types = graph.Types,
        Members = graph.Members,
        Values = graph.Values,
        AccessorGroups = graph.AccessorGroups,
        SupportObservations = graph.SupportObservations,
        Aliases = graph.Aliases,
        Unifications = graph.Unifications,
        InvocationEvidence = graph.InvocationEvidence,
        AccessorEvidence = graph.AccessorEvidence,
        Unknowns = graph.Unknowns,
        AbsentFacts = graph.AbsentFacts,
        StaleRecords = graph.StaleRecords,
        Ambiguities = graph.Ambiguities
    };

    private class GraphPayload
    {
        public string SchemaVersion { get; init; } = "";
        public string Serialization { get; init; } = "";
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
        public IReadOnlyList<UnknownRecord> Unknowns { get; init; } = Array.Empty<UnknownRecord>();
        public IReadOnlyList<AbsentFactRecord> AbsentFacts { get; init; } = Array.Empty<AbsentFactRecord>();
        public IReadOnlyList<StaleRecord> StaleRecords { get; init; } = Array.Empty<StaleRecord>();
        public IReadOnlyList<AmbiguityRecord> Ambiguities { get; init; } = Array.Empty<AmbiguityRecord>();
    }
    private sealed class GraphDocument : GraphPayload
    {
        public string Digest { get; init; } = "";
    }

}
