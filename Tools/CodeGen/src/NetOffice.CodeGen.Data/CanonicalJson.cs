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
        WriteIndented = false
    };

    public static string Serialize(DataGraph graph)
    {
        ArgumentNullException.ThrowIfNull(graph);
        var normalized = Normalize(graph);
        var digest = ComputeDigest(normalized);
        return JsonSerializer.Serialize(ToDocument(normalized with { Digest = digest }), Options);
    }

    public static byte[] SerializeUtf8(DataGraph graph) => Encoding.UTF8.GetBytes(Serialize(graph));

    public static DataGraph Parse(string json)
    {
        ArgumentNullException.ThrowIfNull(json);
        var graph = JsonSerializer.Deserialize<DataGraph>(json, Options)
            ?? throw new InvalidDataException("The data graph JSON is empty.");
        return Normalize(graph);
    }

    public static DataGraph Read(string path)
    {
        ArgumentException.ThrowIfNullOrWhiteSpace(path);
        return Parse(File.ReadAllText(path, Encoding.UTF8));
    }

    public static void Write(string path, DataGraph graph)
    {
        ArgumentException.ThrowIfNullOrWhiteSpace(path);
        ArgumentNullException.ThrowIfNull(graph);
        var fullPath = Path.GetFullPath(path);
        Directory.CreateDirectory(Path.GetDirectoryName(fullPath)!);
        File.WriteAllBytes(fullPath, SerializeUtf8(graph));
    }

    public static string ComputeDigest(DataGraph graph)
    {
        ArgumentNullException.ThrowIfNull(graph);
        var normalized = Normalize(graph);
        var payload = JsonSerializer.Serialize(ToPayload(normalized), Options);
        return Convert.ToHexString(SHA256.HashData(Encoding.UTF8.GetBytes(payload))).ToLowerInvariant();
    }

    public static DataGraph Normalize(DataGraph graph)
    {
        ArgumentNullException.ThrowIfNull(graph);
        return graph with
        {
            Libraries = graph.Libraries.OrderBy(static item => item.LogicalId, StringComparer.Ordinal).ToArray(),
            Types = graph.Types.Select(static item => item with
            {
                BaseTypeIds = item.BaseTypeIds.OrderBy(static id => id, StringComparer.Ordinal).ToArray()
            }).OrderBy(static item => item.LogicalId, StringComparer.Ordinal).ToArray(),
            Members = graph.Members.Select(static item => item with
            {
                ParameterTypes = item.ParameterTypes.ToArray()
            }).OrderBy(static item => item.LogicalId, StringComparer.Ordinal).ToArray(),
            AccessorGroups = graph.AccessorGroups.Select(static item => item with
            {
                MemberIds = item.MemberIds.OrderBy(static id => id, StringComparer.Ordinal).ToArray()
            }).OrderBy(static item => item.LogicalId, StringComparer.Ordinal).ToArray(),
            Ambiguities = graph.Ambiguities.Select(static item => item with
            {
                Candidates = item.Candidates.OrderBy(static id => id, StringComparer.Ordinal).ToArray()
            }).OrderBy(static item => item.LogicalId, StringComparer.Ordinal).ToArray()
        };
    }

    private static GraphDocument ToDocument(DataGraph graph) => new()
    {
        SchemaVersion = graph.SchemaVersion,
        Serialization = graph.Serialization,
        Source = graph.Source,
        Libraries = graph.Libraries,
        Types = graph.Types,
        Members = graph.Members,
        AccessorGroups = graph.AccessorGroups,
        Ambiguities = graph.Ambiguities,
        Digest = graph.Digest
    };

    private static GraphPayload ToPayload(DataGraph graph) => new()
    {
        SchemaVersion = graph.SchemaVersion,
        Serialization = graph.Serialization,
        Source = graph.Source,
        Libraries = graph.Libraries,
        Types = graph.Types,
        Members = graph.Members,
        AccessorGroups = graph.AccessorGroups,
        Ambiguities = graph.Ambiguities
    };

    private class GraphPayload
    {
        public string SchemaVersion { get; init; } = "";
        public string Serialization { get; init; } = "";
        public DataSource Source { get; init; } = new();
        public IReadOnlyList<DataLibrary> Libraries { get; init; } = Array.Empty<DataLibrary>();
        public IReadOnlyList<DataType> Types { get; init; } = Array.Empty<DataType>();
        public IReadOnlyList<DataMember> Members { get; init; } = Array.Empty<DataMember>();
        public IReadOnlyList<AccessorGroup> AccessorGroups { get; init; } = Array.Empty<AccessorGroup>();
        public IReadOnlyList<AmbiguityRecord> Ambiguities { get; init; } = Array.Empty<AmbiguityRecord>();
    }

    private sealed class GraphDocument : GraphPayload
    {
        public string Digest { get; init; } = "";
    }
}
