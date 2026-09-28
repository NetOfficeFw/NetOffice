using System.Collections.ObjectModel;
using System.Text.Json;
using System.Text.Json.Serialization;
using System.Xml.Linq;

namespace NetOffice.CodeGen.Docs;

public enum MappingKind { Exact, ApprovedAlias, Manual, Unmatched, Ambiguous }
public enum DocsProfile { Baseline, Vba }
public enum DocumentationTargetKind { Type, Member, BindingKey }

public sealed record DocsPin(string Repository, string Commit, string? License = "CC BY 4.0")
{
    public static DocsPin Load(string path)
    {
        if (!File.Exists(path)) throw new FileNotFoundException($"Documentation pin is missing: {path}", path);
        var pin = JsonSerializer.Deserialize<DocsPin>(File.ReadAllText(path), JsonOptions.Default)
            ?? throw new InvalidDataException("Documentation pin is empty.");
        if (string.IsNullOrWhiteSpace(pin.Repository)
            || !pin.Repository.Equals("MicrosoftDocs/VBA-Docs", StringComparison.OrdinalIgnoreCase)
            || string.IsNullOrWhiteSpace(pin.Commit)
            || pin.Commit.Length != 40
            || !pin.Commit.All(Uri.IsHexDigit))
            throw new InvalidDataException("Documentation pin must target MicrosoftDocs/VBA-Docs and contain a 40-character commit SHA.");
        return pin;
    }
}

/// <summary>A projected declaration to which documentation will be attached.</summary>
public sealed record DocumentationTarget
{
    /// <summary>The stable Data v2 logical ID.</summary>
    public string LogicalId { get; init; } = "";
    /// <summary>The projection binding key used by emission.</summary>
    public string BindingKey { get; init; } = "";
    public DocumentationTargetKind Kind { get; init; }
    public string ContractTypeLogicalId { get; init; } = "";
    public string? ContractTypeName { get; init; }
    public string? ContractTypeNamespace { get; init; }
    public string? ContractSourcePath { get; init; }
    public string Name { get; init; } = "";
    public string? ContractMemberKind { get; init; }
    public string? ContractSignature { get; init; }
    /// <summary>ValueText names in emitted parameter order. Null means that the caller has no signature information.</summary>
    public IReadOnlyList<string>? EmittedParameterNames { get; init; }

    public static DocumentationTarget ForType(string logicalId, string bindingKey, string contractTypeLogicalId, string name, string? contractTypeNamespace = null, string? contractSourcePath = null)
        => new() { LogicalId = logicalId, BindingKey = bindingKey, Kind = DocumentationTargetKind.Type, ContractTypeLogicalId = contractTypeLogicalId, ContractTypeName = name, ContractTypeNamespace = contractTypeNamespace, ContractSourcePath = contractSourcePath, Name = name };

    public static DocumentationTarget ForMember(string logicalId, string bindingKey, string contractTypeLogicalId, string name,
        string memberKind, string contractSignature, IEnumerable<string> emittedParameterNames, string? contractTypeName = null, string? contractTypeNamespace = null, string? contractSourcePath = null)
        => new()
        {
            LogicalId = logicalId,
            BindingKey = bindingKey,
            Kind = DocumentationTargetKind.Member,
            ContractTypeLogicalId = contractTypeLogicalId,
            ContractTypeName = contractTypeName,
            ContractTypeNamespace = contractTypeNamespace,
            ContractSourcePath = contractSourcePath,
            Name = name,
            ContractMemberKind = memberKind,
            ContractSignature = contractSignature,
            EmittedParameterNames = emittedParameterNames.Select(ParameterName).ToArray()
        };

    internal static DocumentationTarget Legacy(string id)
        => new() { LogicalId = id, BindingKey = id, Kind = DocumentationTargetKind.BindingKey, Name = id };

    internal void Validate()
    {
        if (string.IsNullOrWhiteSpace(LogicalId)) throw new ArgumentException("A documentation target logical ID is required.");
        if (string.IsNullOrWhiteSpace(BindingKey)) throw new ArgumentException($"Documentation target '{LogicalId}' has no binding key.");
        if (Kind != DocumentationTargetKind.BindingKey && string.IsNullOrWhiteSpace(ContractTypeLogicalId))
            throw new ArgumentException($"Documentation target '{LogicalId}' has no Wrapper Contract type logical ID.");
        if (Kind == DocumentationTargetKind.Member && (string.IsNullOrWhiteSpace(Name) || string.IsNullOrWhiteSpace(ContractMemberKind)))
            throw new ArgumentException($"Documentation member target '{LogicalId}' is incomplete.");
        if (EmittedParameterNames is { } names && names.GroupBy(ParameterName, StringComparer.Ordinal).Any(static x => x.Count() != 1))
            throw new ArgumentException($"Documentation target '{LogicalId}' has duplicate emitted parameter names.");
    }

    internal static string ParameterName(string value)
    {
        var result = value?.Trim() ?? "";
        return result.StartsWith('@') ? result[1..] : result;
    }
}

public sealed record DocumentationArticle(string Id, string Title, string RelativePath, string Markdown, string CanonicalUrl);

public sealed record MappingRecord(
    string LogicalId,
    string BindingKey,
    MappingKind Kind,
    string? ArticleId,
    string? CanonicalUrl,
    string? ContractKey = null,
    string? Reason = null,
    IReadOnlyList<string>? Candidates = null);

public sealed record XmlDocumentationElement(string Name, IReadOnlyDictionary<string, string> Attributes, string Xml);

/// <summary>Contract XML documentation. RawXml is always preserved, including source records whose XML is malformed.</summary>
public sealed record XmlDocumentation(
    string RawXml,
    string? Summary,
    string? Remarks,
    string? Returns,
    string? Value,
    IReadOnlyDictionary<string, string> Parameters,
    IReadOnlyList<XmlDocumentationElement> Elements,
    string ParseStatus = "parsed",
    string? ParseError = null)
{
    public static XmlDocumentation Parse(string rawXml)
    {
        if (rawXml is null) throw new ArgumentNullException(nameof(rawXml));
        XElement root;
        try { root = XElement.Parse("<root>" + rawXml + "</root>", LoadOptions.PreserveWhitespace); }
        catch (Exception error) when (error is System.Xml.XmlException or ArgumentException)
        { throw new InvalidDataException("Malformed XML documentation fragment.", error); }

        var parameters = new Dictionary<string, string>(StringComparer.Ordinal);
        foreach (var parameter in root.Descendants().Where(static x => x.Name.LocalName == "param"))
        {
            var name = (string?)parameter.Attribute("name");
            if (string.IsNullOrWhiteSpace(name)) throw new InvalidDataException("XML documentation contains a <param> without a name.");
            name = DocumentationTarget.ParameterName(name);
            if (!parameters.TryAdd(name, NormalizeText(parameter.Value))) throw new InvalidDataException($"XML documentation contains duplicate <param> elements for '{name}'.");
        }

        var elements = root.Elements().Select(static element => new XmlDocumentationElement(
            element.Name.LocalName,
            new ReadOnlyDictionary<string, string>(element.Attributes().OrderBy(static x => x.Name.ToString(), StringComparer.Ordinal).ToDictionary(static x => x.Name.ToString(), static x => x.Value, StringComparer.Ordinal)),
            element.ToString(SaveOptions.DisableFormatting))).ToArray();
        static string? Text(XElement parent, string name)
        {
            var element = parent.Elements().FirstOrDefault(x => x.Name.LocalName == name);
            return element is null ? null : NormalizeText(element.Value);
        }
        return new XmlDocumentation(rawXml, Text(root, "summary"), Text(root, "remarks"), Text(root, "returns"), Text(root, "value"),
            new ReadOnlyDictionary<string, string>(parameters), elements);
    }

    internal static XmlDocumentation Invalid(string rawXml, string? parseError, string? summary, string? remarks, string? returns, string? value, IReadOnlyDictionary<string, string> parameters)
        => new(rawXml, summary, remarks, returns, value, parameters, Array.Empty<XmlDocumentationElement>(), "invalid", parseError ?? "Malformed XML documentation.");

    private static string NormalizeText(string value)
        => string.Join(" ", value.Split((char[]?)null, StringSplitOptions.RemoveEmptyEntries));
}

public sealed record BoundDocumentation(
    string LogicalId,
    string BindingKey,
    MappingKind MappingKind,
    string? ContractKey,
    XmlDocumentation Documentation);

public sealed record DocumentationResult(
    IReadOnlyList<MappingRecord> Mappings,
    IReadOnlyDictionary<string, BoundDocumentation> Documents,
    string Digest,
    DocsPin? Pin,
    DocsProfile Profile)
{
    public IReadOnlyDictionary<string, string> XmlDocs { get; } = new RawDocumentationDictionary(Documents);
}

internal sealed class RawDocumentationDictionary : IReadOnlyDictionary<string, string>
{
    private readonly IReadOnlyDictionary<string, BoundDocumentation> _documents;

    public RawDocumentationDictionary(IReadOnlyDictionary<string, BoundDocumentation> documents) => _documents = documents;
    public int Count => _documents.Count;
    public IEnumerable<string> Keys => _documents.Keys;
    public IEnumerable<string> Values => _documents.Values.Select(static value => value.Documentation.RawXml);
    public string this[string key] => _documents[key].Documentation.RawXml;
    public bool ContainsKey(string key) => _documents.ContainsKey(key);

    public bool TryGetValue(string key, out string value)
    {
        if (_documents.TryGetValue(key, out var document))
        {
            value = document.Documentation.RawXml;
            return true;
        }
        value = "";
        return false;
    }

    public IEnumerator<KeyValuePair<string, string>> GetEnumerator()
    {
        foreach (var document in _documents)
            yield return new KeyValuePair<string, string>(document.Key, document.Value.Documentation.RawXml);
    }

    System.Collections.IEnumerator System.Collections.IEnumerable.GetEnumerator() => GetEnumerator();
}

public sealed record DocsOptions(
    string SourceDirectory,
    string? PinPath = null,
    DocsProfile Profile = DocsProfile.Baseline,
    bool Locked = false,
    string? ManualMappingsPath = null,
    string? AliasMappingsPath = null,
    string? BaselineContractPath = null,
    string? ReportDirectory = null)
{
    public void Validate()
    {
        if (!Enum.IsDefined(Profile)) throw new ArgumentException("Unknown documentation profile.", nameof(Profile));
        if (Profile == DocsProfile.Baseline && string.IsNullOrWhiteSpace(BaselineContractPath))
            throw new InvalidOperationException("The baseline documentation profile requires a Wrapper Contract file or directory.");
        if (Profile == DocsProfile.Vba && Locked && string.IsNullOrWhiteSpace(PinPath))
            throw new InvalidOperationException("Locked VBA documentation requires an explicit pin.");
    }
}

public sealed class DocumentationMappingException : Exception
{
    public DocumentationMappingException(DocumentationResult result)
        : base($"Ambiguous documentation mappings are not allowed: {string.Join(", ", result.Mappings.Where(static x => x.Kind == MappingKind.Ambiguous).Select(static x => x.LogicalId).Order(StringComparer.Ordinal))}")
        => Result = result;

    public DocumentationResult Result { get; }
}

internal static class JsonOptions
{
    public static readonly JsonSerializerOptions Default = Create(false);
    public static readonly JsonSerializerOptions Indented = Create(true);

    private static JsonSerializerOptions Create(bool indented)
    {
        var options = new JsonSerializerOptions { PropertyNameCaseInsensitive = true, PropertyNamingPolicy = JsonNamingPolicy.CamelCase, WriteIndented = indented };
        options.Converters.Add(new JsonStringEnumConverter(JsonNamingPolicy.CamelCase));
        return options;
    }
}
