using System.Security.Cryptography;
using System.Text;
using System.Text.Json;
using System.Text.RegularExpressions;
using System.Xml.Linq;

namespace NetOffice.CodeGen.Docs;

/// <summary>Versioned, deterministic index of current-source XML documentation from Wrapper Contracts.</summary>
public sealed class BaselineContractCorpus
{
    private readonly IReadOnlyDictionary<string, BaselineTypeContract> _types;
    private readonly Dictionary<string, BaselineTypeContract[]> _typesByNormalizedId;
    private readonly Dictionary<string, BaselineTypeContract[]> _typesByNormalizedName;
    private readonly Dictionary<(string Owner, string Name), BaselineMemberContract[]> _membersByName;
    private readonly Dictionary<(string Owner, string Name, string Kind), BaselineMemberContract[]> _membersByNameAndKind;
    private static readonly IReadOnlyDictionary<string, string> EmptyParameters = new Dictionary<string, string>(StringComparer.Ordinal);

    private BaselineContractCorpus(IReadOnlyDictionary<string, BaselineTypeContract> types)
    {
        _types = types;
        _typesByNormalizedId = Index(
            types.Values.Select(static type => (NormalizeQualifiedIdentifier(type.LogicalId), type)), static type => type.ContractKey);
        _typesByNormalizedName = Index(
            types.Values.SelectMany(static type =>
            {
                var name = NormalizeIdentifier(type.Name);
                var logicalName = NormalizeIdentifier(type.LogicalId.Split('.').Last());
                return name == logicalName
                    ? new[] { (name, type) }
                    : new[] { (name, type), (logicalName, type) };
            }), static type => type.ContractKey);
        _membersByName = Index(
            types.Values.SelectMany(static type => type.Members.Select(member =>
                ((type.ContractKey, NormalizeIdentifier(member.Name)), member))), static member => member.ContractKey);
        _membersByNameAndKind = Index(
            types.Values.SelectMany(static type => type.Members.Select(member =>
                ((type.ContractKey, NormalizeIdentifier(member.Name), NormalizeKind(member.Kind)), member))), static member => member.ContractKey);
    }

    private static Dictionary<TKey, TValue[]> Index<TKey, TValue>(IEnumerable<(TKey Key, TValue Value)> values, Func<TValue, string> orderKey)
        where TKey : notnull
        => values.GroupBy(static value => value.Key)
            .ToDictionary(static group => group.Key, group => group.Select(static value => value.Value)
                .OrderBy(orderKey, StringComparer.Ordinal).ToArray());

    public int TypeCount => _types.Count;
    public int MemberCount => _types.Values.Sum(static x => x.Members.Count);
    public int PartCount => _types.Values.Sum(static x => x.Parts.Count);

    public static BaselineContractCorpus Load(string path)
    {
        if (string.IsNullOrWhiteSpace(path)) throw new ArgumentException("A Wrapper Contract path is required.", nameof(path));
        var fullPath = Path.GetFullPath(path);
        string[] files;
        if (File.Exists(fullPath)) files = new[] { fullPath };
        else if (Directory.Exists(fullPath))
            files = Directory.EnumerateFiles(fullPath, "*.wrapper-contract.json", SearchOption.AllDirectories)
                .OrderBy(static x => x, StringComparer.OrdinalIgnoreCase).ToArray();
        else throw new FileNotFoundException($"Baseline documentation contract is missing: {path}", path);
        if (files.Length == 0) throw new InvalidDataException($"No *.wrapper-contract.json files were found beneath '{path}'.");

        var types = new Dictionary<string, BaselineTypeContract>(StringComparer.Ordinal);
        foreach (var file in files) ReadContract(file, types);
        return new BaselineContractCorpus(types);
    }

    /// <summary>Creates contract-identity targets for corpus validation. Generation should instead pass projected Data logical IDs.</summary>
    public IReadOnlyList<DocumentationTarget> CreateValidationTargets()
    {
        var targets = new List<DocumentationTarget>(TypeCount + PartCount + MemberCount);
        foreach (var type in _types.Values.OrderBy(static x => x.LogicalId, StringComparer.Ordinal))
        {
            targets.Add(DocumentationTarget.ForType(type.LogicalId, "baseline:" + type.LogicalId, type.LogicalId, type.Name, type.Namespace));
            foreach (var part in type.Parts.OrderBy(static x => x.Source, StringComparer.Ordinal))
            {
                var partDigest = Convert.ToHexString(SHA256.HashData(Encoding.UTF8.GetBytes(part.Source))).ToLowerInvariant();
                targets.Add(DocumentationTarget.ForType(
                    $"{type.LogicalId}/part-{partDigest}", $"baseline:{type.LogicalId}:part-{partDigest}",
                    type.LogicalId, type.Name, type.Namespace, part.Source));
            }
            foreach (var member in type.Members.OrderBy(static x => x.ContractKey, StringComparer.Ordinal))
            {
                var digest = Convert.ToHexString(SHA256.HashData(Encoding.UTF8.GetBytes(member.ContractKey))).ToLowerInvariant();
                targets.Add(DocumentationTarget.ForMember(
                    $"{type.LogicalId}/member-{digest}", $"baseline:{type.LogicalId}:member-{digest}", type.LogicalId,
                    member.Name, member.Kind, member.Signature, member.ParameterNames, type.Name, type.Namespace));
            }
        }
        return targets;
    }

    internal BaselineResolution Resolve(DocumentationTarget target)
    {
        if (target.Kind == DocumentationTargetKind.BindingKey)
            return ResolveLegacy(target.BindingKey);

        if (target.ContractTypeLogicalId.EndsWith(":invalid-documentation", StringComparison.Ordinal))
            return BaselineResolution.Unmatched($"Wrapper Contract type '{target.ContractTypeLogicalId}' was not found.");

        var owners = ResolveTypes(target.ContractTypeLogicalId, target.ContractTypeName ?? (target.Kind == DocumentationTargetKind.Type ? target.Name : ""), target.ContractTypeNamespace);
        if (owners.Count == 0)
            return BaselineResolution.Unmatched($"Wrapper Contract type '{target.ContractTypeLogicalId}' was not found.");
        if (owners.Count != 1)
            return BaselineResolution.Ambiguous(owners.Select(static x => x.ContractKey).ToArray(), $"Multiple Wrapper Contract types matched '{target.ContractTypeLogicalId}'.");
        var owner = owners[0];
        if (target.Kind == DocumentationTargetKind.Type)
        {
            var part = owner.Parts.FirstOrDefault(candidate => PathEquals(candidate.Source, target.ContractSourcePath));
            return part is null
                ? BaselineResolution.Exact(owner.ContractKey, owner.Document, Array.Empty<string>())
                : BaselineResolution.Exact(owner.ContractKey + "|part|" + part.Source, part.Document, Array.Empty<string>());
        }

        if (target.ContractSignature?.EndsWith(":invalid-documentation", StringComparison.Ordinal) == true)
            return BaselineResolution.Unmatched($"Wrapper Contract member '{target.ContractSignature}' was not found.");

        var normalizedMemberName = NormalizeIdentifier(target.Name);
        BaselineMemberContract[] candidates;
        if (string.IsNullOrWhiteSpace(target.ContractMemberKind))
            _membersByName.TryGetValue((owner.ContractKey, normalizedMemberName), out candidates!);
        else
            _membersByNameAndKind.TryGetValue((owner.ContractKey, normalizedMemberName, NormalizeKind(target.ContractMemberKind)), out candidates!);
        candidates ??= Array.Empty<BaselineMemberContract>();
        if (candidates.Length == 0)
            return BaselineResolution.Unmatched($"No Wrapper Contract member matched '{target.Name}'.");

        if (candidates.Length > 1)
        {
            if (!string.IsNullOrWhiteSpace(target.ContractSignature))
            {
                var exact = candidates.Where(member => string.Equals(member.Signature, target.ContractSignature, StringComparison.Ordinal)).ToArray();
                if (exact.Length != 0) candidates = exact;
                else
                {
                    var normalizedSignature = NormalizeSignature(target.ContractSignature);
                    var normalized = candidates.Where(member => member.NormalizedSignature == normalizedSignature).ToArray();
                    if (normalized.Length != 0) candidates = normalized;
                    else candidates = MatchSignatureShape(candidates, target);
                }
            }
            else candidates = MatchSignatureShape(candidates, target);
        }
        return ResolveMember(target, candidates);
    }

    private IReadOnlyList<BaselineTypeContract> ResolveTypes(string logicalId, string name, string? typeNamespace)
    {
        var normalizedId = NormalizeQualifiedIdentifier(logicalId);
        if (!_typesByNormalizedId.TryGetValue(normalizedId, out var byId))
            byId = Array.Empty<BaselineTypeContract>();
        var normalizedName = NormalizeIdentifier(name);
        if (normalizedName.Length == 0) return byId;
        var normalizedNamespace = string.IsNullOrWhiteSpace(typeNamespace) ? null : NormalizeQualifiedIdentifier(typeNamespace);
        if (byId.Length == 1)
        {
            var type = byId[0];
            if (NormalizeIdentifier(type.Name) == normalizedName
                && (normalizedNamespace is null || NormalizeQualifiedIdentifier(type.Namespace) == normalizedNamespace))
                return byId;
        }
        else if (byId.Length > 1)
        {
            var matchingIdAndName = byId.Where(type => NormalizeIdentifier(type.Name) == normalizedName
                && (normalizedNamespace is null || NormalizeQualifiedIdentifier(type.Namespace) == normalizedNamespace)).ToArray();
            if (matchingIdAndName.Length != 0) return matchingIdAndName;
        }

        if (!_typesByNormalizedName.TryGetValue(normalizedName, out var byName))
            byName = Array.Empty<BaselineTypeContract>();
        BaselineTypeContract[] named;
        if (normalizedNamespace is null || byName.Length == 0) named = byName;
        else if (byName.Length == 1)
            named = NormalizeQualifiedIdentifier(byName[0].Namespace) == normalizedNamespace ? byName : Array.Empty<BaselineTypeContract>();
        else named = byName.Where(type => NormalizeQualifiedIdentifier(type.Namespace) == normalizedNamespace).ToArray();
        // Split/companion projections may retain their primary canonical ID while
        // emitting a separately documented Wrapper Contract type (for example Foo_).
        // Prefer that unique type name; otherwise preserve an explicit stable ID.
        return named.Length != 0 ? named : byId;
    }

    private static bool PathEquals(string source, string? targetSource)
    {
        if (string.IsNullOrWhiteSpace(targetSource)) return false;
        static string NormalizePath(string value) => value.Replace('\\', '/').TrimStart('/');
        var normalizedSource = NormalizePath(source);
        var normalizedTarget = NormalizePath(targetSource);
        return string.Equals(normalizedSource, normalizedTarget, StringComparison.OrdinalIgnoreCase)
            || normalizedTarget.EndsWith("/" + normalizedSource, StringComparison.OrdinalIgnoreCase);
    }
    private static BaselineMemberContract[] MatchSignatureShape(IReadOnlyList<BaselineMemberContract> candidates, DocumentationTarget target)
    {
        if (target.EmittedParameterNames is { } emittedNames)
        {
            var byCount = candidates.Where(member => member.ParameterNames.Count == emittedNames.Count).ToArray();
            if (byCount.Length != 0) candidates = byCount;
            var normalizedNames = emittedNames.Select(NormalizeIdentifier).ToArray();
            var byNames = candidates.Where(member => member.ParameterNames.Select(NormalizeIdentifier).SequenceEqual(normalizedNames, StringComparer.Ordinal)).ToArray();
            if (byNames.Length != 0) candidates = byNames;
        }
        if (!string.IsNullOrWhiteSpace(target.ContractSignature))
        {
            var targetTypes = ParseSignatureParameterTypes(target.ContractSignature);
            if (targetTypes is not null)
            {
                var byTypes = candidates.Where(member => member.ParameterTypes.SequenceEqual(targetTypes, StringComparer.Ordinal)).ToArray();
                if (byTypes.Length != 0) candidates = byTypes;
            }
        }
        return candidates.ToArray();
    }

    private BaselineResolution ResolveLegacy(string key)
    {
        if (_types.TryGetValue(key, out var type))
            return BaselineResolution.Exact(type.ContractKey, type.Document, Array.Empty<string>());

        var separator = key.LastIndexOf('/');
        if (separator <= 0) separator = key.LastIndexOf('.');
        if (separator <= 0) return BaselineResolution.Unmatched($"Wrapper Contract documentation key '{key}' was not found.");
        var typeKey = key[..separator];
        var memberName = key[(separator + 1)..];
        if (!_types.TryGetValue(typeKey, out type)) return BaselineResolution.Unmatched($"Wrapper Contract documentation key '{key}' was not found.");
        return ResolveMember(DocumentationTarget.Legacy(key), type.Members.Where(x => x.Name == memberName).OrderBy(static x => x.ContractKey, StringComparer.Ordinal).ToArray());
    }

    private static BaselineResolution ResolveMember(DocumentationTarget target, IReadOnlyList<BaselineMemberContract> candidates)
    {
        if (candidates.Count == 0) return BaselineResolution.Unmatched($"No Wrapper Contract member matched '{target.Name}'.");
        if (candidates.Count != 1)
            return BaselineResolution.Ambiguous(candidates.Select(static x => x.ContractKey).ToArray(), $"Multiple Wrapper Contract overloads matched '{target.Name}'.");
        var member = candidates[0];
        var emittedNames = target.EmittedParameterNames ?? member.ParameterNames;
        if (emittedNames.Count != member.ParameterNames.Count)
            throw new InvalidDataException($"Target '{target.LogicalId}' emits {emittedNames.Count} parameters, but Wrapper Contract member '{member.ContractKey}' has {member.ParameterNames.Count}.");
        var normalizedNames = emittedNames.Select(DocumentationTarget.ParameterName).ToArray();
        if (normalizedNames.Any(string.IsNullOrWhiteSpace) || normalizedNames.Distinct(StringComparer.Ordinal).Count() != normalizedNames.Length)
            throw new InvalidDataException($"Target '{target.LogicalId}' has invalid emitted parameter names.");
        var document = member.Document is null ? null : ReconcileParameterNames(member.Document, member.ParameterNames, normalizedNames, member.ContractKey);
        return BaselineResolution.Exact(member.ContractKey, document, normalizedNames);
    }

    private static XmlDocumentation ReconcileParameterNames(XmlDocumentation source, IReadOnlyList<string> sourceNames, IReadOnlyList<string> emittedNames, string contractKey)
    {
        if (sourceNames.SequenceEqual(emittedNames, StringComparer.Ordinal)) return source;
        if (string.Equals(source.ParseStatus, "invalid", StringComparison.OrdinalIgnoreCase))
            return source;
        XElement root;
        try { root = XElement.Parse("<root>" + source.RawXml + "</root>", LoadOptions.PreserveWhitespace); }
        catch (Exception error) when (error is System.Xml.XmlException or ArgumentException)
        { throw new InvalidDataException($"Malformed XML documentation for '{contractKey}'.", error); }
        var map = sourceNames.Select((name, index) => (name, emittedNames[index])).ToDictionary(static x => x.name, static x => x.Item2, StringComparer.Ordinal);
        foreach (var element in root.Descendants().Where(static x => x.Name.LocalName is "param" or "paramref"))
        {
            var attribute = element.Attribute("name");
            if (attribute is null || !map.TryGetValue(DocumentationTarget.ParameterName(attribute.Value), out var replacement))
                throw new InvalidDataException($"XML documentation for '{contractKey}' refers to a parameter absent from its source signature.");
            attribute.Value = replacement;
        }
        var raw = string.Concat(root.Nodes().Select(static x => x.ToString(SaveOptions.DisableFormatting)));
        var result = XmlDocumentation.Parse(raw);
        if (result.Parameters.Keys.Any(x => !emittedNames.Contains(x, StringComparer.Ordinal)))
            throw new InvalidDataException($"Reconciled XML documentation for '{contractKey}' contains a parameter absent from the emitted signature.");
        return result;
    }

    private static void ReadContract(string path, IDictionary<string, BaselineTypeContract> result)
    {
        JsonDocument document;
        try { document = JsonDocument.Parse(File.ReadAllBytes(path)); }
        catch (JsonException error) { throw new InvalidDataException($"Malformed Wrapper Contract JSON: {path}", error); }
        using (document)
        {
            var root = document.RootElement;
            if (Required(root, "SchemaVersion") != "1.0" || Required(root, "ContractKind") != "NetOffice.WrapperContract")
                throw new InvalidDataException($"Baseline documentation requires Wrapper Contract schema 1.0: {path}");
            if (!root.TryGetProperty("Types", out var types) || types.ValueKind != JsonValueKind.Array)
                throw new InvalidDataException($"Wrapper Contract has no Types array: {path}");
            foreach (var typeElement in types.EnumerateArray())
            {
                var logicalId = Required(typeElement, "LogicalId");
                var typeName = Required(typeElement, "Name", logicalId.Split('.').Last());
                var typeNamespace = Required(typeElement, "Namespace", logicalId.Contains('.') ? logicalId[..logicalId.LastIndexOf('.')] : "");
                var typeDocument = ReadDocument(typeElement, Array.Empty<string>(), logicalId);
                var members = new List<BaselineMemberContract>();
                if (!typeElement.TryGetProperty("Members", out var memberElements) || memberElements.ValueKind != JsonValueKind.Array)
                    throw new InvalidDataException($"Wrapper Contract type '{logicalId}' has no Members array.");
                foreach (var memberElement in memberElements.EnumerateArray())
                {
                    var name = Required(memberElement, "Name");
                    var kind = Required(memberElement, "Kind", "member");
                    var signature = Required(memberElement, "Signature", name + (memberElement.TryGetProperty("Parameters", out var p) ? p.GetString() : ""));
                    var parameterText = memberElement.TryGetProperty("Parameters", out var parameters) && parameters.ValueKind == JsonValueKind.String ? parameters.GetString() : null;
                    var parameterNames = ParseParameterNames(parameterText);
                    var parameterTypes = ParseParameterTypes(parameterText);
                    var contractKey = logicalId + "|" + kind + "|" + signature;
                    members.Add(new BaselineMemberContract(contractKey, name, kind, signature, NormalizeSignature(signature), parameterNames, parameterTypes, ReadDocument(memberElement, parameterNames, contractKey)));
                }
                var duplicate = members.GroupBy(static x => x.ContractKey, StringComparer.Ordinal).FirstOrDefault(static x => x.Count() != 1);
                if (duplicate is not null) throw new InvalidDataException($"Ambiguous baseline documentation key '{duplicate.Key}'.");
                var parts = new List<BaselineTypePart>();
                if (typeElement.TryGetProperty("Parts", out var partElements) && partElements.ValueKind == JsonValueKind.Array)
                {
                    foreach (var partElement in partElements.EnumerateArray())
                    {
                        var source = Required(partElement, "Source");
                        parts.Add(new BaselineTypePart(source, ReadDocument(partElement, Array.Empty<string>(), logicalId + "|part|" + source)));
                    }
                }
                if (!result.TryAdd(logicalId, new BaselineTypeContract(logicalId, typeName, typeNamespace, logicalId, typeDocument, parts, members)))
                    throw new InvalidDataException($"Duplicate Wrapper Contract type logical ID '{logicalId}'.");
            }
        }
    }

    private static XmlDocumentation? ReadDocument(JsonElement owner, IReadOnlyList<string> parameters, string contractKey)
    {
        if (!owner.TryGetProperty("Documentation", out var value) || value.ValueKind == JsonValueKind.Null) return null;
        if (value.ValueKind != JsonValueKind.Object) throw new InvalidDataException($"Documentation for '{contractKey}' is not an object.");
        var raw = value.TryGetProperty("Raw", out var rawValue) && rawValue.ValueKind == JsonValueKind.String ? rawValue.GetString() : null;
        if (string.IsNullOrWhiteSpace(raw)) raw = BuildStructuredFragment(value, parameters);
        if (string.IsNullOrWhiteSpace(raw)) return null;

        var status = OptionalString(value, "ParseStatus");
        if (string.Equals(status, "invalid", StringComparison.OrdinalIgnoreCase))
            return ReadInvalidDocument(value, raw, OptionalString(value, "ParseError"));
        if (status is not null && !string.Equals(status, "parsed", StringComparison.OrdinalIgnoreCase))
            throw new InvalidDataException($"Documentation for '{contractKey}' has unknown parse status '{status}'.");
        if (status is not null
            && raw.IndexOf("<exception", StringComparison.OrdinalIgnoreCase) < 0
            && raw.IndexOf("<example", StringComparison.OrdinalIgnoreCase) < 0)
            return new XmlDocumentation(raw, OptionalString(value, "Summary"), OptionalString(value, "Remarks"),
                OptionalString(value, "Returns"), OptionalString(value, "Value"), ReadParameters(value),
                Array.Empty<XmlDocumentationElement>(), status, OptionalString(value, "ParseError"));

        try { return XmlDocumentation.Parse(raw); }
        catch (InvalidDataException error)
        {
            return ReadInvalidDocument(value, raw, error.Message);
        }
    }

    private static XmlDocumentation ReadInvalidDocument(JsonElement value, string raw, string? error)
        => XmlDocumentation.Invalid(raw, error, OptionalString(value, "Summary"), OptionalString(value, "Remarks"),
            OptionalString(value, "Returns"), OptionalString(value, "Value"), ReadParameters(value));

    private static IReadOnlyDictionary<string, string> ReadParameters(JsonElement value)
    {
        if (!value.TryGetProperty("Parameters", out var parameterValues) || parameterValues.ValueKind != JsonValueKind.Object)
            return EmptyParameters;
        var entries = parameterValues.EnumerateObject();
        if (!entries.MoveNext()) return EmptyParameters;
        var parameters = new Dictionary<string, string>(StringComparer.Ordinal);
        do
        {
            var parameter = entries.Current;
            parameters[DocumentationTarget.ParameterName(parameter.Name)] = parameter.Value.ValueKind == JsonValueKind.String ? parameter.Value.GetString() ?? "" : "";
        } while (entries.MoveNext());
        return parameters;
    }

    private static string? OptionalString(JsonElement owner, string name)
        => owner.TryGetProperty(name, out var value) && value.ValueKind == JsonValueKind.String ? value.GetString() : null;

    private static string? BuildStructuredFragment(JsonElement value, IReadOnlyList<string> parameterNames)
    {
        var elements = new List<XElement>();
        AddText("summary", "Summary");
        if (value.TryGetProperty("Parameters", out var parameters) && parameters.ValueKind == JsonValueKind.Object)
        {
            var values = parameters.EnumerateObject().ToDictionary(static x => DocumentationTarget.ParameterName(x.Name), static x => x.Value.GetString() ?? "", StringComparer.Ordinal);
            foreach (var name in parameterNames)
                if (values.TryGetValue(name, out var text)) elements.Add(new XElement("param", new XAttribute("name", name), text));
            foreach (var unknown in values.Keys.Except(parameterNames, StringComparer.Ordinal))
                throw new InvalidDataException($"Structured baseline documentation names unknown parameter '{unknown}'.");
        }
        AddText("returns", "Returns");
        AddText("value", "Value");
        AddText("remarks", "Remarks");
        return elements.Count == 0 ? null : string.Join("\n", elements.Select(static x => x.ToString(SaveOptions.DisableFormatting)));

        void AddText(string elementName, string propertyName)
        {
            if (value.TryGetProperty(propertyName, out var property) && property.ValueKind == JsonValueKind.String && !string.IsNullOrWhiteSpace(property.GetString()))
                elements.Add(new XElement(elementName, property.GetString()));
        }
    }

    internal static IReadOnlyList<string> ParseParameterNames(string? text)
    {
        if (string.IsNullOrWhiteSpace(text)) return Array.Empty<string>();
        var value = text.Trim();
        if (value.Length >= 2 && ((value[0] == '(' && value[^1] == ')') || (value[0] == '[' && value[^1] == ']'))) value = value[1..^1];
        if (string.IsNullOrWhiteSpace(value)) return Array.Empty<string>();
        var segments = SplitParameters(value);
        var names = new List<string>(segments.Count);
        foreach (var segment in segments)
        {
            var declaration = RemoveDefaultValue(segment);
            var matches = Regex.Matches(declaration, @"@?[\p{L}_][\p{L}\p{N}_]*", RegexOptions.CultureInvariant);
            if (matches.Count == 0) throw new InvalidDataException($"Unable to read parameter name from Wrapper Contract signature segment '{segment}'.");
            names.Add(DocumentationTarget.ParameterName(matches[^1].Value));
        }
        return names;
    }

    internal static IReadOnlyList<string> ParseParameterTypes(string? text)
    {
        if (string.IsNullOrWhiteSpace(text)) return Array.Empty<string>();
        var value = text.Trim();
        if (value.Length >= 2 && ((value[0] == '(' && value[^1] == ')') || (value[0] == '[' && value[^1] == ']'))) value = value[1..^1];
        if (string.IsNullOrWhiteSpace(value)) return Array.Empty<string>();
        return SplitParameters(value).Select(ParameterType).ToArray();
    }

    private static IReadOnlyList<string>? ParseSignatureParameterTypes(string signature)
    {
        var open = signature.IndexOf('(');
        var close = signature.LastIndexOf(')');
        if (open >= 0 && close > open) return ParseParameterTypes(signature[open..(close + 1)]);
        open = signature.IndexOf('[');
        close = signature.LastIndexOf(']');
        if (open >= 0 && close > open) return ParseParameterTypes(signature[open..(close + 1)]);
        return Array.Empty<string>();
    }

    private static string ParameterType(string segment)
    {
        var declaration = RemoveDefaultValue(segment).Trim();
        while (declaration.StartsWith("[", StringComparison.Ordinal))
        {
            var close = declaration.IndexOf(']');
            if (close < 0) break;
            declaration = declaration[(close + 1)..].TrimStart();
        }
        var matches = Regex.Matches(declaration, @"@?[\p{L}_][\p{L}\p{N}_]*", RegexOptions.CultureInvariant);
        if (matches.Count == 0) throw new InvalidDataException($"Unable to read parameter type from Wrapper Contract signature segment '{segment}'.");
        var type = declaration[..matches[^1].Index].Trim();
        type = Regex.Replace(type, @"^(?:(?:this|ref|out|in|params)\s+)+", "", RegexOptions.CultureInvariant | RegexOptions.IgnoreCase);
        return NormalizeType(type);
    }

    private static string NormalizeSignature(string signature)
        => Regex.Replace(NormalizeType(signature.Trim().TrimEnd('{', ';').Trim()), @"\s+", "", RegexOptions.CultureInvariant);

    private static string NormalizeType(string value)
    {
        var normalized = Regex.Replace(value.Replace("global::", "", StringComparison.Ordinal), @"(?:[\p{L}_][\p{L}\p{N}_]*\.)*[\p{L}_][\p{L}\p{N}_]*", static match =>
        {
            var token = match.Value.Split('.').Last().ToLowerInvariant();
            return token switch
            {
                "int" or "int32" => "int32",
                "uint" or "uint32" => "uint32",
                "short" or "int16" => "int16",
                "ushort" or "uint16" => "uint16",
                "long" or "int64" => "int64",
                "ulong" or "uint64" => "uint64",
                "byte" or "sbyte" => token,
                "float" or "single" => "single",
                "double" => "double",
                "decimal" => "decimal",
                "bool" or "boolean" => "boolean",
                "char" => "char",
                "string" => "string",
                "object" => "object",
                "void" => "void",
                _ => token
            };
        }, RegexOptions.CultureInvariant);
        return Regex.Replace(normalized, @"\s+", "", RegexOptions.CultureInvariant);
    }

    private static string NormalizeIdentifier(string value) => DocumentationTarget.ParameterName(value).ToLowerInvariant();
    private static string NormalizeKind(string value)
    {
        Span<char> buffer = value.Length <= 64 ? stackalloc char[value.Length] : new char[value.Length];
        var length = 0;
        foreach (var character in value)
        {
            if (character is >= 'A' and <= 'Z') buffer[length++] = (char)(character + ('a' - 'A'));
            else if (character is (>= 'a' and <= 'z') or (>= '0' and <= '9')) buffer[length++] = character;
        }
        return new string(buffer[..length]);
    }
    private static string NormalizeQualifiedIdentifier(string value)
        => value.Replace("global::", "", StringComparison.Ordinal).Trim().ToLowerInvariant();

    private static IReadOnlyList<string> SplitParameters(string value)
    {
        var result = new List<string>();
        var start = 0;
        var angle = 0; var round = 0; var square = 0; var brace = 0; var quote = '\0'; var escaped = false;
        for (var index = 0; index < value.Length; index++)
        {
            var current = value[index];
            if (quote != '\0')
            {
                if (escaped) escaped = false;
                else if (current == '\\') escaped = true;
                else if (current == quote) quote = '\0';
                continue;
            }
            if (current is '\'' or '"') { quote = current; continue; }
            switch (current)
            {
                case '<': angle++; break; case '>': angle--; break;
                case '(': round++; break; case ')': round--; break;
                case '[': square++; break; case ']': square--; break;
                case '{': brace++; break; case '}': brace--; break;
                case ',' when angle == 0 && round == 0 && square == 0 && brace == 0:
                    result.Add(value[start..index].Trim()); start = index + 1; break;
            }
            if (angle < 0 || round < 0 || square < 0 || brace < 0) throw new InvalidDataException($"Unbalanced Wrapper Contract parameter list '{value}'.");
        }
        result.Add(value[start..].Trim());
        if (result.Any(string.IsNullOrWhiteSpace)) throw new InvalidDataException($"Empty parameter in Wrapper Contract parameter list '{value}'.");
        return result;
    }

    private static string RemoveDefaultValue(string value)
    {
        var angle = 0; var round = 0; var square = 0; var brace = 0; var quote = '\0'; var escaped = false;
        for (var index = 0; index < value.Length; index++)
        {
            var current = value[index];
            if (quote != '\0')
            {
                if (escaped) escaped = false;
                else if (current == '\\') escaped = true;
                else if (current == quote) quote = '\0';
                continue;
            }
            if (current is '\'' or '"') { quote = current; continue; }
            switch (current)
            {
                case '<': angle++; break; case '>': angle--; break;
                case '(': round++; break; case ')': round--; break;
                case '[': square++; break; case ']': square--; break;
                case '{': brace++; break; case '}': brace--; break;
                case '=' when angle == 0 && round == 0 && square == 0 && brace == 0: return value[..index].Trim();
            }
        }
        return value.Trim();
    }

    private static string Required(JsonElement owner, string name, string? fallback = null)
    {
        if (owner.TryGetProperty(name, out var value) && value.ValueKind == JsonValueKind.String && !string.IsNullOrWhiteSpace(value.GetString())) return value.GetString()!;
        if (fallback is not null) return fallback;
        throw new InvalidDataException($"Baseline contract field '{name}' is missing.");
    }
}

public static class BaselineDocumentation
{
    public const string RepresentationVersion = "wrapper-contract-docs-v2";
    public static BaselineContractCorpus Load(string path) => BaselineContractCorpus.Load(path);
}


internal sealed record BaselineTypeContract(string ContractKey, string Name, string Namespace, string LogicalId, XmlDocumentation? Document, IReadOnlyList<BaselineTypePart> Parts, IReadOnlyList<BaselineMemberContract> Members);
internal sealed record BaselineTypePart(string Source, XmlDocumentation? Document);
internal sealed record BaselineMemberContract(string ContractKey, string Name, string Kind, string Signature, string NormalizedSignature, IReadOnlyList<string> ParameterNames, IReadOnlyList<string> ParameterTypes, XmlDocumentation? Document);
internal sealed record BaselineResolution(MappingKind Kind, string? ContractKey, XmlDocumentation? Document, IReadOnlyList<string> ParameterNames, string? Reason, IReadOnlyList<string> Candidates)
{
    public static BaselineResolution Exact(string key, XmlDocumentation? document, IReadOnlyList<string> names) => new(MappingKind.Exact, key, document, names, null, Array.Empty<string>());
    public static BaselineResolution Unmatched(string reason) => new(MappingKind.Unmatched, null, null, Array.Empty<string>(), reason, Array.Empty<string>());
    public static BaselineResolution Ambiguous(IReadOnlyList<string> candidates, string reason) => new(MappingKind.Ambiguous, null, null, Array.Empty<string>(), reason, candidates);
}
