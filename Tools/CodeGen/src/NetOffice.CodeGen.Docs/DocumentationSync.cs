using System.Collections.ObjectModel;
using System.Security.Cryptography;
using System.Text;
using System.Text.Json;
using System.Xml.Linq;

namespace NetOffice.CodeGen.Docs;

public static class DocumentationSync
{
    public const string ResultSchemaVersion = "netoffice-documentation-result-v1";

    /// <summary>Compatibility entry point. New consumers should pass typed <see cref="DocumentationTarget"/> values.</summary>
    public static DocumentationResult Sync(IEnumerable<string> logicalIds, DocsOptions options)
        => Sync(logicalIds.Select(DocumentationTarget.Legacy), options);

    /// <summary>Binds offline documentation to projected declarations. Any ambiguous mapping fails the operation.</summary>
    public static DocumentationResult Sync(IEnumerable<DocumentationTarget> targets, DocsOptions options)
    {
        ArgumentNullException.ThrowIfNull(targets);
        ArgumentNullException.ThrowIfNull(options);
        options.Validate();
        var orderedTargets = OrderAndValidateTargets(targets);

        var result = options.Profile switch
        {
            DocsProfile.Baseline => BindBaseline(orderedTargets, options),
            DocsProfile.Vba => BindVba(orderedTargets, options),
            _ => throw new ArgumentException("Unknown documentation profile.", nameof(options))
        };
        if (!string.IsNullOrWhiteSpace(options.ReportDirectory)) WriteReports(options.ReportDirectory, result);
        if (result.Mappings.Any(static x => x.Kind == MappingKind.Ambiguous)) throw new DocumentationMappingException(result);
        return result;
    }

    internal static DocumentationTarget[] OrderAndValidateTargets(IEnumerable<DocumentationTarget> targets)
    {
        var orderedTargets = targets.OrderBy(static x => x.LogicalId, StringComparer.Ordinal).ThenBy(static x => x.BindingKey, StringComparer.Ordinal).ToArray();
        var logicalIds = new HashSet<string>(StringComparer.Ordinal);
        var bindingKeys = new HashSet<string>(StringComparer.Ordinal);
        foreach (var target in orderedTargets)
        {
            target.Validate();
            if (!logicalIds.Add(target.LogicalId)) throw new InvalidDataException($"Duplicate documentation logical ID '{target.LogicalId}'.");
            if (!bindingKeys.Add(target.BindingKey)) throw new InvalidDataException($"Duplicate documentation binding key '{target.BindingKey}'.");
        }
        return orderedTargets;
    }

    public static void WriteReports(string directory, DocumentationResult result)
    {
        ArgumentException.ThrowIfNullOrWhiteSpace(directory);
        ArgumentNullException.ThrowIfNull(result);
        Directory.CreateDirectory(directory);
        var records = result.Mappings.OrderBy(static x => x.LogicalId, StringComparer.Ordinal).ThenBy(static x => x.BindingKey, StringComparer.Ordinal).ToArray();
        WriteReport(Path.Combine(directory, "mapping-ledger.json"), result, records);
        WriteReport(Path.Combine(directory, "unmatched-report.json"), result, records.Where(static x => x.Kind == MappingKind.Unmatched).ToArray());
        WriteReport(Path.Combine(directory, "ambiguous-report.json"), result, records.Where(static x => x.Kind == MappingKind.Ambiguous).ToArray());
    }


    internal static string Hash(string text) => Convert.ToHexString(SHA256.HashData(Encoding.UTF8.GetBytes(text))).ToLowerInvariant();

    private static DocumentationResult BindBaseline(IReadOnlyList<DocumentationTarget> targets, DocsOptions options)
        => BindBaselineBatch(targets, BaselineContractCorpus.Load(options.BaselineContractPath!));

    internal static DocumentationResult BindBaselineBatch(IReadOnlyList<DocumentationTarget> targets, BaselineContractCorpus corpus)
    {
        var mappings = new List<MappingRecord>(targets.Count);
        var documents = new Dictionary<string, BoundDocumentation>(targets.Count, StringComparer.Ordinal);
        foreach (var target in targets)
        {
            var resolution = corpus.Resolve(target);
            XmlDocumentation? document = null;
            MappingKind kind;
            string? reason = resolution.Reason;
            if (resolution.Kind == MappingKind.Exact)
            {
                if (resolution.Document is not null)
                {
                    kind = MappingKind.Exact;
                    document = resolution.Document;
                    ValidateParameterNames(target, document, resolution.ParameterNames);
                }
                else
                {
                    kind = MappingKind.Unmatched;
                    reason = "Current-source Wrapper Contract documentation is absent.";
                }
            }
            else
            {
                kind = resolution.Kind == MappingKind.Ambiguous ? MappingKind.Ambiguous : MappingKind.Unmatched;
                document = Fallback(target, DocsProfile.Baseline);
            }
            mappings.Add(new MappingRecord(target.LogicalId, target.BindingKey, kind,
                kind == MappingKind.Exact ? resolution.ContractKey : null, null, resolution.ContractKey, reason,
                resolution.Candidates.Count == 0 ? null : resolution.Candidates));
            if (document is not null)
                documents.Add(target.BindingKey, new BoundDocumentation(target.LogicalId, target.BindingKey, kind, resolution.ContractKey, document));
        }
        return CreateResult(mappings, documents, null, DocsProfile.Baseline);
    }

    private static DocumentationResult BindVba(IReadOnlyList<DocumentationTarget> targets, DocsOptions options)
    {
        // This profile deliberately remains an offline, pin-validated reader. Network acquisition belongs to docs sync orchestration.
        if (!Directory.Exists(options.SourceDirectory))
            throw new DirectoryNotFoundException($"Documentation source is missing (generation performs no network access): {options.SourceDirectory}");
        var pin = DocsPin.Load(options.PinPath ?? Path.Combine(options.SourceDirectory, "pin.json"));
        var articles = LoadArticles(options.SourceDirectory);
        var aliases = LoadMap(options.AliasMappingsPath);
        var manuals = LoadMap(options.ManualMappingsPath);
        var mappings = new List<MappingRecord>(targets.Count);
        var documents = new Dictionary<string, BoundDocumentation>(StringComparer.Ordinal);
        foreach (var target in targets)
        {
            DocumentationArticle? article = null;
            MappingKind kind;
            string? reason = null;
            IReadOnlyList<string>? candidates = null;
            var configuredKey = manuals.ContainsKey(target.BindingKey) ? target.BindingKey : target.LogicalId;
            if (manuals.TryGetValue(configuredKey, out var manualId))
            {
                article = FindArticle(articles, manualId) ?? throw new InvalidDataException($"Manual mapping '{configuredKey}' points to missing article '{manualId}'.");
                kind = MappingKind.Manual;
            }
            else
            {
                configuredKey = aliases.ContainsKey(target.BindingKey) ? target.BindingKey : target.LogicalId;
                if (aliases.TryGetValue(configuredKey, out var aliasId))
                {
                    article = FindArticle(articles, aliasId) ?? throw new InvalidDataException($"Alias mapping '{configuredKey}' points to missing article '{aliasId}'.");
                    kind = MappingKind.ApprovedAlias;
                }
                else
                {
                    var keys = new[] { target.BindingKey, target.LogicalId, target.Name }.Where(static x => !string.IsNullOrWhiteSpace(x)).Select(Normalize).Distinct(StringComparer.Ordinal).ToArray();
                    var matches = articles.Where(articleCandidate => ArticleKeys(articleCandidate).Any(keys.Contains)).OrderBy(static x => x.Id, StringComparer.Ordinal).ToArray();
                    if (matches.Length == 1) { article = matches[0]; kind = MappingKind.Exact; }
                    else if (matches.Length == 0) { kind = MappingKind.Unmatched; reason = "No pinned article matched."; }
                    else { kind = MappingKind.Ambiguous; reason = "Multiple pinned articles matched."; candidates = matches.Select(static x => x.Id).ToArray(); }
                }
            }

            XmlDocumentation document;
            if (article is null) document = Fallback(target, DocsProfile.Vba);
            else
            {
                document = XmlDocumentation.Parse(MarkdownXmlConverter.Convert(article.Markdown, article.Title));
                // Unlike baseline reconciliation, VBA parameter names are never positionally guessed.
                ValidateParameterNames(target, document, target.EmittedParameterNames);
            }
            mappings.Add(new MappingRecord(target.LogicalId, target.BindingKey, kind, article?.Id, article?.CanonicalUrl, null, reason, candidates));
            documents.Add(target.BindingKey, new BoundDocumentation(target.LogicalId, target.BindingKey, kind, null, document));
        }
        return CreateResult(mappings, documents, pin, DocsProfile.Vba);
    }

    internal static DocumentationResult CreateResult(IReadOnlyList<MappingRecord> mappings, Dictionary<string, BoundDocumentation> documents, DocsPin? pin, DocsProfile profile, bool canonicalize = false)
    {
        IReadOnlyList<MappingRecord> finalMappings = canonicalize
            ? mappings.OrderBy(static x => x.LogicalId, StringComparer.Ordinal).ThenBy(static x => x.BindingKey, StringComparer.Ordinal).ToArray()
            : mappings;
        Dictionary<string, BoundDocumentation> finalDocuments;
        if (canonicalize)
            finalDocuments = documents.OrderBy(static x => x.Key, StringComparer.Ordinal)
                .ToDictionary(static x => x.Key, static x => x.Value, StringComparer.Ordinal);
        else finalDocuments = documents;
        var readOnlyDocuments = new ReadOnlyDictionary<string, BoundDocumentation>(finalDocuments);
        return new DocumentationResult(finalMappings, readOnlyDocuments, Digest(finalMappings, readOnlyDocuments, pin, profile), pin, profile);
    }

    internal static DocumentationResult CreateSummaryResult(IReadOnlyList<MappingRecord> mappings, IReadOnlyDictionary<string, string> documentHashes, DocsProfile profile)
    {
        var finalMappings = mappings.OrderBy(static x => x.LogicalId, StringComparer.Ordinal)
            .ThenBy(static x => x.BindingKey, StringComparer.Ordinal).ToArray();
        var documents = new ReadOnlyDictionary<string, BoundDocumentation>(
            new Dictionary<string, BoundDocumentation>(StringComparer.Ordinal));
        return new DocumentationResult(finalMappings, documents, DigestHashes(finalMappings, documentHashes, null, profile), null, profile);
    }

    private static void ValidateParameterNames(DocumentationTarget target, XmlDocumentation document, IReadOnlyList<string>? effectiveNames)
    {
        if (target.Kind == DocumentationTargetKind.Type || effectiveNames is null || string.Equals(document.ParseStatus, "invalid", StringComparison.OrdinalIgnoreCase)) return;
        var emitted = effectiveNames.Select(DocumentationTarget.ParameterName).ToHashSet(StringComparer.Ordinal);
        var invalid = document.Parameters.Keys.Where(x => !emitted.Contains(x)).Order(StringComparer.Ordinal).ToArray();
        if (invalid.Length != 0)
            throw new InvalidDataException($"Documentation for '{target.LogicalId}' contains parameters absent from the emitted signature: {string.Join(", ", invalid)}.");
    }

    private static XmlDocumentation Fallback(DocumentationTarget target, DocsProfile profile)
    {
        var label = string.IsNullOrWhiteSpace(target.Name) ? target.LogicalId : target.Name;
        var source = profile == DocsProfile.Baseline ? "baseline Wrapper Contract" : "approved pinned VBA documentation";
        return XmlDocumentation.Parse(new XElement("summary", $"Documentation is not available in the {source} for {label}.").ToString(SaveOptions.DisableFormatting));
    }

    private static string Digest(IEnumerable<MappingRecord> records, IReadOnlyDictionary<string, BoundDocumentation> documents, DocsPin? pin, DocsProfile profile)
    {
        var canonical = AppendDigestRecords(records, pin, profile);
        foreach (var document in documents.OrderBy(static x => x.Key, StringComparer.Ordinal))
            canonical.Append(document.Key).Append('\t').Append(Hash(document.Value.Documentation.RawXml)).Append('\n');
        return Hash(canonical.ToString());
    }

    private static string DigestHashes(IEnumerable<MappingRecord> records, IReadOnlyDictionary<string, string> documentHashes, DocsPin? pin, DocsProfile profile)
    {
        var canonical = AppendDigestRecords(records, pin, profile);
        foreach (var document in documentHashes.OrderBy(static x => x.Key, StringComparer.Ordinal))
            canonical.Append(document.Key).Append('\t').Append(document.Value).Append('\n');
        return Hash(canonical.ToString());
    }

    private static StringBuilder AppendDigestRecords(IEnumerable<MappingRecord> records, DocsPin? pin, DocsProfile profile)
    {
        var canonical = new StringBuilder(ResultSchemaVersion).Append('\n').Append(profile).Append('\n')
            .Append(pin?.Repository).Append('\t').Append(pin?.Commit).Append('\t').Append(pin?.License).Append('\n');
        foreach (var record in records)
            canonical.Append(record.LogicalId).Append('\t').Append(record.BindingKey).Append('\t').Append(record.Kind).Append('\t')
                .Append(record.ArticleId).Append('\t').Append(record.CanonicalUrl).Append('\t').Append(record.ContractKey).Append('\t')
                .Append(record.Reason).Append('\t').AppendJoin(',', record.Candidates ?? Array.Empty<string>()).Append('\n');
        return canonical;
    }

    private static void WriteReport(string path, DocumentationResult result, IReadOnlyList<MappingRecord> records)
    {
        var report = new DocumentationReport(ResultSchemaVersion, result.Profile, result.Digest, result.Pin, records.Count, records);
        File.WriteAllText(path, JsonSerializer.Serialize(report, JsonOptions.Indented) + "\n", new UTF8Encoding(false));
    }

    private static IReadOnlyList<DocumentationArticle> LoadArticles(string root)
    {
        var result = new List<DocumentationArticle>();
        foreach (var file in Directory.EnumerateFiles(root, "*.md", SearchOption.AllDirectories).OrderBy(static x => x, StringComparer.OrdinalIgnoreCase))
        {
            var relativePath = Path.GetRelativePath(root, file).Replace('\\', '/');
            if (relativePath.Equals("README.md", StringComparison.OrdinalIgnoreCase)) continue;
            var markdown = File.ReadAllText(file, Encoding.UTF8);
            var title = markdown.Replace("\r\n", "\n", StringComparison.Ordinal).Split('\n').Select(static x => x.Trim()).FirstOrDefault(static x => x.StartsWith("# ", StringComparison.Ordinal))?[2..].Trim();
            if (string.IsNullOrWhiteSpace(title)) throw new InvalidDataException($"Malformed VBA documentation article (missing H1): {relativePath}");
            var id = Path.GetFileNameWithoutExtension(relativePath);
            result.Add(new DocumentationArticle(id, title, relativePath, markdown, $"https://learn.microsoft.com/en-us/office/vba/{relativePath[..^3].ToLowerInvariant()}"));
        }
        return result;
    }

    private static DocumentationArticle? FindArticle(IEnumerable<DocumentationArticle> articles, string id)
        => articles.FirstOrDefault(x => x.Id.Equals(id, StringComparison.Ordinal));
    private static IEnumerable<string> ArticleKeys(DocumentationArticle article)
        => new[] { Normalize(article.Id), Normalize(Path.GetFileNameWithoutExtension(article.RelativePath)), Normalize(article.Title) };
    private static string Normalize(string value) => new(value.Where(char.IsLetterOrDigit).Select(char.ToLowerInvariant).ToArray());

    private static Dictionary<string, string> LoadMap(string? path)
    {
        if (string.IsNullOrWhiteSpace(path)) return new Dictionary<string, string>(StringComparer.Ordinal);
        if (!File.Exists(path)) throw new FileNotFoundException($"Documentation mapping file is missing: {path}", path);
        return JsonSerializer.Deserialize<Dictionary<string, string>>(File.ReadAllText(path), JsonOptions.Default)
            ?? new Dictionary<string, string>(StringComparer.Ordinal);
    }

    private sealed record DocumentationReport(string SchemaVersion, DocsProfile Profile, string Digest, DocsPin? Pin, int Count, IReadOnlyList<MappingRecord> Records);
}
