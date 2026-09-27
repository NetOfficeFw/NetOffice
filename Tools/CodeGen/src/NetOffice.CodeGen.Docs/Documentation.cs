using System.Security.Cryptography;
using System.Text;
using System.Text.Json;
using System.Xml.Linq;

namespace NetOffice.CodeGen.Docs;

public enum MappingKind { Exact, ApprovedAlias, Manual, Unmatched, Ambiguous }
public enum DocsProfile { Baseline, Vba }

public sealed record DocsPin(string Repository, string Commit, string? License = "CC BY 4.0")
{
    public static DocsPin Load(string path)
    {
        if (!File.Exists(path)) throw new FileNotFoundException($"Documentation pin is missing: {path}", path);
        var pin = JsonSerializer.Deserialize<DocsPin>(File.ReadAllText(path), JsonOptions.Default)
            ?? throw new InvalidDataException("Documentation pin is empty.");
        if (string.IsNullOrWhiteSpace(pin.Repository) || !pin.Repository.Equals("MicrosoftDocs/VBA-Docs", StringComparison.OrdinalIgnoreCase) || string.IsNullOrWhiteSpace(pin.Commit) || !IsSha(pin.Commit))
            throw new InvalidDataException("Documentation pin must target MicrosoftDocs/VBA-Docs and contain a 40-character commit SHA.");
        return pin;
    }
    private static bool IsSha(string value) => value.Length == 40 && value.All(Uri.IsHexDigit);
}

public sealed record DocumentationArticle(string Id, string Title, string RelativePath, string Markdown, string CanonicalUrl);
public sealed record MappingRecord(string LogicalId, MappingKind Kind, string? ArticleId, string? CanonicalUrl, string? Reason = null);
public sealed record DocumentationResult(IReadOnlyList<MappingRecord> Mappings, IReadOnlyDictionary<string,string> XmlDocs, string Digest, DocsPin? Pin);
public sealed record DocsOptions(string SourceDirectory, string? PinPath = null, DocsProfile Profile = DocsProfile.Baseline, bool Locked = false, string? ManualMappingsPath = null, string? AliasMappingsPath = null)
{
    public void Validate()
    {
        if (!Enum.IsDefined(Profile)) throw new ArgumentException("Unknown documentation profile.", nameof(Profile));
        if (Locked && Profile == DocsProfile.Vba && string.IsNullOrWhiteSpace(PinPath)) throw new InvalidOperationException("Locked VBA documentation requires a pin.");
    }
}

public static class DocumentationSync
{
    public static DocumentationResult Sync(IEnumerable<string> logicalIds, DocsOptions options)
    {
        options.Validate();
        var ids = logicalIds.Order(StringComparer.Ordinal).ToArray();
        if (options.Profile == DocsProfile.Baseline)
        {
            var mappings = ids.Select(id => new MappingRecord(id, MappingKind.Unmatched, null, null, "baseline profile does not attach VBA documentation")).ToArray();
            return new DocumentationResult(mappings, ids.ToDictionary(x => x, _ => FallbackXml("Documentation is not available in the baseline profile.")), Digest(mappings, Array.Empty<DocumentationArticle>()), null);
        }
        if (!Directory.Exists(options.SourceDirectory))
            throw new DirectoryNotFoundException($"Documentation source is missing (locked mode performs no network access): {options.SourceDirectory}");
        var pin = DocsPin.Load(options.PinPath ?? Path.Combine(options.SourceDirectory, "pin.json"));
        var articles = LoadArticles(options.SourceDirectory, pin);
        var aliases = LoadMap(options.AliasMappingsPath);
        var manuals = LoadMap(options.ManualMappingsPath);
        var records = new List<MappingRecord>();
        var docs = new Dictionary<string,string>(StringComparer.Ordinal);
        foreach (var id in ids)
        {
            MappingRecord record;
            DocumentationArticle? article = null;
            if (manuals.TryGetValue(id, out var manualId))
            { article = FindArticle(articles, manualId) ?? throw new InvalidDataException($"Manual mapping '{id}' points to missing article '{manualId}'."); record = Map(id, MappingKind.Manual, article); }
            else if (aliases.TryGetValue(id, out var aliasId))
            { article = FindArticle(articles, aliasId) ?? throw new InvalidDataException($"Alias mapping '{id}' points to missing article '{aliasId}'."); record = Map(id, MappingKind.ApprovedAlias, article); }
            else
            {
                var matches = articles.Where(a => Keys(a).Contains(Normalize(id), StringComparer.Ordinal)).ToArray();
                record = matches.Length switch { 1 => Map(id, MappingKind.Exact, matches[0]), 0 => new(id, MappingKind.Unmatched, null, null, "no article matched"), _ => new(id, MappingKind.Ambiguous, null, null, "multiple articles matched") };
                article = matches.Length == 1 ? matches[0] : null;
            }
            records.Add(record);
            docs[id] = article is null ? FallbackXml("No approved documentation mapping is available.") : MarkdownXmlConverter.Convert(article.Markdown, article.Title);
        }
        return new DocumentationResult(records, docs, Digest(records, articles), pin);
    }

    public static void WriteLedger(string path, IEnumerable<MappingRecord> records)
    {
        Directory.CreateDirectory(Path.GetDirectoryName(Path.GetFullPath(path))!);
        File.WriteAllText(path, JsonSerializer.Serialize(records.OrderBy(x => x.LogicalId), JsonOptions.Indented) + "\n", new UTF8Encoding(false));
    }

    private static DocumentationArticle? FindArticle(IEnumerable<DocumentationArticle> articles, string id) => articles.FirstOrDefault(x => x.Id.Equals(id, StringComparison.Ordinal));
    private static MappingRecord Map(string id, MappingKind kind, DocumentationArticle a) => new(id, kind, a.Id, a.CanonicalUrl);
    private static IEnumerable<string> Keys(DocumentationArticle a) => new[] { Normalize(a.Id), Normalize(Path.GetFileNameWithoutExtension(a.RelativePath)), Normalize(a.Title) };
    private static string Normalize(string s) => new(s.Where(char.IsLetterOrDigit).Select(char.ToLowerInvariant).ToArray());
    private static Dictionary<string,string> LoadMap(string? path) => string.IsNullOrWhiteSpace(path) ? new(StringComparer.Ordinal) : JsonSerializer.Deserialize<Dictionary<string,string>>(File.ReadAllText(path), JsonOptions.Default) ?? new(StringComparer.Ordinal);
    private static string FallbackXml(string text) => new XElement("summary", text).ToString(SaveOptions.DisableFormatting);

    private static IReadOnlyList<DocumentationArticle> LoadArticles(string root, DocsPin pin)
    {
        var result = new List<DocumentationArticle>();
        foreach (var file in Directory.EnumerateFiles(root, "*.md", SearchOption.AllDirectories).Order(StringComparer.OrdinalIgnoreCase))
        {
            var rel = Path.GetRelativePath(root, file).Replace('\\', '/');
            if (rel.Equals("README.md", StringComparison.OrdinalIgnoreCase)) continue;
            var markdown = File.ReadAllText(file, Encoding.UTF8);
            var title = markdown.Split('\n').Select(x => x.Trim()).FirstOrDefault(x => x.StartsWith("# ", StringComparison.Ordinal))?[2..].Trim();
            if (string.IsNullOrWhiteSpace(title)) throw new InvalidDataException($"Malformed VBA documentation article (missing H1): {rel}");
            var id = Path.GetFileNameWithoutExtension(rel);
            var url = $"https://learn.microsoft.com/en-us/office/vba/{rel[..^3].ToLowerInvariant()}";
            result.Add(new DocumentationArticle(id, title, rel, markdown, url));
        }
        return result;
    }
    private static string Digest(IEnumerable<MappingRecord> records, IEnumerable<DocumentationArticle> articles)
    {
        var canonical = string.Join("\n", records.OrderBy(x => x.LogicalId).Select(x => $"{x.LogicalId}\t{x.Kind}\t{x.ArticleId}\t{x.CanonicalUrl}")) + "\n" + string.Join("\n", articles.OrderBy(x => x.Id).Select(x => x.Id + "\t" + SHA256.HashData(Encoding.UTF8.GetBytes(x.Markdown)).ToHex()));
        return Convert.ToHexString(SHA256.HashData(Encoding.UTF8.GetBytes(canonical))).ToLowerInvariant();
    }
}

public static class MarkdownXmlConverter
{
    public static string Convert(string markdown, string title)
    {
        if (markdown is null) throw new ArgumentNullException(nameof(markdown));
        var lines = markdown.Replace("\r\n", "\n").Split('\n');
        var summary = new List<string>(); var remarks = new List<string>(); var parameters = new List<(string,string)>(); var target = summary;
        foreach (var raw in lines.Skip(1))
        {
            var line = raw.Trim(); if (line.Length == 0) { if (target == summary) target = remarks; continue; }
            if (line.StartsWith("## ", StringComparison.Ordinal)) { target = remarks; continue; }
            if (line.StartsWith("- `", StringComparison.Ordinal) && line.Contains("`:") ) { var end = line.IndexOf("`:", StringComparison.Ordinal); parameters.Add((line[3..end], Clean(line[(end + 2)..]))); continue; }
            if (!line.StartsWith("#", StringComparison.Ordinal)) target.Add(Clean(line));
        }
        var root = new XElement("member", new XAttribute("name", ""), new XElement("summary", string.Join(" ", summary).Trim()));
        if (parameters.Count > 0) root.Add(parameters.Select(p => new XElement("param", new XAttribute("name", p.Item1), p.Item2)));
        if (remarks.Count > 0) root.Add(new XElement("remarks", string.Join(" ", remarks).Trim()));
        return root.ToString(SaveOptions.DisableFormatting);
    }
    private static string Clean(string value) => value.Replace("**", "", StringComparison.Ordinal).Replace("`", "", StringComparison.Ordinal).Trim();
}

internal static class JsonOptions { public static readonly JsonSerializerOptions Default = new() { PropertyNameCaseInsensitive = true }; public static readonly JsonSerializerOptions Indented = new() { WriteIndented = true, PropertyNamingPolicy = JsonNamingPolicy.CamelCase }; }
internal static class HashExtensions { public static string ToHex(this byte[] bytes) => Convert.ToHexString(bytes).ToLowerInvariant(); }
