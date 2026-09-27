// Copyright (c) 2026 NetOffice contributors
// SPDX-License-Identifier: MIT

using System.Security.Cryptography;
using System.Text;
using System.Text.Json;
using System.Text.Json.Serialization;
using System.Text.RegularExpressions;

namespace NetOffice.CodeGen.ContractExtractor;

internal static class Program
{
    private const string SchemaVersion = "1.0";
    private const string ToolVersion = "1.0.0";

    public static int Main(string[] args)
    {
        try
        {
            var options = Options.Parse(args);
            if (options.ShowHelp)
            {
                Console.WriteLine(Options.Usage);
                return 0;
            }

            if (string.IsNullOrWhiteSpace(options.Source) || string.IsNullOrWhiteSpace(options.Api))
                throw new UsageException("--source and --api are required.");

            var source = Path.GetFullPath(options.Source);
            if (!Directory.Exists(source))
                throw new UsageException("Source directory does not exist: " + source);

            var apiRoot = Path.Combine(source, options.Api);
            if (!Directory.Exists(apiRoot))
                throw new UsageException("API source directory does not exist: " + apiRoot);

            var contract = Extractor.Extract(source, options.Api, apiRoot);
            var output = options.Output ?? Path.Combine("contracts", "wrapper", options.Api + ".wrapper-contract.json");
            var outputPath = Path.GetFullPath(output);
            Directory.CreateDirectory(Path.GetDirectoryName(outputPath)!);
            JsonFile.Write(outputPath, contract);

            var outputDirectory = Path.GetDirectoryName(outputPath)!;
            var stem = Path.GetFileNameWithoutExtension(outputPath);
            if (stem.EndsWith(".wrapper-contract", StringComparison.Ordinal))
                stem = stem[..^".wrapper-contract".Length];

            var ledgerPath = options.Ledger ?? Path.Combine(outputDirectory, stem + ".compatibility-ledger.json");
            var classificationPath = options.Classification ?? Path.Combine(outputDirectory, stem + ".classification.json");
            JsonFile.Write(Path.GetFullPath(ledgerPath), Records.CreateLedger(contract));
            JsonFile.Write(Path.GetFullPath(classificationPath), Records.CreateClassification(contract));

            Console.WriteLine("Extracted " + contract.Types.Count.ToString(System.Globalization.CultureInfo.InvariantCulture) +
                              " types and " + contract.Types.Sum(t => t.Members.Count).ToString(System.Globalization.CultureInfo.InvariantCulture) +
                              " members from " + options.Api + ".");
            Console.WriteLine("Contract: " + outputPath);
            Console.WriteLine("Ledger: " + Path.GetFullPath(ledgerPath));
            Console.WriteLine("Classification: " + Path.GetFullPath(classificationPath));
            return 0;
        }
        catch (UsageException ex)
        {
            Console.Error.WriteLine("error: " + ex.Message);
            Console.Error.WriteLine(Options.Usage);
            return 2;
        }
        catch (Exception ex)
        {
            Console.Error.WriteLine("error: " + ex.Message);
            return 1;
        }
    }
}

internal sealed class UsageException : Exception
{
    public UsageException(string message) : base(message) { }
}

internal sealed class Options
{
    public string Source { get; private set; }
    public string Api { get; private set; }
    public string Output { get; private set; }
    public string Ledger { get; private set; }
    public string Classification { get; private set; }
    public bool ShowHelp { get; private set; }

    public static readonly string Usage = "Usage: dotnet run --project NetOffice.CodeGen.ContractExtractor.csproj -- --source <Source> --api <Api> [--output <file>] [--ledger <file>] [--classification <file>]";

    public static Options Parse(string[] args)
    {
        var result = new Options();
        for (var i = 0; i < args.Length; i++)
        {
            var arg = args[i];
            if (arg is "-h" or "--help")
            {
                result.ShowHelp = true;
                continue;
            }

            if (!arg.StartsWith("--", StringComparison.Ordinal) || i + 1 >= args.Length)
                throw new UsageException("Unknown or incomplete argument: " + arg);

            var value = args[++i];
            switch (arg)
            {
                case "--source": result.Source = value; break;
                case "--api": result.Api = value; break;
                case "--output": result.Output = value; break;
                case "--ledger": result.Ledger = value; break;
                case "--classification": result.Classification = value; break;
                default: throw new UsageException("Unknown argument: " + arg);
            }
        }
        return result;
    }
}

internal static class Extractor
{
    private static readonly Regex NamespaceRegex = new Regex(@"^\s*namespace\s+([A-Za-z_][\w.]*)", RegexOptions.Compiled);
    private static readonly Regex TypeRegex = new Regex(@"^\s*(?<access>public|protected|internal|private)?\s*(?<mods>(?:(?:abstract|sealed|static|partial|unsafe|readonly)\s+)*)?(?<kind>class|interface|struct|enum|delegate)\s+(?<name>[A-Za-z_]\w*(?:\s*<[^>{}]+>)?)\s*(?::\s*(?<bases>[^\{]+))?", RegexOptions.Compiled);
    private static readonly Regex AccessRegex = new Regex(@"^\s*(?<access>public|protected(?:\s+internal)?|internal)\b(?<rest>.*)$", RegexOptions.Compiled);
    private static readonly Regex NameBeforeParenRegex = new Regex(@"(?<name>[A-Za-z_]\w*)\s*(?:<[^>]+>)?\s*\(", RegexOptions.Compiled);
    private static readonly Regex IdentifierRegex = new Regex(@"[A-Za-z_]\w*", RegexOptions.Compiled);

    public static WrapperContract Extract(string sourceRoot, string api, string apiRoot)
    {
        var sourceLabel = StableRootLabel(sourceRoot);
        var contract = new WrapperContract
        {
            SchemaVersion = "1.0",
            ContractKind = "NetOffice.WrapperContract",
            Generator = new GeneratorInfo { Name = "NetOffice.CodeGen.ContractExtractor", Version = "1.0.0" },
            Source = new SourceInfo
            {
                Root = sourceLabel,
                Api = api,
                SourceRoots = new List<string> { NormalizePath(sourceLabel + "/" + api) }
            }
        };

        var paths = Directory.EnumerateFiles(apiRoot, "*.cs", SearchOption.AllDirectories)
            .Where(p => !IsBuildPath(p))
            .OrderBy(p => NormalizePath(Path.GetRelativePath(apiRoot, p)), StringComparer.Ordinal)
            .ToArray();
        if (paths.Length == 0)
        {
            contract.Unknowns.Add(new UnknownRecord { Code = "NO_SOURCE_FILES", Message = "No C# files were found below the API source root.", Path = contract.Source.SourceRoots[0] });
            return contract;
        }

        foreach (var path in paths)
        {
            var relative = NormalizePath(Path.GetRelativePath(apiRoot, path));
            var text = File.ReadAllText(path);
            var file = new SourceFile
            {
                Path = relative,
                Sha256 = Hash(text),
                LineCount = text.Replace("\r\n", "\n", StringComparison.Ordinal).Split('\n').Length
            };
            contract.Source.Files.Add(file);
            ParseFile(text, relative, contract);
        }

        contract.Source.Files = contract.Source.Files.OrderBy(f => f.Path, StringComparer.Ordinal).ToList();
        contract.Types = contract.Types.OrderBy(t => t.Namespace, StringComparer.Ordinal)
            .ThenBy(t => t.Name, StringComparer.Ordinal).ToList();
        foreach (var type in contract.Types)
            type.Members = type.Members.OrderBy(m => m.Line).ThenBy(m => m.Name, StringComparer.Ordinal).ToList();
        foreach (var duplicate in contract.Types.GroupBy(t => t.LogicalId, StringComparer.Ordinal).Where(g => g.Count() > 1).OrderBy(g => g.Key, StringComparer.Ordinal))
        {
            contract.Ambiguities.Add(new AmbiguityRecord
            {
                Code = "DUPLICATE_LOGICAL_ID",
                Message = "More than one source type has the same logical ID.",
                Candidates = duplicate.Select(t => t.Source + ":" + t.Line.ToString(System.Globalization.CultureInfo.InvariantCulture)).OrderBy(v => v, StringComparer.Ordinal).ToList()
            });
        }
        return contract;
    }

    private static void ParseFile(string text, string relative, WrapperContract contract)
    {
        var lines = text.Replace("\r\n", "\n", StringComparison.Ordinal).Replace('\r', '\n').Split('\n');
        var ns = "";
        var pendingAttributes = new List<string>();
        TypeRecord type = null;
        var bodyDepth = -1;
        var depth = 0;
        MemberBuilder member = null;

        for (var index = 0; index < lines.Length; index++)
        {
            var raw = lines[index];
            var line = StripLineComment(raw);
            var namespaceMatch = NamespaceRegex.Match(line);
            if (namespaceMatch.Success && type == null)
                ns = namespaceMatch.Groups[1].Value;

            if (type == null)
            {
                if (line.TrimStart().StartsWith("[", StringComparison.Ordinal))
                {
                    pendingAttributes.Add(line.Trim());
                }
                else
                {
                    var typeMatch = TypeRegex.Match(line);
                    if (typeMatch.Success)
                    {
                        type = new TypeRecord
                        {
                            LogicalId = (ns.Length == 0 ? "" : ns + ".") + typeMatch.Groups["name"].Value,
                            Namespace = ns,
                            Name = typeMatch.Groups["name"].Value,
                            Kind = typeMatch.Groups["kind"].Value,
                            Accessibility = typeMatch.Groups["access"].Success ? typeMatch.Groups["access"].Value : "private",
                            Modifiers = SplitWords(typeMatch.Groups["mods"].Value).ToList(),
                            Attributes = ExtractAttributes(string.Join(" ", pendingAttributes)),
                            Source = relative,
                            Line = index + 1,
                            Signature = Normalize(line)
                        };
                        ParseBases(type, typeMatch.Groups["bases"].Value);
                        contract.Types.Add(type);
                        pendingAttributes.Clear();
                        var opens = Count(line, '{');
                        if (opens > 0)
                            bodyDepth = depth + opens;
                        // A declaration may put its opening brace on the following line.
                    }
                }
            }
            else if (bodyDepth >= 0)
            {
                if (member == null && depth == bodyDepth)
                {
                    if (line.TrimStart().StartsWith("[", StringComparison.Ordinal))
                    {
                        pendingAttributes.Add(line.Trim());
                    }
                    else
                    {
                        var declaration = type.Kind == "enum" ? ParseEnumMemberDeclaration(line) : ParseMemberDeclaration(line, type);
                        if (declaration != null)
                        {
                            declaration.Line = index + 1;
                            declaration.Attributes = ExtractAttributes(string.Join(" ", pendingAttributes));
                            declaration.Source = relative;
                            member = declaration;
                            pendingAttributes.Clear();
                        }
                        else if (IsPotentialPublicDeclaration(line))
                        {
                            contract.Unknowns.Add(new UnknownRecord { Code = "UNPARSED_DECLARATION", Message = Normalize(line), Path = relative, Line = index + 1, LogicalId = type.LogicalId });
                            pendingAttributes.Clear();
                        }
                        else if (!string.IsNullOrWhiteSpace(line) && !line.TrimStart().StartsWith("///", StringComparison.Ordinal) && !line.TrimStart().StartsWith("//", StringComparison.Ordinal))
                        {
                            pendingAttributes.Clear();
                        }
                    }
                }

                if (member != null)
                {
                    member.Lines.Add(line);
                    if (line.Contains("Factory.", StringComparison.Ordinal) || line.Contains("Invoke", StringComparison.Ordinal) || line.Contains("Execute", StringComparison.Ordinal))
                    {
                        var invocation = Normalize(line);
                        if (invocation.Length > 0)
                            member.InvocationText.Add(invocation);
                    }
                }
            }

            var before = depth;
            depth += Count(line, '{') - Count(line, '}');
            if (type != null && bodyDepth < 0 && line.Contains('{', StringComparison.Ordinal))
                bodyDepth = depth;
            if (member != null)
            {
                var hasBody = member.Lines.Any(l => l.Contains('{', StringComparison.Ordinal));
                var complete = (!hasBody && line.Contains(';', StringComparison.Ordinal)) || (hasBody && depth == bodyDepth);
                if (complete)
                {
                    var finished = member.ToRecord();
                    type.Members.Add(finished);
                    member = null;
                }
            }
            if (type != null && bodyDepth >= 0 && depth < bodyDepth)
            {
                if (member != null)
                    type.Members.Add(member.ToRecord());
                member = null;
                type = null;
                bodyDepth = -1;
                pendingAttributes.Clear();
            }
            _ = before;
        }

        if (type != null)
            contract.Unknowns.Add(new UnknownRecord { Code = "UNBALANCED_TYPE", Message = "Type body did not close before end of file.", Path = relative, Line = type.Line, LogicalId = type.LogicalId });
    }

    private static MemberBuilder ParseEnumMemberDeclaration(string line)
    {
        var match = Regex.Match(line, @"^\s*(?<name>[A-Za-z_]\w*)\s*(?:=\s*(?<value>[^,]+))?\s*,?\s*$");
        if (!match.Success)
            return null;
        var value = match.Groups["value"].Success ? Normalize(match.Groups["value"].Value) : null;
        return new MemberBuilder
        {
            Name = match.Groups["name"].Value,
            Kind = "enumValue",
            Accessibility = "public",
            Signature = value == null ? match.Groups["name"].Value : match.Groups["name"].Value + " = " + value
        };
    }

    private static MemberBuilder ParseMemberDeclaration(string line, TypeRecord type)
    {
        var match = AccessRegex.Match(line);
        if (!match.Success)
            return null;
        var rest = match.Groups["rest"].Value.Trim();
        if (rest.Length == 0 || rest.StartsWith("if ", StringComparison.Ordinal) || rest.StartsWith("if(", StringComparison.Ordinal))
            return null;
        var modifiers = new List<string>();
        while (true)
        {
            var m = Regex.Match(rest, @"^(static|virtual|override|abstract|sealed|new|extern|unsafe|async|readonly|const|event)\s+(.*)$");
            if (!m.Success) break;
            modifiers.Add(m.Groups[1].Value);
            rest = m.Groups[2].Value.Trim();
        }

        var builder = new MemberBuilder { Accessibility = match.Groups["access"].Value, Modifiers = modifiers };
        var paren = NameBeforeParenRegex.Match(rest);
        if (paren.Success)
        {
            builder.Name = paren.Groups["name"].Value;
            var prefix = rest[..paren.Index].Trim();
            builder.Kind = string.Equals(builder.Name, type.Name, StringComparison.Ordinal) ? "constructor" : "method";
            builder.ReturnType = builder.Kind == "constructor" ? null : prefix;
            var open = rest.IndexOf('(', paren.Index);
            var close = FindClosingParen(rest, open);
            if (close >= 0)
                builder.Parameters = Normalize(rest.Substring(open, close - open + 1));
            builder.Signature = Normalize(match.Groups["access"].Value + " " + string.Join(" ", modifiers) + " " + rest);
            return builder;
        }

        var declaration = rest;
        var equals = declaration.IndexOf('=');
        if (equals >= 0) declaration = declaration[..equals].Trim();
        declaration = declaration.TrimEnd(';').Trim();
        var identifiers = IdentifierRegex.Matches(declaration).Cast<Match>().Select(m => m.Value).ToList();
        if (identifiers.Count < 1)
            return null;
        builder.Name = identifiers[^1];
        builder.ReturnType = declaration[..Math.Max(0, declaration.LastIndexOf(builder.Name, StringComparison.Ordinal))].Trim();
        if (declaration.Contains("this[", StringComparison.Ordinal))
        {
            builder.Name = "this";
            builder.Kind = "indexer";
        }
        else if (modifiers.Contains("event", StringComparer.Ordinal))
            builder.Kind = "event";
        else if (line.Contains('{', StringComparison.Ordinal) || line.Contains(" get", StringComparison.Ordinal) || line.Contains(" set", StringComparison.Ordinal))
            builder.Kind = "property";
        else
            builder.Kind = "field";
        builder.Signature = Normalize(match.Groups["access"].Value + " " + string.Join(" ", modifiers) + " " + rest);
        return builder;
    }

    private static void ParseBases(TypeRecord type, string bases)
    {
        if (string.IsNullOrWhiteSpace(bases)) return;
        var values = SplitCommaSeparated(bases);
        if (type.Kind == "class" && values.Count > 0)
        {
            type.BaseType = values[0];
            type.Interfaces.AddRange(values.Skip(1));
        }
        else
            type.Interfaces.AddRange(values);
    }

    private static List<string> ExtractAttributes(string text)
    {
        var values = new List<string>();
        foreach (Match match in Regex.Matches(text, @"\[(?<value>[^\]]+)\]"))
        {
            var value = Normalize(match.Groups["value"].Value);
            if (value.Length > 0) values.Add(value);
        }
        return values;
    }

    private static bool IsPotentialPublicDeclaration(string line) => AccessRegex.IsMatch(line) && !line.TrimStart().StartsWith("public class", StringComparison.Ordinal);
    private static int FindClosingParen(string text, int open)
    {
        var level = 0;
        for (var i = open; i < text.Length; i++)
        {
            if (text[i] == '(') level++;
            else if (text[i] == ')' && --level == 0) return i;
        }
        return -1;
    }

    private static string StripLineComment(string line)
    {
        var index = line.IndexOf("//", StringComparison.Ordinal);
        return index >= 0 && !line.TrimStart().StartsWith("///", StringComparison.Ordinal) ? line[..index] : line;
    }

    private static List<string> SplitCommaSeparated(string value)
    {
        var result = new List<string>();
        var level = 0;
        var start = 0;
        for (var i = 0; i < value.Length; i++)
        {
            if (value[i] == '<') level++;
            else if (value[i] == '>') level--;
            else if (value[i] == ',' && level == 0)
            {
                result.Add(Normalize(value[start..i]));
                start = i + 1;
            }
        }
        var last = Normalize(value[start..]);
        if (last.Length > 0) result.Add(last);
        return result;
    }

    private static string[] SplitWords(string value) => value.Split(new[] { ' ', '\t', '\r', '\n' }, StringSplitOptions.RemoveEmptyEntries);
    private static int Count(string value, char character) => value.Count(c => c == character);
    private static string Normalize(string value) => Regex.Replace(value.Trim(), @"\s+", " ");
    private static string NormalizePath(string value) => value.Replace('\\', '/');
    private static string StableRootLabel(string path)
    {
        var trimmed = path.TrimEnd(Path.DirectorySeparatorChar, Path.AltDirectorySeparatorChar);
        var label = Path.GetFileName(trimmed);
        return string.IsNullOrEmpty(label) ? NormalizePath(trimmed) : NormalizePath(label);
    }
    private static bool IsBuildPath(string path) => path.Split(Path.DirectorySeparatorChar, Path.AltDirectorySeparatorChar).Any(p => p is "bin" or "obj");
    private static string Hash(string value)
    {
        var bytes = SHA256.HashData(Encoding.UTF8.GetBytes(value.Replace("\r\n", "\n", StringComparison.Ordinal).Replace('\r', '\n')));
        return Convert.ToHexString(bytes).ToLowerInvariant();
    }
}

internal sealed class MemberBuilder
{
    public string Accessibility { get; set; }
    public List<string> Modifiers { get; set; } = new List<string>();
    public string Name { get; set; }
    public string Kind { get; set; }
    public string ReturnType { get; set; }
    public string Parameters { get; set; }
    public string Signature { get; set; }
    public string Source { get; set; }
    public int Line { get; set; }
    public List<string> Attributes { get; set; } = new List<string>();
    public List<string> InvocationText { get; set; } = new List<string>();
    public List<string> Lines { get; } = new List<string>();

    public MemberRecord ToRecord() => new MemberRecord
    {
        Name = Name,
        Kind = Kind,
        Accessibility = Accessibility,
        Modifiers = Modifiers,
        ReturnType = ReturnType,
        Parameters = Parameters,
        Signature = Signature,
        Source = Source,
        Line = Line,
        Attributes = Attributes,
        InvocationText = InvocationText.Distinct(StringComparer.Ordinal).ToList()
    };
}

internal sealed class TypeRecord
{
    public string LogicalId { get; set; }
    public string Namespace { get; set; }
    public string Name { get; set; }
    public string Kind { get; set; }
    public string Accessibility { get; set; }
    public List<string> Modifiers { get; set; } = new List<string>();
    public string BaseType { get; set; }
    public List<string> Interfaces { get; set; } = new List<string>();
    public List<string> Attributes { get; set; } = new List<string>();
    public string Source { get; set; }
    public int Line { get; set; }
    public string Signature { get; set; }
    public List<MemberRecord> Members { get; set; } = new List<MemberRecord>();
}

internal sealed class MemberRecord
{
    public string Name { get; set; }
    public string Kind { get; set; }
    public string Accessibility { get; set; }
    public List<string> Modifiers { get; set; } = new List<string>();
    public string ReturnType { get; set; }
    public string Parameters { get; set; }
    public string Signature { get; set; }
    public List<string> Attributes { get; set; } = new List<string>();
    public List<string> InvocationText { get; set; } = new List<string>();
    public string Source { get; set; }
    public int Line { get; set; }
}

internal sealed class WrapperContract
{
    public string SchemaVersion { get; set; }
    public string ContractKind { get; set; }
    public GeneratorInfo Generator { get; set; }
    public SourceInfo Source { get; set; }
    public List<TypeRecord> Types { get; set; } = new List<TypeRecord>();
    public List<UnknownRecord> Unknowns { get; set; } = new List<UnknownRecord>();
    public List<AmbiguityRecord> Ambiguities { get; set; } = new List<AmbiguityRecord>();
}

internal sealed class GeneratorInfo { public string Name { get; set; } public string Version { get; set; } }
internal sealed class SourceInfo
{
    public string Root { get; set; }
    public string Api { get; set; }
    public List<string> SourceRoots { get; set; } = new List<string>();
    public List<SourceFile> Files { get; set; } = new List<SourceFile>();
}
internal sealed class SourceFile { public string Path { get; set; } public string Sha256 { get; set; } public int LineCount { get; set; } }
internal sealed class UnknownRecord { public string Code { get; set; } public string Message { get; set; } public string Path { get; set; } public int? Line { get; set; } public string LogicalId { get; set; } }
internal sealed class AmbiguityRecord { public string Code { get; set; } public string Message { get; set; } public List<string> Candidates { get; set; } = new List<string>(); }

internal static class Records
{
    public static CompatibilityLedger CreateLedger(WrapperContract contract)
    {
        var ledger = new CompatibilityLedger { SchemaVersion = "1.0", ContractKind = "NetOffice.WrapperCompatibilityLedger", Api = contract.Source.Api };
        foreach (var type in contract.Types)
        {
            ledger.Entries.Add(new LedgerEntry { LogicalId = type.LogicalId, RecordKind = "type", Source = type.Source, Status = "extracted", ExpectedMatchCount = 1 });
            foreach (var member in type.Members)
                ledger.Entries.Add(new LedgerEntry { LogicalId = type.LogicalId + "." + member.Name, RecordKind = "member", Source = member.Source, Status = "extracted", ExpectedMatchCount = 1 });
        }
        ledger.Entries = ledger.Entries.OrderBy(e => e.LogicalId, StringComparer.Ordinal).ThenBy(e => e.RecordKind, StringComparer.Ordinal).ToList();
        return ledger;
    }

    public static ClassificationReport CreateClassification(WrapperContract contract)
    {
        var report = new ClassificationReport { SchemaVersion = "1.0", ContractKind = "NetOffice.WrapperClassification", Api = contract.Source.Api };
        foreach (var type in contract.Types)
        {
            var folder = type.Source.Contains("/", StringComparison.Ordinal) ? type.Source[..type.Source.IndexOf('/', StringComparison.Ordinal)] : "Root";
            report.Entries.Add(new ClassificationEntry { LogicalId = type.LogicalId, Source = type.Source, Category = folder, Kind = type.Kind });
        }
        report.Entries = report.Entries.OrderBy(e => e.LogicalId, StringComparer.Ordinal).ToList();
        return report;
    }
}

internal sealed class CompatibilityLedger
{
    public string SchemaVersion { get; set; }
    public string ContractKind { get; set; }
    public string Api { get; set; }
    public List<LedgerEntry> Entries { get; set; } = new List<LedgerEntry>();
}
internal sealed class LedgerEntry
{
    public string LogicalId { get; set; }
    public string RecordKind { get; set; }
    public string Source { get; set; }
    public string Status { get; set; }
    public int ExpectedMatchCount { get; set; }
}
internal sealed class ClassificationReport
{
    public string SchemaVersion { get; set; }
    public string ContractKind { get; set; }
    public string Api { get; set; }
    public List<ClassificationEntry> Entries { get; set; } = new List<ClassificationEntry>();
}
internal sealed class ClassificationEntry
{
    public string LogicalId { get; set; }
    public string Source { get; set; }
    public string Category { get; set; }
    public string Kind { get; set; }
}

internal static class JsonFile
{
    private static readonly JsonSerializerOptions Options = new JsonSerializerOptions
    {
        WriteIndented = true,
        Encoder = System.Text.Encodings.Web.JavaScriptEncoder.UnsafeRelaxedJsonEscaping,
        DefaultIgnoreCondition = JsonIgnoreCondition.Never
    };

    public static void Write<T>(string path, T value)
    {
        var json = JsonSerializer.Serialize(value, Options).Replace("\r\n", "\n", StringComparison.Ordinal).Replace('\r', '\n') + "\n";
        File.WriteAllText(path, json, new UTF8Encoding(false));
    }
}
