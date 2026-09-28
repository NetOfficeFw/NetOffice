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
            return CommandRunner.Run(args);
        }
        catch (CommandLineException ex)
        {
            Console.Error.WriteLine("error: " + ex.Message);
            Console.Error.WriteLine(CommandRunner.Usage);
            return 2;
        }
        catch (Exception ex)
        {
            Console.Error.WriteLine("error: " + ex.Message);
            return 1;
        }
    }
}

internal static class Extractor
{
    private static readonly Regex NamespaceRegex = new Regex(@"^\s*namespace\s+([A-Za-z_][\w.]*)", RegexOptions.Compiled);
    private static readonly Regex TypeRegex = new Regex(@"^\s*(?<access>public|protected|internal|private)?\s*(?<mods>(?:(?:abstract|sealed|static|partial|unsafe|readonly)\s+)*)?(?<kind>class|interface|struct|enum)\s+(?<name>[A-Za-z_]\w*(?:\s*<[^>{}]+>)?)\s*(?::\s*(?<bases>[^\{]+))?", RegexOptions.Compiled);
    private static readonly Regex DelegateRegex = new Regex(@"^\s*(?<access>public|protected|internal|private)?\s*(?<mods>(?:(?:static|unsafe)\s+)*)?delegate\s+(?<return>.+?)\s+(?<name>[A-Za-z_]\w*(?:\s*<[^>{}]+>)?)\s*(?<parameters>\([^;]*\))\s*;", RegexOptions.Compiled);
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
            .Where(p => !IsBuildPath(Path.GetRelativePath(apiRoot, p)))
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
        contract.Types = MergePartialTypes(contract.Types).OrderBy(t => t.LogicalId, StringComparer.Ordinal)
            .ThenBy(t => t.Source, StringComparer.Ordinal).ToList();
        foreach (var type in contract.Types)
        {
            type.Interfaces = type.Interfaces.Distinct(StringComparer.Ordinal).OrderBy(value => value, StringComparer.Ordinal).ToList();
            type.Attributes = type.Attributes.OrderBy(value => value, StringComparer.Ordinal).ToList();
            type.SupportVersions = ContractIdentity.GetSupportVersions(type.Attributes);
            foreach (var member in type.Members)
            {
                member.LogicalId = ContractIdentity.MemberLogicalId(type.LogicalId, member);
                member.Attributes = member.Attributes.OrderBy(value => value, StringComparer.Ordinal).ToList();
                member.InvocationText = member.InvocationText.Distinct(StringComparer.Ordinal).OrderBy(value => value, StringComparer.Ordinal).ToList();
                member.SupportVersions = ContractIdentity.GetSupportVersions(member.Attributes);
            }
            type.Members = type.Members.OrderBy(m => m.LogicalId, StringComparer.Ordinal)
                .ThenBy(m => m.Source, StringComparer.Ordinal).ThenBy(m => m.Line).ToList();
            foreach (var duplicate in type.Members.GroupBy(m => m.LogicalId, StringComparer.Ordinal).Where(g => g.Count() > 1).OrderBy(g => g.Key, StringComparer.Ordinal))
            {
                contract.Ambiguities.Add(new AmbiguityRecord
                {
                    Code = "DUPLICATE_MEMBER_LOGICAL_ID",
                    Message = "More than one source member has the same semantic logical ID.",
                    Candidates = duplicate.Select(m => m.LogicalId + ":" + m.Source + ":" + m.Line.ToString(System.Globalization.CultureInfo.InvariantCulture)).OrderBy(v => v, StringComparer.Ordinal).ToList()
                });
            }
        }
        foreach (var duplicate in contract.Types.GroupBy(t => t.LogicalId, StringComparer.Ordinal)
            .Where(group => group.Count() > 1 && !group.All(type => type.Modifiers.Contains("partial", StringComparer.Ordinal)))
            .OrderBy(group => group.Key, StringComparer.Ordinal))
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

    private static List<TypeRecord> MergePartialTypes(IEnumerable<TypeRecord> records)
    {
        var result = new List<TypeRecord>();
        foreach (var group in records.GroupBy(type => type.LogicalId, StringComparer.Ordinal)
                     .OrderBy(group => group.Key, StringComparer.Ordinal))
        {
            var declarations = group.OrderBy(type => type.Source, StringComparer.Ordinal).ThenBy(type => type.Line).ToList();
            foreach (var declaration in declarations)
            {
                declaration.Parts.Add(new TypePartRecord
                {
                    Source = declaration.Source,
                    Line = declaration.Line,
                    Signature = declaration.Signature,
                    Attributes = declaration.Attributes.ToList(),
                    BaseType = declaration.BaseType,
                    Interfaces = declaration.Interfaces.ToList(),
                    Documentation = declaration.Documentation
                });
            }

            if (declarations.Count == 1 || !declarations.All(type => type.Modifiers.Contains("partial", StringComparer.Ordinal)))
            {
                result.AddRange(declarations);
                continue;
            }

            var merged = declarations[0];
            merged.Modifiers = declarations.SelectMany(type => type.Modifiers).Distinct(StringComparer.Ordinal)
                .OrderBy(value => value, StringComparer.Ordinal).ToList();
            merged.Attributes = declarations.SelectMany(type => type.Attributes).Distinct(StringComparer.Ordinal)
                .OrderBy(value => value, StringComparer.Ordinal).ToList();
            merged.Interfaces = declarations.SelectMany(type => type.Interfaces).Distinct(StringComparer.Ordinal)
                .OrderBy(value => value, StringComparer.Ordinal).ToList();
            merged.BaseType = declarations.Select(type => type.BaseType).FirstOrDefault(value => !string.IsNullOrWhiteSpace(value));
            merged.Members = declarations.SelectMany(type => type.Members).ToList();
            merged.Parts = declarations.SelectMany(type => type.Parts).OrderBy(part => part.Source, StringComparer.Ordinal)
                .ThenBy(part => part.Line).ToList();
            result.Add(merged);
        }
        return result;
    }

    private static void ParseFile(string text, string relative, WrapperContract contract)
    {
        var lines = text.Replace("\r\n", "\n", StringComparison.Ordinal).Replace('\r', '\n').Split('\n');
        var ns = "";
        var pendingAttributes = new List<string>();
        var pendingDocumentation = new List<string>();
        TypeRecord type = null;
        var bodyDepth = -1;
        var depth = 0;
        var member = (MemberBuilder)null;
        var typeHeader = new List<string>();

        for (var index = 0; index < lines.Length; index++)
        {
            var raw = lines[index];
            var line = StripLineComment(raw);
            var trimmed = line.Trim();
            var namespaceMatch = NamespaceRegex.Match(line);
            if (namespaceMatch.Success && type == null)
                ns = namespaceMatch.Groups[1].Value;

            if (type == null)
            {
                if (trimmed.StartsWith("///", StringComparison.Ordinal))
                {
                    pendingDocumentation.Add(trimmed[3..].TrimStart());
                }
                else if (trimmed.StartsWith("[", StringComparison.Ordinal) || HasOpenAttribute(pendingAttributes))
                {
                    pendingAttributes.Add(trimmed);
                }
                else
                {
                    var delegateMatch = DelegateRegex.Match(line);
                    if (delegateMatch.Success)
                    {
                        contract.Types.Add(CreateDelegate(delegateMatch, ns, relative, index + 1, pendingAttributes, pendingDocumentation, contract));
                        pendingAttributes.Clear();
                        pendingDocumentation.Clear();
                    }
                    else
                    {
                        var typeMatch = TypeRegex.Match(line);
                        if (typeMatch.Success)
                        {
                            typeHeader.Clear();
                            typeHeader.Add(line);
                            type = CreateType(typeMatch, ns, relative, index + 1, pendingAttributes, pendingDocumentation, contract);
                            contract.Types.Add(type);
                            pendingAttributes.Clear();
                            pendingDocumentation.Clear();
                            if (line.Contains('{', StringComparison.Ordinal))
                            {
                                bodyDepth = depth + Count(line, '{') - Count(line, '}');
                            }
                        }
                    }
                }
            }
            else if (bodyDepth < 0)
            {
                // Complete declarations whose opening brace is on a later line, including
                // multiline base/interface lists.
                if (!string.IsNullOrWhiteSpace(line))
                {
                    typeHeader.Add(line);
                    var header = Normalize(string.Join(" ", typeHeader));
                    var typeMatch = TypeRegex.Match(header);
                    if (typeMatch.Success)
                    {
                        ParseBases(type, typeMatch.Groups["bases"].Value);
                        type.Signature = Normalize(header[..Math.Max(0, header.IndexOf('{', StringComparison.Ordinal) < 0 ? header.Length : header.IndexOf('{', StringComparison.Ordinal))]);
                    }
                }
                if (line.Contains('{', StringComparison.Ordinal))
                    bodyDepth = depth + Count(line, '{') - Count(line, '}');
            }
            else
            {
                if (member == null && depth == bodyDepth)
                {
                    if (trimmed.StartsWith("///", StringComparison.Ordinal))
                    {
                        pendingDocumentation.Add(trimmed[3..].TrimStart());
                    }
                    else if (trimmed.StartsWith("[", StringComparison.Ordinal) || HasOpenAttribute(pendingAttributes))
                    {
                        pendingAttributes.Add(trimmed);
                    }
                    else if (type.Kind == "enum" && !string.IsNullOrWhiteSpace(trimmed) &&
                             !trimmed.StartsWith("#", StringComparison.Ordinal) && !trimmed.StartsWith("}", StringComparison.Ordinal))
                    {
                        member = new MemberBuilder { Source = relative, Line = index + 1 };
                        member.Lines.Add(line);
                        member.Attributes = ExtractAttributes(string.Join(" ", pendingAttributes));
                        member.Documentation = ParseDocumentation(pendingDocumentation, relative, index + 1, type.LogicalId, contract);
                        pendingAttributes.Clear();
                        pendingDocumentation.Clear();
                    }
                    else if (IsDeclarationStart(line, type))
                    {
                        member = new MemberBuilder { Source = relative, Line = index + 1 };
                        member.Lines.Add(line);
                        member.Attributes = ExtractAttributes(string.Join(" ", pendingAttributes));
                        member.Documentation = ParseDocumentation(pendingDocumentation, relative, index + 1, type.LogicalId, contract);
                        pendingAttributes.Clear();
                        pendingDocumentation.Clear();
                    }
                    else if (IsPotentialDeclaration(line))
                    {
                        AddUnknown(contract, "UNPARSED_DECLARATION", Normalize(line), relative, index + 1, type.LogicalId);
                        pendingAttributes.Clear();
                        pendingDocumentation.Clear();
                    }
                    else if (!string.IsNullOrWhiteSpace(trimmed))
                    {
                        // Documentation and attributes belong only to the immediately
                        // following declaration. Consume trivia for ignored private/manual
                        // declarations instead of leaking it to the next contract member.
                        pendingAttributes.Clear();
                        pendingDocumentation.Clear();
                    }
                }

                if (member != null)
                {
                    if ((member.Lines.Count == 0 || !ReferenceEquals(member.Lines[^1], line)) &&
                        !(type.Kind == "enum" && trimmed.StartsWith("}", StringComparison.Ordinal)))
                        member.Lines.Add(line);
                    if (line.Contains("Factory.", StringComparison.Ordinal) || line.Contains("Invoke", StringComparison.Ordinal) || line.Contains("Execute", StringComparison.Ordinal))
                    {
                        var invocation = Normalize(line);
                        if (invocation.Length > 0)
                            member.InvocationText.Add(invocation);
                    }
                }
            }

            depth += Count(line, '{') - Count(line, '}');
            if (member != null && ((type.Kind == "enum" && Normalize(string.Join(" ", member.Lines)).EndsWith(",", StringComparison.Ordinal)) || IsCompleteMember(member, depth, bodyDepth)))
            {
                var finished = type.Kind == "enum" ? ParseEnumMemberDeclaration(string.Join(" ", member.Lines)) : ParseMemberDeclaration(string.Join(" ", member.Lines), type);
                if (finished == null)
                {
                    AddUnknown(contract, "UNPARSED_DECLARATION", Normalize(string.Join(" ", member.Lines)), relative, member.Line, type.LogicalId);
                }
                else
                {
                    finished.Line = member.Line;
                    finished.Source = relative;
                    finished.Attributes = member.Attributes;
                    finished.InvocationText = CompleteInvocationText(finished, member);
                    finished.Documentation = member.Documentation;
                    type.Members.Add(finished.ToRecord());
                }
                member = null;
            }

            if (type != null && bodyDepth >= 0 && depth < bodyDepth)
            {
                if (member != null)
                {
                    var finished = type.Kind == "enum" ? ParseEnumMemberDeclaration(string.Join(" ", member.Lines)) : ParseMemberDeclaration(string.Join(" ", member.Lines), type);
                    if (finished != null)
                    {
                        finished.Line = member.Line;
                        finished.Source = relative;
                        finished.Attributes = member.Attributes;
                        finished.InvocationText = CompleteInvocationText(finished, member);
                        finished.Documentation = member.Documentation;
                        type.Members.Add(finished.ToRecord());
                    }
                    else
                        AddUnknown(contract, "UNPARSED_DECLARATION", Normalize(string.Join(" ", member.Lines)), relative, member.Line, type.LogicalId);
                }
                member = null;
                type = null;
                bodyDepth = -1;
                pendingAttributes.Clear();
                pendingDocumentation.Clear();
                typeHeader.Clear();
            }
        }


        if (type != null)
        {
            if (member != null)
                AddUnknown(contract, "UNPARSED_DECLARATION", Normalize(string.Join(" ", member.Lines)), relative, member.Line, type.LogicalId);
            AddUnknown(contract, "UNBALANCED_TYPE", "Type body did not close before end of file.", relative, type.Line, type.LogicalId);
        }
    }
    private static List<string> CompleteInvocationText(MemberBuilder finished, MemberBuilder parsed)
    {
        if (string.Equals(finished.Kind, "enumValue", StringComparison.Ordinal))
            return new List<string>();
        if (!string.Equals(finished.Kind, "event", StringComparison.Ordinal))
            return parsed.InvocationText;
        var invocationText = new List<string>();
        var declaration = string.Join(" ", parsed.Lines);
        var bodyStart = FindTopLevel(declaration, '{');
        if (bodyStart >= 0)
            invocationText.Add(Normalize(declaration[bodyStart..]));
        return invocationText;
    }

    private static TypeRecord CreateType(Match typeMatch, string ns, string relative, int line, List<string> attributes, List<string> documentation, WrapperContract contract)
    {
        var type = new TypeRecord
        {
            LogicalId = (ns.Length == 0 ? "" : ns + ".") + typeMatch.Groups["name"].Value,
            Namespace = ns,
            Name = typeMatch.Groups["name"].Value,
            Kind = typeMatch.Groups["kind"].Value,
            Accessibility = typeMatch.Groups["access"].Success ? typeMatch.Groups["access"].Value : "private",
            Modifiers = SplitWords(typeMatch.Groups["mods"].Value).ToList(),
            Attributes = ExtractAttributes(string.Join(" ", attributes)),
            Documentation = ParseDocumentation(documentation, relative, line, null, contract),
            Source = relative,
            Line = line,
            Signature = Normalize(typeMatch.Value)
        };
        ParseBases(type, typeMatch.Groups["bases"].Value);
        return type;
    }

    private static TypeRecord CreateDelegate(Match match, string ns, string relative, int line, List<string> attributes, List<string> documentation, WrapperContract contract)
    {
        var name = match.Groups["name"].Value;
        return new TypeRecord
        {
            LogicalId = (ns.Length == 0 ? "" : ns + ".") + name,
            Namespace = ns,
            Name = name,
            Kind = "delegate",
            Accessibility = match.Groups["access"].Success ? match.Groups["access"].Value : "private",
            Modifiers = SplitWords(match.Groups["mods"].Value).ToList(),
            Attributes = ExtractAttributes(string.Join(" ", attributes)),
            Documentation = ParseDocumentation(documentation, relative, line, null, contract),
            Source = relative,
            Line = line,
            Signature = Normalize(match.Value)
        };
    }

    private static bool IsCompleteMember(MemberBuilder member, int depth, int bodyDepth)
    {
        var declaration = Normalize(string.Join(" ", member.Lines));
        var hasBody = declaration.Contains('{', StringComparison.Ordinal);
        if (hasBody)
            return depth == bodyDepth && (declaration.Contains('}', StringComparison.Ordinal) || declaration.Contains("=>", StringComparison.Ordinal));
        return declaration.Contains(';', StringComparison.Ordinal);
    }

    private static bool IsDeclarationStart(string line, TypeRecord type)
    {
        var trimmed = line.Trim();
        if (string.IsNullOrWhiteSpace(trimmed) || trimmed.StartsWith("#", StringComparison.Ordinal) ||
            trimmed.StartsWith("}", StringComparison.Ordinal) || trimmed.StartsWith("else", StringComparison.Ordinal) ||
            trimmed.StartsWith("if ", StringComparison.Ordinal) || trimmed.StartsWith("if(", StringComparison.Ordinal) ||
            trimmed.StartsWith("private ", StringComparison.Ordinal))
            return false;
        if (AccessRegex.IsMatch(line) || trimmed.StartsWith("event ", StringComparison.Ordinal) ||
            trimmed.StartsWith("new ", StringComparison.Ordinal))
            return true;
        // Explicit interface implementations have no accessibility modifier.
        if (trimmed.Contains('.', StringComparison.Ordinal) &&
            (trimmed.Contains('(', StringComparison.Ordinal) || trimmed.Contains('[', StringComparison.Ordinal) ||
             trimmed.Contains('{', StringComparison.Ordinal) || !trimmed.EndsWith(";", StringComparison.Ordinal)))
            return true;
        return type.Kind == "interface" && (trimmed.Contains('(', StringComparison.Ordinal) ||
            trimmed.Contains(';', StringComparison.Ordinal) || trimmed.Contains('{', StringComparison.Ordinal));
    }

    private static bool IsPotentialDeclaration(string line)
    {
        var trimmed = line.Trim();
        return !trimmed.StartsWith("private ", StringComparison.Ordinal) &&
               (AccessRegex.IsMatch(line) || trimmed.StartsWith("event ", StringComparison.Ordinal) ||
                (line.Contains('.', StringComparison.Ordinal) && line.Contains('(', StringComparison.Ordinal)));
    }

    private static void AddUnknown(WrapperContract contract, string code, string message, string path, int line, string logicalId)
    {
        if (OwnershipClassifier.GetOwnership(path) != "wrapper-generated")
            return;
        contract.Unknowns.Add(new UnknownRecord { Code = code, Message = message, Path = path, Line = line, LogicalId = logicalId });
    }

    private static DocumentationRecord ParseDocumentation(List<string> lines, string path, int line, string logicalId, WrapperContract contract)
    {
        if (lines == null || lines.Count == 0)
            return null;
        var raw = string.Join("\n", lines.Select(x => x.Trim()));
        var result = new DocumentationRecord { Raw = raw };
        try
        {
            var xml = System.Xml.Linq.XElement.Parse("<root>" + raw + "</root>", System.Xml.Linq.LoadOptions.PreserveWhitespace);
            foreach (var element in xml.Elements())
            {
                var value = Normalize(element.Value);
                switch (element.Name.LocalName)
                {
                    case "summary": result.Summary = value; break;
                    case "remarks": result.Remarks = value; break;
                    case "returns": result.Returns = value; break;
                    case "value": result.Value = value; break;
                    case "param":
                        var parameter = (string)element.Attribute("name");
                        if (!string.IsNullOrWhiteSpace(parameter)) result.Parameters[parameter] = value;
                        break;
                    case "typeparam":
                        var typeParameter = (string)element.Attribute("name");
                        if (!string.IsNullOrWhiteSpace(typeParameter)) result.TypeParameters[typeParameter] = value;
                        break;
                    case "exception":
                        result.Exceptions.Add(value);
                        break;
                }
            }
        }
        catch (System.Xml.XmlException)
        {
            result.ParseStatus = "invalid";
            result.ParseError = "Malformed XML documentation.";
        }
        return result;
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
        var declaration = Normalize(RemoveDeclarationBody(line));
        declaration = Regex.Replace(declaration, @"^\s*(?:\[[^\]]*\]\s*)+", "");
        var accessMatch = AccessRegex.Match(declaration);
        var explicitInterface = !accessMatch.Success && IsExplicitInterfaceDeclaration(declaration);
        if (!accessMatch.Success && !explicitInterface && type.Kind != "interface")
            return null;

        var accessibility = accessMatch.Success ? accessMatch.Groups["access"].Value : "public";
        var rest = accessMatch.Success ? accessMatch.Groups["rest"].Value.Trim() : declaration;
        if (rest.Length == 0 || rest.StartsWith("if ", StringComparison.Ordinal) || rest.StartsWith("if(", StringComparison.Ordinal))
            return null;

        var modifiers = new List<string>();
        while (true)
        {
            var modifier = Regex.Match(rest, @"^(static|virtual|override|abstract|sealed|new|extern|unsafe|async|readonly|const|event|partial)\s+(.*)$");
            if (!modifier.Success) break;
            modifiers.Add(modifier.Groups[1].Value);
            rest = modifier.Groups[2].Value.Trim();
        }

        var builder = new MemberBuilder { Accessibility = accessibility, Modifiers = modifiers };
        var paren = FindTopLevelParen(rest);
        if (paren >= 0)
        {
            var before = rest[..paren].Trim();
            var methodName = ExtractMemberName(before, type.Name);
            builder.Name = methodName.Name;
            builder.Kind = string.Equals(methodName.Name, type.Name, StringComparison.Ordinal) ? "constructor" : "method";
            builder.ReturnType = builder.Kind == "constructor" ? null : methodName.Prefix;
            var close = FindClosingParen(rest, paren);
            if (close < 0)
                return null;
            builder.Parameters = Normalize(rest.Substring(paren, close - paren + 1));
            builder.DefaultValues = ParseDefaultValues(rest.Substring(paren + 1, close - paren - 1));
            builder.Signature = Normalize(accessibility + " " + string.Join(" ", modifiers) + " " + rest);
            return builder;
        }

        var withoutInitializer = RemoveInitializer(rest);
        var indexerMarker = Regex.Match(withoutInitializer, @"\bthis\s*\[", RegexOptions.Singleline);
        if (indexerMarker.Success)
        {
            var open = withoutInitializer.IndexOf('[', indexerMarker.Index);
            var close = FindClosingBracket(withoutInitializer, open);
            if (close < 0)
                return null;
            var parameters = withoutInitializer.Substring(open, close - open + 1);
            builder.Name = "this";
            builder.Kind = "indexer";
            builder.ReturnType = Normalize(withoutInitializer[..indexerMarker.Index]);
            builder.Parameters = Normalize(parameters);
            builder.DefaultValues = ParseDefaultValues(parameters.Trim('[', ']'));
        }
        else
        {
            var identifiers = IdentifierRegex.Matches(withoutInitializer).Cast<Match>().Select(m => m.Value).ToList();
            if (identifiers.Count < 1)
                return null;
            builder.Name = explicitInterface ? ExtractMemberName(withoutInitializer, null).Name : identifiers[^1];
            var nameIndex = withoutInitializer.LastIndexOf(builder.Name, StringComparison.Ordinal);
            builder.ReturnType = Normalize(withoutInitializer[..Math.Max(0, nameIndex)]);
            if (modifiers.Contains("event", StringComparer.Ordinal))
                builder.Kind = "event";
            else if (withoutInitializer.Contains('{', StringComparison.Ordinal) || withoutInitializer.Contains("=>", StringComparison.Ordinal) ||
                Regex.IsMatch(line, @"\b(get|set|init)\b"))
                builder.Kind = "property";
            else
                builder.Kind = "field";
        }

        builder.Signature = Normalize(accessibility + " " + string.Join(" ", modifiers) + " " + rest);
        return builder;
    }

    private static string RemoveDeclarationBody(string value)
    {
        var brace = FindTopLevel(value, '{');
        if (brace >= 0)
            value = value[..brace] + " {";
        var arrow = FindTopLevel(value, '=');
        if (arrow >= 0 && (arrow + 1 >= value.Length || value[arrow + 1] != '>'))
            value = value[..arrow].TrimEnd();
        return value.Trim().TrimEnd(';').Trim();
    }

    private static string RemoveInitializer(string value)
    {
        var equals = FindTopLevel(value, '=');
        return equals >= 0 ? value[..equals].Trim() : value;
    }

    private static bool IsExplicitInterfaceDeclaration(string value) =>
        !Regex.IsMatch(value, @"\b(class|interface|struct|enum|delegate)\b") &&
        value.Contains('.', StringComparison.Ordinal) &&
        (value.Contains('(', StringComparison.Ordinal) || value.Contains('{', StringComparison.Ordinal) ||
         value.Contains('[', StringComparison.Ordinal));

    private static (string Name, string Prefix) ExtractMemberName(string value, string typeName)
    {
        var identifierMatches = Regex.Matches(value, @"[A-Za-z_]\w*(?:\s*<[^>]+>)?");
        if (identifierMatches.Count == 0)
            return (string.Empty, string.Empty);
        var last = identifierMatches[^1];
        var name = last.Value.Trim();
        var before = value[..last.Index].Trim();
        if (before.EndsWith(".", StringComparison.Ordinal))
        {
            var prefixName = Regex.Match(before, @"(?<qualified>[A-Za-z_][\w.]*)\.$").Groups["qualified"].Value;
            if (prefixName.Length > 0)
                name = prefixName + "." + name;
        }
        var prefix = value[..Math.Max(0, value.LastIndexOf(name, StringComparison.Ordinal))].Trim();
        if (typeName != null && name.Contains('.', StringComparison.Ordinal))
            prefix = prefix.Replace(name[..(name.LastIndexOf(".", StringComparison.Ordinal) + 1)], "", StringComparison.Ordinal).Trim();
        return (name, prefix);
    }

    private static int FindTopLevelParen(string value) => FindTopLevel(value, '(');

    private static int FindTopLevel(string value, char target)
    {
        var angle = 0;
        var square = 0;
        var parenthesis = 0;
        for (var i = 0; i < value.Length; i++)
        {
            if (value[i] == target && angle == 0 && square == 0 && parenthesis == 0)
                return i;
            switch (value[i])
            {
                case '<': angle++; break;
                case '>': if (angle > 0) angle--; break;
                case '[': square++; break;
                case ']': if (square > 0) square--; break;
                case '(': parenthesis++; break;
                case ')': if (parenthesis > 0) parenthesis--; break;
            }
        }
        return -1;
    }

    private static Dictionary<string, string> ParseDefaultValues(string value)
    {
        var result = new Dictionary<string, string>(StringComparer.Ordinal);
        foreach (var parameter in SplitCommaSeparated(value))
        {
            var equals = FindTopLevel(parameter, '=');
            if (equals < 0) continue;
            var left = parameter[..equals].Trim();
            var right = Normalize(parameter[(equals + 1)..]);
            var names = IdentifierRegex.Matches(left).Cast<Match>().Select(m => m.Value).ToArray();
            if (names.Length > 0 && right.Length > 0)
                result[names[^1]] = right;
        }
        return result;
    }

    private static void ParseBases(TypeRecord type, string bases)
    {
        if (string.IsNullOrWhiteSpace(bases)) return;
        var values = SplitCommaSeparated(bases);
        if (type.Kind == "class" && values.Count > 0)
        {
            if (string.IsNullOrWhiteSpace(type.BaseType))
                type.BaseType = values[0];
            foreach (var value in values.Skip(1))
                if (!type.Interfaces.Contains(value, StringComparer.Ordinal))
                    type.Interfaces.Add(value);
        }
        else
        {
            foreach (var value in values)
                if (!type.Interfaces.Contains(value, StringComparer.Ordinal))
                    type.Interfaces.Add(value);
        }
    }

    private static bool HasOpenAttribute(IEnumerable<string> lines)
    {
        var balance = 0;
        foreach (var line in lines)
            balance += Count(line, '[') - Count(line, ']');
        return balance > 0;
    }

    private static List<string> ExtractAttributes(string text)
    {
        var values = new List<string>();
        foreach (Match match in Regex.Matches(text, @"\[(?<value>[^\]]+)\]"))
        {
            foreach (var attribute in SplitCommaSeparated(match.Groups["value"].Value))
            {
                var value = Normalize(attribute);
                if (value.Length > 0)
                    values.Add(value);
            }
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

    private static int FindClosingBracket(string text, int open)
    {
        var level = 0;
        for (var index = open; index < text.Length; index++)
        {
            if (text[index] == '[')
                level++;
            else if (text[index] == ']' && --level == 0)
                return index;
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
        var angle = 0;
        var parenthesis = 0;
        var square = 0;
        var braces = 0;
        var quoted = false;
        var start = 0;
        for (var i = 0; i < value.Length; i++)
        {
            if (value[i] == '"' && (i == 0 || value[i - 1] != '\\'))
                quoted = !quoted;
            if (quoted) continue;
            switch (value[i])
            {
                case '<': angle++; break;
                case '>': if (angle > 0) angle--; break;
                case '(': parenthesis++; break;
                case ')': if (parenthesis > 0) parenthesis--; break;
                case '[': square++; break;
                case ']': if (square > 0) square--; break;
                case '{': braces++; break;
                case '}': if (braces > 0) braces--; break;
                case ',' when angle == 0 && parenthesis == 0 && square == 0 && braces == 0:
                    result.Add(Normalize(value[start..i]));
                    start = i + 1;
                    break;
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
    public Dictionary<string, string> DefaultValues { get; set; } = new Dictionary<string, string>(StringComparer.Ordinal);
    public string Signature { get; set; }
    public string Source { get; set; }
    public int Line { get; set; }
    public List<string> Attributes { get; set; } = new List<string>();
    public List<string> InvocationText { get; set; } = new List<string>();
    public DocumentationRecord Documentation { get; set; }
    public List<string> Lines { get; } = new List<string>();

    public MemberRecord ToRecord() => new MemberRecord
    {
        Name = Name,
        Kind = Kind,
        Accessibility = Accessibility,
        Modifiers = Modifiers,
        ReturnType = ReturnType,
        Parameters = Parameters,
        DefaultValues = DefaultValues.OrderBy(x => x.Key, StringComparer.Ordinal).ToDictionary(x => x.Key, x => x.Value, StringComparer.Ordinal),
        Signature = Signature,
        Source = Source,
        Line = Line,
        Attributes = Attributes,
        InvocationText = InvocationText.Distinct(StringComparer.Ordinal).ToList(),
        Documentation = Documentation
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
    public List<string> SupportVersions { get; set; } = new List<string>();
    public string BaseType { get; set; }
    public List<string> Interfaces { get; set; } = new List<string>();
    public List<string> Attributes { get; set; } = new List<string>();
    public DocumentationRecord Documentation { get; set; }
    public string Source { get; set; }
    public int Line { get; set; }
    public string Signature { get; set; }
    public List<TypePartRecord> Parts { get; set; } = new List<TypePartRecord>();
    public List<MemberRecord> Members { get; set; } = new List<MemberRecord>();
}

internal sealed class TypePartRecord
{
    public string Source { get; set; }
    public int Line { get; set; }
    public string Signature { get; set; }
    public List<string> Attributes { get; set; } = new List<string>();
    public string BaseType { get; set; }
    public List<string> Interfaces { get; set; } = new List<string>();
    public DocumentationRecord Documentation { get; set; }
}

internal sealed class MemberRecord
{
    public string LogicalId { get; set; }
    public string Name { get; set; }
    public string Kind { get; set; }
    public string Accessibility { get; set; }
    public List<string> Modifiers { get; set; } = new List<string>();
    public string ReturnType { get; set; }
    public string Parameters { get; set; }
    public Dictionary<string, string> DefaultValues { get; set; } = new Dictionary<string, string>(StringComparer.Ordinal);
    public string Signature { get; set; }
    public List<string> Attributes { get; set; } = new List<string>();
    public List<string> InvocationText { get; set; } = new List<string>();
    public List<string> SupportVersions { get; set; } = new List<string>();
    public DocumentationRecord Documentation { get; set; }
    public string Source { get; set; }
    public int Line { get; set; }
}

internal sealed class DocumentationRecord
{
    public string ParseStatus { get; set; } = "parsed";
    public string ParseError { get; set; }
    public string Raw { get; set; }
    public string Summary { get; set; }
    public string Remarks { get; set; }
    public string Returns { get; set; }
    public string Value { get; set; }
    public Dictionary<string, string> Parameters { get; set; } = new Dictionary<string, string>(StringComparer.Ordinal);
    public Dictionary<string, string> TypeParameters { get; set; } = new Dictionary<string, string>(StringComparer.Ordinal);
    public List<string> Exceptions { get; set; } = new List<string>();
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
        var records = contract.Types.Select(type => new LedgerEntry
            {
                LogicalId = type.LogicalId,
                RecordKind = "type",
                Source = string.Join(";", type.Parts.Select(part => ContractIdentity.NormalizePartition(part.Source)).Distinct(StringComparer.Ordinal).OrderBy(value => value, StringComparer.Ordinal)),
                Status = "extracted",
                ExpectedMatchCount = 1
            })
            .Concat(contract.Types.SelectMany(type => type.Members.Select(member => new LedgerEntry
            {
                LogicalId = member.LogicalId, RecordKind = "member", Source = ContractIdentity.NormalizePartition(member.Source),
                Status = "extracted", ExpectedMatchCount = 1
            })));
        ledger.Entries = records.GroupBy(entry => entry.LogicalId + "\0" + entry.RecordKind, StringComparer.Ordinal)
            .Select(group => new LedgerEntry
            {
                LogicalId = group.First().LogicalId,
                RecordKind = group.First().RecordKind,
                Source = string.Join(";", group.Select(entry => entry.Source).Distinct(StringComparer.Ordinal).OrderBy(value => value, StringComparer.Ordinal)),
                Status = "extracted",
                ExpectedMatchCount = group.Count()
            })
            .OrderBy(entry => entry.LogicalId, StringComparer.Ordinal).ThenBy(entry => entry.RecordKind, StringComparer.Ordinal).ToList();
        return ledger;
    }

    public static ClassificationReport CreateClassification(WrapperContract contract)
    {
        var report = new ClassificationReport { SchemaVersion = "1.0", ContractKind = "NetOffice.WrapperClassification", Api = contract.Source.Api };
        foreach (var file in contract.Source.Files)
        {
            var fileTypes = contract.Types.Where(type => type.Parts.Any(part => string.Equals(part.Source, file.Path, StringComparison.Ordinal))).ToList();
            report.Files.Add(OwnershipClassifier.Classify(file, fileTypes));
        }
        foreach (var type in contract.Types)
        {
            foreach (var part in type.Parts)
            {
                var path = ContractIdentity.NormalizePartition(part.Source);
                var folder = path.Contains("/", StringComparison.Ordinal) ? path[..path.IndexOf('/', StringComparison.Ordinal)] : "Root";
                var ownership = report.Files.Single(file => file.Path == path).Ownership;
                report.Entries.Add(new ClassificationEntry
                {
                    LogicalId = type.LogicalId, Source = path, Category = folder, Kind = type.Kind, Ownership = ownership
                });
            }
        }
        report.Files = report.Files.OrderBy(file => file.Path, StringComparer.Ordinal).ToList();
        report.Entries = report.Entries.OrderBy(entry => entry.LogicalId, StringComparer.Ordinal)
            .ThenBy(entry => entry.Source, StringComparer.Ordinal).ToList();
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
    public string Facet { get; set; }
    public string Rationale { get; set; }
    public string Provenance { get; set; }
    public string ApprovedBy { get; set; }
}
internal sealed class ClassificationReport
{
    public string SchemaVersion { get; set; }
    public string ContractKind { get; set; }
    public string Api { get; set; }
    public List<FileClassification> Files { get; set; } = new List<FileClassification>();
    public List<ClassificationEntry> Entries { get; set; } = new List<ClassificationEntry>();
}
internal sealed class ClassificationEntry
{
    public string LogicalId { get; set; }
    public string Source { get; set; }
    public string Category { get; set; }
    public string Ownership { get; set; }
    public string Kind { get; set; }
}
internal sealed class FileClassification
{
    public string Path { get; set; }
    public string Ownership { get; set; }
    public string BuildAction { get; set; }
    public bool RequiredForIsolatedBuild { get; set; }
    public string Reason { get; set; }
    public List<string> TypeLogicalIds { get; set; } = new List<string>();
}


internal static class JsonFile
{
    internal static readonly JsonSerializerOptions SerializerOptions = new JsonSerializerOptions
    {
        WriteIndented = true,
        Encoder = System.Text.Encodings.Web.JavaScriptEncoder.UnsafeRelaxedJsonEscaping,
        DefaultIgnoreCondition = JsonIgnoreCondition.Never,
        PropertyNameCaseInsensitive = false
    };

    public static T Read<T>(string path)
    {
        var result = JsonSerializer.Deserialize<T>(File.ReadAllText(path), SerializerOptions);
        return result ?? throw new InvalidDataException("JSON artifact is empty: " + path);
    }

    public static string Serialize<T>(T value) =>
        JsonSerializer.Serialize(value, SerializerOptions).Replace("\r\n", "\n", StringComparison.Ordinal).Replace('\r', '\n') + "\n";

    public static void Write<T>(string path, T value)
    {
        Directory.CreateDirectory(Path.GetDirectoryName(Path.GetFullPath(path))!);
        File.WriteAllText(path, Serialize(value), new UTF8Encoding(false));
    }
}
