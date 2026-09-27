using System.Globalization;
using System.Security.Cryptography;
using System.Text;
using System.Text.RegularExpressions;
using NetOffice.CodeGen.Data;

namespace NetOffice.CodeGen.Projection;

/// <summary>Pure, deterministic Data v2 plus Wrapper Contract projection.</summary>
public static class ProjectionEngine
{
    private static readonly Regex SupportRegex = new(@"SupportByVersion\s*\(\s*""(?<product>[^""]+)""\s*,(?<versions>[^)]*)\)", RegexOptions.Compiled | RegexOptions.CultureInvariant);
    private static readonly Regex AttributeRegex = new("[A-Za-z_][A-Za-z0-9_]*", RegexOptions.Compiled | RegexOptions.CultureInvariant);

    public static ProjectionResult Project(DataGraph graph, WrapperContract contract, ProjectionPolicy policy, ProjectionOptions? options = null)
    {
        ArgumentNullException.ThrowIfNull(graph);
        ArgumentNullException.ThrowIfNull(contract);
        ArgumentNullException.ThrowIfNull(policy);
        options ??= new ProjectionOptions();
        policy.ValidateOrThrow(options.ExpectedPolicyDigest);

        var dataValidation = DataGraphValidator.Validate(graph, options.ExpectedDataDigest, policy.DataSchemaVersion);
        if (!dataValidation.IsValid)
            throw new ProjectionValidationException(dataValidation.Issues.Select(static x => new ProjectionIssue(x.Code, x.Message, x.Path ?? "")).ToArray());
        ValidateContract(contract, policy.ContractSchemaVersion, options.ExpectedContractDigest);
        if (options.RejectUnresolvedAmbiguities && graph.Ambiguities.Any(static x => string.IsNullOrWhiteSpace(x.Resolution)))
            throw new ProjectionValidationException(graph.Ambiguities.Where(static x => string.IsNullOrWhiteSpace(x.Resolution)).Select(static x => new ProjectionIssue("data.ambiguity", x.Reason, x.LogicalId)).ToArray());

        var issues = new List<ProjectionIssue>();
        var dataByName = graph.Types.GroupBy(static x => x.Name, StringComparer.Ordinal).ToDictionary(static x => x.Key, static x => x.ToArray(), StringComparer.Ordinal);
        var dataById = graph.Types.ToDictionary(static x => x.LogicalId, StringComparer.Ordinal);
        var contractByName = contract.Types.GroupBy(static x => x.Name, StringComparer.Ordinal).Where(static x => x.Count() == 1).ToDictionary(static x => x.Key, static x => x.Single(), StringComparer.Ordinal);
        var projected = new List<(ContractType Contract, DataType Data, ProjectedType Type)>();
        var traces = new List<ProjectionTrace>();

        foreach (var sourceType in contract.Types.OrderBy(static x => x.LogicalId, StringComparer.Ordinal))
        {
            if (!dataByName.TryGetValue(sourceType.Name, out var candidates) || candidates.Length == 0)
            {
                issues.Add(new("type.missing", $"No Data v2 type named {sourceType.Name} exists.", sourceType.LogicalId));
                continue;
            }
            if (candidates.Length != 1)
            {
                issues.Add(new("type.conflict", $"Data v2 type name {sourceType.Name} has {candidates.Length} candidates.", sourceType.LogicalId));
                continue;
            }
            var dataType = candidates[0];
            var steps = new List<ProjectionTraceStep>
            {
                new ProjectionTraceStep { Stage = "match", Rule = "contract-name-to-data-type", Result = dataType.LogicalId },
                new ProjectionTraceStep { Stage = "name", Rule = "sanitize-csharp-identifier", Result = Sanitize(sourceType.Name, policy.Names) },
                new ProjectionTraceStep { Stage = "inheritance", Rule = "expand-base-type-graph", Result = string.Join(",", dataType.BaseTypeIds.OrderBy(static x => x, StringComparer.Ordinal)) },
                new ProjectionTraceStep { Stage = "partition", Rule = policy.FilePartitionStrategy, Result = PartitionPath(sourceType, policy) }
            };
            var projectedType = new ProjectedType
            {
                LogicalId = sourceType.LogicalId,
                DataLogicalId = dataType.LogicalId,
                Namespace = sourceType.Namespace,
                Name = sourceType.Name,
                CSharpName = Sanitize(sourceType.Name, policy.Names),
                Kind = sourceType.Kind,
                BaseType = sourceType.BaseType,
                Interfaces = sourceType.Interfaces.OrderBy(static x => x, StringComparer.Ordinal).ToArray(),
                EffectiveBaseTypes = EffectiveBaseTypes(dataType, dataById, contractByName),
                DuplicateGroups = Array.Empty<string>(),
                Signature = sourceType.Signature,
                SupportVersions = ParseSupport(sourceType.Attributes),
                Capabilities = TypeCapabilities(sourceType, policy),
                RuntimeRequirements = RuntimeRequirements("type", sourceType.Kind, sourceType.Attributes, policy),
                DocsBindingKey = DocsKey(policy, dataType.LogicalId),
                Source = sourceType.Source,
                Line = sourceType.Line
            };
            traces.Add(new ProjectionTrace { LogicalId = sourceType.LogicalId, Steps = steps });
            projected.Add((sourceType, dataType, projectedType));
        }
        if (issues.Count != 0) throw new ProjectionValidationException(issues);

        var byDataTypeId = projected.ToDictionary(static x => x.Data.LogicalId, StringComparer.Ordinal);
        var files = new List<WrapperFile>();
        foreach (var item in projected.OrderBy(x => PartitionPath(x.Contract, policy), StringComparer.Ordinal).ThenBy(static x => x.Contract.LogicalId, StringComparer.Ordinal))
        {
            var memberCandidates = graph.Members.Where(x => x.TypeId == item.Data.LogicalId).GroupBy(static x => x.Name, StringComparer.Ordinal).ToDictionary(static x => x.Key, static x => x.OrderBy(y => y.LogicalId, StringComparer.Ordinal).ToArray(), StringComparer.Ordinal);
            var usedData = new HashSet<string>(StringComparer.Ordinal);
            var members = new List<ProjectedMember>();
            foreach (var contractMember in item.Contract.Members.OrderBy(static x => x.Name, StringComparer.Ordinal).ThenBy(static x => x.Signature, StringComparer.Ordinal))
            {
                if (!memberCandidates.TryGetValue(contractMember.Name, out var candidates) || candidates.Length == 0)
                {
                    issues.Add(new("member.missing", $"No Data v2 member named {contractMember.Name} exists on {item.Contract.Name}.", item.Contract.LogicalId));
                    continue;
                }
                var dataMember = ChooseMember(candidates, contractMember, usedData);
                if (dataMember is null)
                {
                    issues.Add(new("member.conflict", $"Member {contractMember.Name} has no unambiguous Data v2 candidate.", item.Contract.LogicalId));
                    continue;
                }
                usedData.Add(dataMember.LogicalId);
                var memberId = $"{item.Contract.LogicalId}/{StableMemberKey(contractMember)}";
                var overloadGroup = $"{item.Contract.LogicalId}:overload:{Sanitize(contractMember.Name, policy.Names)}";
                var capabilities = MemberCapabilities(contractMember, policy);
                var invocation = Invocation(contractMember, dataMember);
                var member = new ProjectedMember
                {
                    LogicalId = memberId,
                    DataLogicalId = dataMember.LogicalId,
                    Name = contractMember.Name,
                    CSharpName = Sanitize(contractMember.Name, policy.Names),
                    Kind = contractMember.Kind,
                    Accessibility = contractMember.Accessibility,
                    Modifiers = contractMember.Modifiers.OrderBy(static x => x, StringComparer.Ordinal).ToArray(),
                    ReturnType = contractMember.ReturnType ?? dataMember.ReturnType,
                    Parameters = contractMember.Parameters,
                    Signature = contractMember.Signature,
                    OverloadGroup = overloadGroup,
                    DuplicateOf = DuplicateOf(dataMember, candidates),
                    Invocation = invocation,
                    SupportVersions = ParseSupport(contractMember.Attributes),
                    Capabilities = capabilities,
                    RuntimeRequirements = RuntimeRequirements("member", contractMember.Kind, contractMember.Attributes, policy).Concat(capabilities.Select(static x => "capability:" + x)).Distinct(StringComparer.Ordinal).OrderBy(static x => x, StringComparer.Ordinal).ToArray(),
                    DocsBindingKey = DocsKey(policy, dataMember.LogicalId),
                    Attributes = contractMember.Attributes.OrderBy(static x => x, StringComparer.Ordinal).ToArray(),
                    Source = contractMember.Source,
                    Line = contractMember.Line
                };
                members.Add(member);
                traces.Add(new ProjectionTrace
                {
                    LogicalId = memberId,
                    Steps = new[]
                    {
                        new ProjectionTraceStep { Stage = "match", Rule = "contract-member-to-data-member", Result = dataMember.LogicalId },
                        new ProjectionTraceStep { Stage = "overload", Rule = "group-by-name", Result = overloadGroup },
                        new ProjectionTraceStep { Stage = "invocation", Rule = "derive-dispatch-plan", Result = invocation.Operation },
                        new ProjectionTraceStep { Stage = "docs", Rule = "stable-logical-binding-key", Result = member.DocsBindingKey }
                    }
                });
            }
            if (issues.Count != 0) throw new ProjectionValidationException(issues);
            var duplicateGroups = members.Where(static x => x.DuplicateOf is not null).Select(static x => x.OverloadGroup).Distinct(StringComparer.Ordinal).OrderBy(static x => x, StringComparer.Ordinal).ToArray();
            var finalType = item.Type with { Members = members, DuplicateGroups = duplicateGroups };
            finalType = ApplyTypeOverrides(finalType, policy, issues);
            var path = PartitionPath(item.Contract, policy);
            files.Add(new WrapperFile { Path = path, Namespace = finalType.Namespace, Types = new[] { finalType } });
        }
        ApplyMemberOverrides(files, policy, issues);
        if (issues.Count != 0) throw new ProjectionValidationException(issues);
        return new ProjectionResult
        {
            DataSchemaVersion = graph.SchemaVersion,
            DataDigest = graph.Digest,
            ContractSchemaVersion = contract.SchemaVersion,
            ContractDigest = ContractDigest(contract),
            PolicyDigest = policy.Digest,
            Files = files.OrderBy(static x => x.Path, StringComparer.Ordinal).ToArray(),
            Explain = traces.OrderBy(static x => x.LogicalId, StringComparer.Ordinal).ToArray()
        };
    }

    public static string ProjectJson(string dataGraphJson, string contractJson, string policyJson, ProjectionOptions? options = null, bool indented = false)
    {
        var graph = CanonicalJson.Parse(dataGraphJson);
        var contract = WrapperContract.Parse(contractJson);
        var policy = ProjectionPolicy.Parse(policyJson);
        return Project(graph, contract, policy, options).ToJson(indented);
    }

    public static string Serialize(ProjectionResult result, bool indented = false) => result.ToJson(indented);

    private static void ValidateContract(WrapperContract contract, string expectedSchema, string? expectedDigest)
    {
        var issues = new List<ProjectionIssue>();
        if (!string.Equals(contract.SchemaVersion, expectedSchema, StringComparison.Ordinal)) issues.Add(new("contract.schema", $"Expected {expectedSchema}, got {contract.SchemaVersion}.", "schemaVersion"));
        if (!string.Equals(contract.ContractKind, "NetOffice.WrapperContract", StringComparison.Ordinal)) issues.Add(new("contract.kind", "Unexpected contract kind.", "contractKind"));
        if (contract.Source is null || string.IsNullOrWhiteSpace(contract.Source.Api)) issues.Add(new("contract.source", "Source.Api is required.", "source.api"));
        if (contract.Types is null) issues.Add(new("contract.types", "Types is required.", "types"));
        else
        {
            var duplicateTypes = contract.Types.GroupBy(static x => x.LogicalId, StringComparer.Ordinal).Where(static x => string.IsNullOrWhiteSpace(x.Key) || x.Count() > 1);
            foreach (var duplicate in duplicateTypes) issues.Add(new("contract.type-id", "Contract type logical IDs must be unique and non-empty.", duplicate.Key));
            foreach (var type in contract.Types)
            {
                if (string.IsNullOrWhiteSpace(type.LogicalId) || string.IsNullOrWhiteSpace(type.Name) || string.IsNullOrWhiteSpace(type.Signature) || string.IsNullOrWhiteSpace(type.Source) || type.Line < 1)
                    issues.Add(new("contract.type-shape", $"Contract type {type.LogicalId} is incomplete.", type.LogicalId));
                if (type.Members is null) { issues.Add(new("contract.members", "Members is required.", type.LogicalId)); continue; }
                foreach (var member in type.Members)
                    if (string.IsNullOrWhiteSpace(member.Name) || string.IsNullOrWhiteSpace(member.Kind) || string.IsNullOrWhiteSpace(member.Signature) || string.IsNullOrWhiteSpace(member.Source) || member.Line < 1)
                        issues.Add(new("contract.member-shape", $"Contract member {member.Name} is incomplete.", type.LogicalId));
            }
        }
        if (contract.Unknowns is null || contract.Unknowns.Count != 0) issues.Add(new("contract.unknown", "Contract contains unknown records and cannot be projected safely.", "unknowns"));
        if (contract.Ambiguities is null || contract.Ambiguities.Count != 0) issues.Add(new("contract.ambiguity", "Contract contains unresolved ambiguity records.", "ambiguities"));
        var digest = ContractDigest(contract);
        if (expectedDigest is not null && !string.Equals(expectedDigest, digest, StringComparison.OrdinalIgnoreCase)) issues.Add(new("contract.expected-digest", $"Expected pinned digest {expectedDigest}, got {digest}.", "digest"));
        if (issues.Count != 0) throw new ProjectionValidationException(issues);
    }

    public static string ContractDigest(WrapperContract contract)
    {
        var json = System.Text.Json.JsonSerializer.Serialize(contract, ProjectionPolicy.JsonOptions);
        return Convert.ToHexString(SHA256.HashData(Encoding.UTF8.GetBytes(json))).ToLowerInvariant();
    }

    private static DataMember? ChooseMember(IReadOnlyList<DataMember> candidates, ContractMember contract, HashSet<string> used)
    {
        var unused = candidates.Where(x => !used.Contains(x.LogicalId)).ToArray();
        if (unused.Length == 1) return unused[0];
        if (unused.Length == 0) return null;
        var parameterCount = ParameterCount(contract.Parameters);
        var matching = unused.Where(x => x.ParameterTypes.Count == parameterCount).ToArray();
        return matching.Length == 1 ? matching[0] : unused.OrderBy(static x => x.LogicalId, StringComparer.Ordinal).FirstOrDefault();
    }

    private static string? DuplicateOf(DataMember member, IReadOnlyList<DataMember> siblings)
        => siblings.Where(x => x.Name == member.Name && x.LogicalId != member.LogicalId).OrderBy(static x => x.LogicalId, StringComparer.Ordinal).Select(static x => x.LogicalId).FirstOrDefault();

    private static string StableMemberKey(ContractMember member) => member.Name + (member.Parameters ?? "") + ":" + member.Kind;

    private static int ParameterCount(string? parameters)
    {
        if (string.IsNullOrWhiteSpace(parameters) || parameters.Trim() == "()") return 0;
        var value = parameters.Trim();
        if (value.StartsWith("(", StringComparison.Ordinal) && value.EndsWith(")", StringComparison.Ordinal)) value = value[1..^1];
        var depth = 0; var count = 1; var quote = false;
        foreach (var c in value)
        {
            if (c == '"') quote = !quote;
            if (quote) continue;
            if (c == '<') depth++; else if (c == '>') depth--; else if (c == ',' && depth == 0) count++;
        }
        return string.IsNullOrWhiteSpace(value) ? 0 : count;
    }

    private static InvocationPlan Invocation(ContractMember member, DataMember data)
    {
        var kind = member.Kind.ToLowerInvariant();
        var operation = kind.Contains("event", StringComparison.Ordinal) ? "event" : kind.Contains("property", StringComparison.Ordinal) ? "property" : kind.Contains("field", StringComparison.Ordinal) ? "field" : "method";
        if (member.Name.Contains("get", StringComparison.OrdinalIgnoreCase) && kind.Contains("accessor", StringComparison.Ordinal)) operation = "get";
        if (member.Name.Contains("set", StringComparison.OrdinalIgnoreCase) && kind.Contains("accessor", StringComparison.Ordinal)) operation = "set";
        if (member.Attributes.Any(x => x.Contains("IndexProperty", StringComparison.OrdinalIgnoreCase)) || member.Signature.Contains("this[", StringComparison.Ordinal)) operation = "indexer";
        return new InvocationPlan { Operation = operation, DispatchName = member.Name, DispId = data.DispId, ArgumentCount = data.ParameterTypes.Count, ResultType = member.ReturnType ?? data.ReturnType, RequiresProxy = operation is "method" or "property" or "indexer", Text = member.InvocationText.OrderBy(static x => x, StringComparer.Ordinal).ToArray() };
    }

    private static IReadOnlyList<string> TypeCapabilities(ContractType type, ProjectionPolicy policy)
    {
        var caps = new HashSet<string>(StringComparer.Ordinal);
        if (type.Attributes.Any(x => x.Contains("HasIndexProperty", StringComparison.OrdinalIgnoreCase))) caps.Add("indexer");
        if (type.Interfaces.Any(x => x.Contains("IEnumerable", StringComparison.OrdinalIgnoreCase)) || type.Attributes.Any(x => x.Contains("Enumerator", StringComparison.OrdinalIgnoreCase))) caps.Add("enumerator");
        if (type.Attributes.Any(x => x.Contains("Collection", StringComparison.OrdinalIgnoreCase))) caps.Add("collection");
        if (type.Members.Any(x => x.Kind.Contains("event", StringComparison.OrdinalIgnoreCase) || x.Attributes.Any(a => a.Contains("Event", StringComparison.OrdinalIgnoreCase)))) caps.Add("event");
        return AddPolicyCapabilities(caps, "type", policy);
    }

    private static IReadOnlyList<string> MemberCapabilities(ContractMember member, ProjectionPolicy policy)
    {
        var caps = new HashSet<string>(StringComparer.Ordinal);
        if (member.Kind.Contains("event", StringComparison.OrdinalIgnoreCase) || member.Attributes.Any(x => x.Contains("Event", StringComparison.OrdinalIgnoreCase))) caps.Add("event");
        if (member.Name.Contains("GetEnumerator", StringComparison.OrdinalIgnoreCase) || (member.ReturnType?.Contains("IEnumerator", StringComparison.OrdinalIgnoreCase) ?? false) || member.Attributes.Any(x => x.Contains("Enumerator", StringComparison.OrdinalIgnoreCase))) caps.Add("enumerator");
        if (member.Attributes.Any(x => x.Contains("IndexProperty", StringComparison.OrdinalIgnoreCase)) || member.Signature.Contains("this[", StringComparison.Ordinal)) caps.Add("indexer");
        return AddPolicyCapabilities(caps, member.Kind, policy);
    }

    private static IReadOnlyList<string> AddPolicyCapabilities(HashSet<string> caps, string key, ProjectionPolicy policy)
    {
        if (policy.RuntimeCapabilities.TryGetValue(key, out var value)) foreach (var item in value.Split(',', StringSplitOptions.RemoveEmptyEntries | StringSplitOptions.TrimEntries)) caps.Add(item);
        return caps.OrderBy(static x => x, StringComparer.Ordinal).ToArray();
    }

    private static IReadOnlyList<string> RuntimeRequirements(string scope, string kind, IReadOnlyList<string> attributes, ProjectionPolicy policy)
    {
        var result = new HashSet<string>(StringComparer.Ordinal);
        var key = scope + ":" + kind;
        if (policy.RuntimeCapabilities.TryGetValue(key, out var value)) foreach (var item in value.Split(',', StringSplitOptions.RemoveEmptyEntries | StringSplitOptions.TrimEntries)) result.Add(item);
        foreach (var attribute in attributes.SelectMany(static x => AttributeRegex.Matches(x).Select(m => m.Value)).OrderBy(static x => x, StringComparer.Ordinal))
            if (policy.RuntimeCapabilities.TryGetValue("attribute:" + attribute, out var requirement)) result.Add(requirement);
        return result.OrderBy(static x => x, StringComparer.Ordinal).ToArray();
    }

    private static IReadOnlyList<VersionSupport> ParseSupport(IEnumerable<string> attributes)
    {
        var result = new Dictionary<string, HashSet<string>>(StringComparer.Ordinal);
        foreach (var attribute in attributes)
            foreach (Match match in SupportRegex.Matches(attribute))
            {
                var product = match.Groups["product"].Value;
                if (!result.TryGetValue(product, out var versions)) result[product] = versions = new HashSet<string>(StringComparer.Ordinal);
                foreach (var version in match.Groups["versions"].Value.Split(',', StringSplitOptions.RemoveEmptyEntries | StringSplitOptions.TrimEntries)) versions.Add(version);
            }
        return result.OrderBy(static x => x.Key, StringComparer.Ordinal).Select(static x => new VersionSupport { Product = x.Key, Versions = x.Value.OrderBy(static y => y, StringComparer.Ordinal).ToArray() }).ToArray();
    }

    private static IReadOnlyList<string> EffectiveBaseTypes(DataType type, IReadOnlyDictionary<string, DataType> all, IReadOnlyDictionary<string, ContractType> contracts)
    {
        var result = new List<string>(); var seen = new HashSet<string>(StringComparer.Ordinal);
        void Visit(DataType current)
        {
            foreach (var id in current.BaseTypeIds.OrderBy(static x => x, StringComparer.Ordinal))
            {
                if (!seen.Add(id) || !all.TryGetValue(id, out var baseType)) continue;
                result.Add(contracts.TryGetValue(baseType.Name, out var c) ? c.Name : baseType.Name);
                Visit(baseType);
            }
        }
        Visit(type); return result;
    }

    private static string PartitionPath(ContractType type, ProjectionPolicy policy)
    {
        if (string.Equals(policy.FilePartitionStrategy, "source", StringComparison.Ordinal) && !string.IsNullOrWhiteSpace(type.Source)) return type.Source.Replace('\\', '/');
        var ns = string.Join('/', (type.Namespace ?? "").Split('.', StringSplitOptions.RemoveEmptyEntries).Select(SafePathPart));
        var root = policy.GeneratedRoot.Trim('/').Replace('\\', '/');
        return string.Join('/', new[] { root, ns, Sanitize(type.Name, policy.Names) + policy.FileExtension }.Where(static x => !string.IsNullOrEmpty(x)));
    }

    private static string DocsKey(ProjectionPolicy policy, string id) => policy.DocsKeyPrefix + ":" + id.Replace("/", ".", StringComparison.Ordinal);

    private static string Sanitize(string value, ProjectionNameRules rules)
    {
        var replacement = string.IsNullOrEmpty(rules.InvalidCharacterReplacement) ? "_" : rules.InvalidCharacterReplacement;
        var chars = value.Select(c => char.IsLetterOrDigit(c) || c == '_' ? c : replacement[0]).ToArray();
        var result = new string(chars);
        if (result.Length == 0) result = "_";
        if (char.IsDigit(result[0])) result = "_" + result;
        if ((rules.ReservedWords ?? Array.Empty<string>()).Contains(result, StringComparer.Ordinal)) result = "@" + result;
        return result;
    }

    private static string SafePathPart(string value) => Sanitize(value, new ProjectionNameRules { ReservedWords = Array.Empty<string>() });

    private static ProjectedType ApplyTypeOverrides(ProjectedType type, ProjectionPolicy policy, ICollection<ProjectionIssue> issues)
    {
        foreach (var item in policy.Overrides.Where(static x => x.MemberLogicalId is null && x.MemberName is null))
        {
            var matched = MatchesType(item, type);
            if (matched != (item.ExpectedMatches == 1))
            {
                issues.Add(new("policy.override.stale", $"Override {item.Id ?? item.Property} matched {(matched ? 1 : 0)}, expected {item.ExpectedMatches}.", item.Id ?? item.Property));
                continue;
            }
            if (!matched) continue;
            if (!string.Equals(item.Property, "csharpName", StringComparison.Ordinal) && !string.Equals(item.Property, "docsBindingKey", StringComparison.Ordinal) && !string.Equals(item.Property, "runtimeCapability", StringComparison.Ordinal) && !string.Equals(item.Property, "namespace", StringComparison.Ordinal)) continue;
            type = item.Property switch
            {
                "csharpName" => type with { CSharpName = item.Value },
                "docsBindingKey" => type with { DocsBindingKey = item.Value },
                "namespace" => type with { Namespace = item.Value },
                "runtimeCapability" => type with { RuntimeRequirements = type.RuntimeRequirements.Concat(new[] { item.Value }).Distinct(StringComparer.Ordinal).OrderBy(static x => x, StringComparer.Ordinal).ToArray() },
                _ => type
            };
        }
        return type;
    }

    private static void ApplyMemberOverrides(List<WrapperFile> files, ProjectionPolicy policy, ICollection<ProjectionIssue> issues)
    {
        foreach (var item in policy.Overrides)
        {
            var matches = files.SelectMany(f => f.Types.SelectMany(t => t.Members.Select(m => (File: f, Type: t, Member: m))))
                .Where(x => MatchesMember(item, x.Type, x.Member)).ToArray();
            if (item.MemberLogicalId is not null || item.MemberName is not null)
            {
                if (matches.Length != item.ExpectedMatches) issues.Add(new("policy.override.stale", $"Override {item.Id ?? item.Property} matched {matches.Length}, expected {item.ExpectedMatches}.", item.Id ?? item.Property));
                foreach (var match in matches)
                {
                    var member = match.Member;
                    member = item.Property switch
                    {
                        "csharpName" => member with { CSharpName = item.Value },
                        "signature" => member with { Signature = item.Value },
                        "docsBindingKey" => member with { DocsBindingKey = item.Value },
                        "invocationOperation" => member with { Invocation = member.Invocation with { Operation = item.Value } },
                        "runtimeCapability" => member with { RuntimeRequirements = member.RuntimeRequirements.Concat(new[] { item.Value }).Distinct(StringComparer.Ordinal).OrderBy(static x => x, StringComparer.Ordinal).ToArray() },
                        _ => member
                    };
                    var type = match.Type with { Members = match.Type.Members.Select(x => x.LogicalId == member.LogicalId ? member : x).ToArray() };
                    var index = files.IndexOf(match.File);
                    files[index] = match.File with { Types = match.File.Types.Select(x => x.LogicalId == type.LogicalId ? type : x).ToArray() };
                }
            }
        }
    }

    private static bool MatchesType(ProjectionOverride item, ProjectedType type)
        => item.TypeLogicalId is not null && (string.Equals(item.TypeLogicalId, type.LogicalId, StringComparison.Ordinal) || string.Equals(item.TypeLogicalId, type.DataLogicalId, StringComparison.Ordinal));

    private static bool MatchesMember(ProjectionOverride item, ProjectedType type, ProjectedMember member)
    {
        if (item.TypeLogicalId is not null && !MatchesType(item, type)) return false;
        if (item.MemberLogicalId is not null && !string.Equals(item.MemberLogicalId, member.LogicalId, StringComparison.Ordinal) && !string.Equals(item.MemberLogicalId, member.DataLogicalId, StringComparison.Ordinal)) return false;
        if (item.MemberName is not null && !string.Equals(item.MemberName, member.Name, StringComparison.Ordinal)) return false;
        return item.MemberLogicalId is not null || item.MemberName is not null;
    }
}
