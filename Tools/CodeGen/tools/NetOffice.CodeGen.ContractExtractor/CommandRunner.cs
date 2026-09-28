// Copyright (c) 2026 NetOffice contributors
// SPDX-License-Identifier: MIT

using System.Text.Json;
using System.Text.RegularExpressions;

namespace NetOffice.CodeGen.ContractExtractor;

internal sealed class CommandLineException : Exception
{
    public CommandLineException(string message) : base(message) { }
}

internal static class CommandRunner
{
    internal static readonly string[] ApiProducts =
    {
        "ADODB", "Access", "DAO", "Excel", "MSComctlLib", "MSDATASRC",
        "Office", "Outlook", "OWC10", "PowerPoint", "VBIDE", "Word"
    };

    public const string Usage =
        "Usage:\n" +
        "  ContractExtractor extract --source <Source> --api <Api> [--output <file>] [--ledger <file>] [--classification <file>]\n" +
        "  ContractExtractor extract-all --source <Source> --output-dir <contracts> [--apis <comma-list>]\n" +
        "  ContractExtractor validate --contracts <contracts> [--source <Source>] [--apis <comma-list>]\n" +
        "  ContractExtractor compare --expected <contracts> --actual <generated-tree> --report <file> [--apis <comma-list>]";

    public static int Run(string[] args)
    {
        if (args.Length == 0 || args[0] is "-h" or "--help")
        {
            Console.WriteLine(Usage);
            return 0;
        }

        var command = args[0].StartsWith("--", StringComparison.Ordinal) ? "extract" : args[0];
        var start = command == "extract" && args[0].StartsWith("--", StringComparison.Ordinal) ? 0 : 1;
        var options = ParseOptions(args, start);
        return command switch
        {
            "extract" => ExtractOne(options),
            "extract-all" => ExtractAll(options),
            "validate" => Validate(options),
            "compare" => Compare(options),
            _ => throw new CommandLineException("Unknown command: " + command)
        };
    }

    private static int ExtractOne(IReadOnlyDictionary<string, string> options)
    {
        var source = RequiredDirectory(options, "--source");
        var api = Required(options, "--api");
        EnsureKnownApi(api);
        var apiRoot = Path.Combine(source, api);
        if (!Directory.Exists(apiRoot))
            throw new CommandLineException("API source directory does not exist: " + apiRoot);

        var output = Get(options, "--output") ?? Path.Combine("contracts", "wrapper", api + ".wrapper-contract.json");
        WriteArtifacts(source, api, apiRoot, output, Get(options, "--ledger"), Get(options, "--classification"));
        return 0;
    }

    private static int ExtractAll(IReadOnlyDictionary<string, string> options)
    {
        var source = RequiredDirectory(options, "--source");
        var outputDirectory = Path.GetFullPath(Required(options, "--output-dir"));
        Directory.CreateDirectory(outputDirectory);
        var apis = SelectedApis(options);
        foreach (var api in apis)
        {
            var apiRoot = Path.Combine(source, api);
            if (!Directory.Exists(apiRoot))
                throw new CommandLineException("API source directory does not exist: " + apiRoot);
            WriteArtifacts(source, api, apiRoot, Path.Combine(outputDirectory, api + ".wrapper-contract.json"), null, null);
        }
        Console.WriteLine("Extracted deterministic wrapper contracts for " + apis.Count + " products.");
        return 0;
    }

    private static int Validate(IReadOnlyDictionary<string, string> options)
    {
        var contracts = RequiredDirectory(options, "--contracts");
        var source = Get(options, "--source");
        if (source != null)
            source = RequiredDirectory(options, "--source");
        var apis = SelectedApis(options);
        var failures = new List<string>();
        foreach (var api in apis)
        {
            var contractPath = Path.Combine(contracts, api + ".wrapper-contract.json");
            var classificationPath = Path.Combine(contracts, api + ".classification.json");
            var ledgerPath = Path.Combine(contracts, api + ".compatibility-ledger.json");
            ValidateExists(contractPath, failures);
            ValidateExists(classificationPath, failures);
            ValidateExists(ledgerPath, failures);
            if (!File.Exists(contractPath) || !File.Exists(classificationPath) || !File.Exists(ledgerPath))
                continue;

            var contract = JsonFile.Read<WrapperContract>(contractPath);
            var classification = JsonFile.Read<ClassificationReport>(classificationPath);
            var ledger = JsonFile.Read<CompatibilityLedger>(ledgerPath);
            ContractValidator.Validate(api, contract, classification, ledger, failures);
            if (source != null)
            {
                var regenerated = Extractor.Extract(source, api, Path.Combine(source, api));
                if (!string.Equals(JsonFile.Serialize(contract), JsonFile.Serialize(regenerated), StringComparison.Ordinal))
                    failures.Add(api + ": checked-in contract differs from deterministic extraction of current source.");
                var regeneratedClassification = Records.CreateClassification(regenerated);
                if (!string.Equals(JsonFile.Serialize(classification), JsonFile.Serialize(regeneratedClassification), StringComparison.Ordinal))
                    failures.Add(api + ": checked-in classification differs from deterministic extraction of current source.");
            }
        }

        foreach (var failure in failures.OrderBy(value => value, StringComparer.Ordinal))
            Console.Error.WriteLine("error: " + failure);
        if (failures.Count > 0)
            return 1;
        Console.WriteLine("Validated deterministic wrapper contracts for " + apis.Count + " products.");
        return 0;
    }

    private static int Compare(IReadOnlyDictionary<string, string> options)
    {
        var expected = RequiredDirectory(options, "--expected");
        var actual = RequiredDirectory(options, "--actual");
        var reportPath = Path.GetFullPath(Required(options, "--report"));
        var apis = SelectedApis(options, expected);
        var report = SemanticComparer.Compare(expected, actual, apis);
        JsonFile.Write(reportPath, report);
        var reportDirectory = Path.GetDirectoryName(reportPath)!;
        Directory.CreateDirectory(reportDirectory);
        foreach (var product in report.Products)
            JsonFile.Write(Path.Combine(reportDirectory, product.Api + ".semantic-diff.json"), product);

        Console.WriteLine("Compared " + report.Products.Count + " products: " + report.Summary.UnexplainedDifferences + " unexplained, " + report.Summary.IntentionalDifferences + " intentional differences.");
        Console.WriteLine("Report: " + reportPath);
        return report.Summary.UnexplainedDifferences == 0 ? 0 : 3;
    }

    private static void WriteArtifacts(string source, string api, string apiRoot, string output, string ledger, string classification)
    {
        var contract = Extractor.Extract(source, api, apiRoot);
        var outputPath = Path.GetFullPath(output);
        var directory = Path.GetDirectoryName(outputPath)!;
        Directory.CreateDirectory(directory);
        var stem = Path.GetFileNameWithoutExtension(outputPath);
        if (stem.EndsWith(".wrapper-contract", StringComparison.Ordinal))
            stem = stem[..^".wrapper-contract".Length];
        var ledgerPath = Path.GetFullPath(ledger ?? Path.Combine(directory, stem + ".compatibility-ledger.json"));
        var classificationPath = Path.GetFullPath(classification ?? Path.Combine(directory, stem + ".classification.json"));
        JsonFile.Write(outputPath, contract);
        JsonFile.Write(ledgerPath, Records.CreateLedger(contract));
        JsonFile.Write(classificationPath, Records.CreateClassification(contract));
        Console.WriteLine("Extracted " + contract.Types.Count + " type declarations and " + contract.Types.Sum(type => type.Members.Count) + " members from " + api + ".");
    }

    private static Dictionary<string, string> ParseOptions(string[] args, int start)
    {
        var result = new Dictionary<string, string>(StringComparer.Ordinal);
        for (var index = start; index < args.Length; index += 2)
        {
            if (!args[index].StartsWith("--", StringComparison.Ordinal) || index + 1 >= args.Length)
                throw new CommandLineException("Unknown or incomplete argument: " + args[index]);
            if (!result.TryAdd(args[index], args[index + 1]))
                throw new CommandLineException("Duplicate argument: " + args[index]);
        }
        return result;
    }

    private static List<string> SelectedApis(IReadOnlyDictionary<string, string> options, string expectedDirectory = null)
    {
        var selected = Get(options, "--apis");
        if (selected == null && expectedDirectory != null)
        {
            var discovered = Directory.EnumerateFiles(expectedDirectory, "*.wrapper-contract.json", SearchOption.TopDirectoryOnly)
                .Select(path => Path.GetFileName(path)[..^".wrapper-contract.json".Length])
                .Where(api => ApiProducts.Contains(api, StringComparer.Ordinal))
                .OrderBy(api => api, StringComparer.Ordinal).ToList();
            if (discovered.Count > 0)
                return discovered;
        }
        if (selected == null)
            return ApiProducts.ToList();
        var apis = selected.Split(',', StringSplitOptions.RemoveEmptyEntries | StringSplitOptions.TrimEntries)
            .Distinct(StringComparer.Ordinal).OrderBy(api => api, StringComparer.Ordinal).ToList();
        foreach (var api in apis)
            EnsureKnownApi(api);
        return apis;
    }

    private static void EnsureKnownApi(string api)
    {
        if (!ApiProducts.Contains(api, StringComparer.Ordinal))
            throw new CommandLineException("Unknown API product: " + api);
    }

    private static string Required(IReadOnlyDictionary<string, string> options, string name) =>
        Get(options, name) ?? throw new CommandLineException(name + " is required.");

    private static string RequiredDirectory(IReadOnlyDictionary<string, string> options, string name)
    {
        var path = Path.GetFullPath(Required(options, name));
        if (!Directory.Exists(path))
            throw new CommandLineException(name + " directory does not exist: " + path);
        return path;
    }

    private static string Get(IReadOnlyDictionary<string, string> options, string name) =>
        options.TryGetValue(name, out var value) ? value : null;

    private static void ValidateExists(string path, ICollection<string> failures)
    {
        if (!File.Exists(path))
            failures.Add("Missing required artifact: " + path.Replace('\\', '/'));
    }
}

internal static class ContractIdentity
{
    public static string MemberLogicalId(string typeLogicalId, MemberRecord member) =>
        typeLogicalId + "::" + member.Kind + ":" + member.Name + (member.Parameters ?? string.Empty);

    public static List<string> GetSupportVersions(IEnumerable<string> attributes)
    {
        var result = new List<string>();
        foreach (var attribute in attributes)
        {
            var match = Regex.Match(attribute, @"(?:^|\.)SupportByVersion(?:Attribute)?\((?<arguments>.*)\)$");
            if (match.Success)
                result.Add(match.Groups["arguments"].Value);
        }
        return result.Distinct(StringComparer.Ordinal).OrderBy(value => value, StringComparer.Ordinal).ToList();
    }

    public static string NormalizePartition(string path)
    {
        path = path.Replace('\\', '/');
        if (path.StartsWith("Generated/", StringComparison.Ordinal))
            path = path["Generated/".Length..];
        return path;
    }
}

internal static class OwnershipClassifier
{
    private static readonly HashSet<string> GeneratedFolders = new(StringComparer.Ordinal)
    {
        "Classes", "Constants", "DispatchInterfaces", "Enums", "Events", "Interfaces", "Modules"
    };

    public static FileClassification Classify(SourceFile file, IEnumerable<TypeRecord> types)
    {
        var path = ContractIdentity.NormalizePartition(file.Path);
        var ownership = GetOwnership(path);
        var reason = ownership switch
        {
            "wrapper-generated" => "Canonical wrapper category emitted from the logical API graph.",
            "runtime" => "Hand-maintained interop/runtime capability required by the API project.",
            "companion" => "Hand-maintained API companion source required beside generated wrappers.",
            _ => "Hand-maintained project source outside generator ownership."
        };
        return new FileClassification
        {
            Path = path,
            Ownership = ownership,
            BuildAction = ownership == "wrapper-generated" ? "generate" : "copy-companion",
            RequiredForIsolatedBuild = true,
            Reason = reason,
            TypeLogicalIds = types.Select(type => type.LogicalId).Distinct(StringComparer.Ordinal).OrderBy(value => value, StringComparer.Ordinal).ToList()
        };
    }

    public static string GetOwnership(string sourcePath)
    {
        var path = ContractIdentity.NormalizePartition(sourcePath);
        var first = path.Contains('/', StringComparison.Ordinal) ? path[..path.IndexOf('/', StringComparison.Ordinal)] : string.Empty;
        return GeneratedFolders.Contains(first) || string.Equals(path, "Utils/ProjectInfo.cs", StringComparison.Ordinal)
            ? "wrapper-generated"
            : first is "Native" or "NativeCaller" ? "runtime"
            : first is "Tools" or "Properties" or "Utils" ? "companion"
            : "manual";
    }
}

internal static class ContractValidator
{
    public static void Validate(string api, WrapperContract contract, ClassificationReport classification, CompatibilityLedger ledger, ICollection<string> failures)
    {
        if (contract.SchemaVersion != "1.0" || contract.ContractKind != "NetOffice.WrapperContract" || contract.Source.Api != api)
            failures.Add(api + ": invalid wrapper contract identity.");
        if (classification.SchemaVersion != "1.0" || classification.ContractKind != "NetOffice.WrapperClassification" || classification.Api != api)
            failures.Add(api + ": invalid classification identity.");
        if (ledger.SchemaVersion != "1.0" || ledger.ContractKind != "NetOffice.WrapperCompatibilityLedger" || ledger.Api != api)
            failures.Add(api + ": invalid compatibility ledger identity.");
        if (!IsSorted(contract.Source.Files.Select(file => file.Path)))
            failures.Add(api + ": source files are not ordinally sorted.");
        if (!IsSorted(contract.Types.Select(type => type.LogicalId + "\0" + type.Source)))
            failures.Add(api + ": types are not ordinally sorted.");
        if (!IsSorted(classification.Files.Select(file => file.Path)))
            failures.Add(api + ": classified files are not ordinally sorted.");
        if (contract.Types.GroupBy(type => type.LogicalId, StringComparer.Ordinal).Any(group => string.IsNullOrWhiteSpace(group.Key) || group.Count() != 1))
            failures.Add(api + ": type logical IDs must be unique and non-empty.");
        if (contract.Unknowns.Count != 0)
            failures.Add(api + ": projectable contract contains unknown records.");
        if (contract.Ambiguities.Count != 0)
            failures.Add(api + ": projectable contract contains ambiguity records.");

        var contractFiles = contract.Source.Files.Select(file => ContractIdentity.NormalizePartition(file.Path)).ToHashSet(StringComparer.Ordinal);
        var classifiedFiles = classification.Files.Select(file => file.Path).ToHashSet(StringComparer.Ordinal);
        if (!contractFiles.SetEquals(classifiedFiles))
            failures.Add(api + ": classification does not cover every and only contract source file.");
        foreach (var type in contract.Types)
        {
            if (type.Parts.Count == 0 || !IsSorted(type.Parts.Select(part => part.Source + "\0" + part.Line.ToString("D9", System.Globalization.CultureInfo.InvariantCulture))))
                failures.Add(api + ": type parts are missing or unsorted for " + type.LogicalId + ".");
            if (type.Parts.Any(part => string.IsNullOrWhiteSpace(part.Source) || part.Line < 1 || string.IsNullOrWhiteSpace(part.Signature)))
                failures.Add(api + ": type part is incomplete for " + type.LogicalId + ".");
            if (!IsSorted(type.Members.Select(member => member.LogicalId + "\0" + member.Source + "\0" + member.Line.ToString("D9", System.Globalization.CultureInfo.InvariantCulture))))
                failures.Add(api + ": members are not ordinally sorted for " + type.LogicalId + ".");
            foreach (var member in type.Members)
            {
                if (member.LogicalId != ContractIdentity.MemberLogicalId(type.LogicalId, member))
                    failures.Add(api + ": invalid member logical ID " + member.LogicalId + ".");
                if (member.Documentation != null)
                {
                    foreach (var parameterName in member.Documentation.Parameters.Keys)
                    {
                        if (!HasDeclaredParameter(member.Parameters, parameterName))
                            failures.Add(api + ": documentation parameter " + parameterName + " is absent from " + member.LogicalId + ".");
                    }
                }
            }
        }
        var ledgerIds = ledger.Entries.Where(entry => entry.Status == "extracted").Select(entry => entry.LogicalId).ToHashSet(StringComparer.Ordinal);
        var contractIds = contract.Types.Select(type => type.LogicalId)
            .Concat(contract.Types.SelectMany(type => type.Members.Select(member => member.LogicalId))).ToHashSet(StringComparer.Ordinal);
        if (!ledgerIds.SetEquals(contractIds))
            failures.Add(api + ": compatibility ledger does not cover every extracted logical ID.");
    }

    private static bool HasDeclaredParameter(string parameters, string parameterName)
    {
        if (string.IsNullOrWhiteSpace(parameters) || string.IsNullOrWhiteSpace(parameterName))
            return false;
        return Regex.IsMatch(parameters, @"\b" + Regex.Escape(parameterName) + @"\b\s*(?==|,|\)|\])");
    }

    private static bool IsSorted(IEnumerable<string> values)
    {
        string previous = null;
        foreach (var value in values)
        {
            if (previous != null && StringComparer.Ordinal.Compare(previous, value) > 0)
                return false;
            previous = value;
        }
        return true;
    }
}

internal static class SemanticComparer
{
    public static SemanticDiffReport Compare(string expectedDirectory, string actualTree, IReadOnlyList<string> apis)
    {
        var report = new SemanticDiffReport();
        foreach (var api in apis)
        {
            var contractPath = Path.Combine(expectedDirectory, api + ".wrapper-contract.json");
            var classificationPath = Path.Combine(expectedDirectory, api + ".classification.json");
            if (!File.Exists(contractPath) || !File.Exists(classificationPath))
                throw new CommandLineException("Expected contract/classification is missing for " + api + ".");
            var expected = JsonFile.Read<WrapperContract>(contractPath);
            var classification = JsonFile.Read<ClassificationReport>(classificationPath);
            var ledgerPath = Path.Combine(expectedDirectory, api + ".compatibility-ledger.json");
            var ledger = File.Exists(ledgerPath) ? JsonFile.Read<CompatibilityLedger>(ledgerPath) : new CompatibilityLedger { Api = api };
            var actualRoot = FindActualApiRoot(actualTree, api, apis.Count == 1);
            var actual = Extractor.Extract(actualTree, api, actualRoot);
            report.Products.Add(CompareProduct(api, expected, actual, classification, ledger));
        }
        report.Products = report.Products.OrderBy(product => product.Api, StringComparer.Ordinal).ToList();
        report.Summary.Products = report.Products.Count;
        report.Summary.Differences = report.Products.Sum(product => product.Summary.Differences);
        report.Summary.IntentionalDifferences = report.Products.Sum(product => product.Summary.IntentionalDifferences);
        report.Summary.UnexplainedDifferences = report.Products.Sum(product => product.Summary.UnexplainedDifferences);
        report.Summary.ByDifferenceKind = SumCounts(report.Products.Select(product => product.Summary.ByDifferenceKind));
        report.Summary.ByRecordKind = SumCounts(report.Products.Select(product => product.Summary.ByRecordKind));
        report.Summary.ByFacet = SumCounts(report.Products.Select(product => product.Summary.ByFacet));
        return report;
    }

    private static ProductSemanticDiff CompareProduct(string api, WrapperContract expected, WrapperContract actual, ClassificationReport classification, CompatibilityLedger ledger)
    {
        var result = new ProductSemanticDiff { Api = api };
        result.Summary.Products = 1;
        var generatedFiles = classification.Files.Where(file => file.Ownership == "wrapper-generated")
            .Select(file => file.Path).ToHashSet(StringComparer.Ordinal);
        var expectedTypes = expected.Types.Where(type => generatedFiles.Contains(ContractIdentity.NormalizePartition(type.Source))).ToList();
        var actualTypes = actual.Types.Where(type => OwnershipClassifier.Classify(
            new SourceFile { Path = type.Source }, Array.Empty<TypeRecord>()).Ownership == "wrapper-generated").ToList();
        var expectedGroups = expectedTypes.GroupBy(type => type.LogicalId, StringComparer.Ordinal).ToDictionary(group => group.Key, group => group.ToList(), StringComparer.Ordinal);
        var actualGroups = actualTypes.GroupBy(type => type.LogicalId, StringComparer.Ordinal).ToDictionary(group => group.Key, group => group.ToList(), StringComparer.Ordinal);

        foreach (var id in expectedGroups.Keys.Union(actualGroups.Keys, StringComparer.Ordinal).OrderBy(value => value, StringComparer.Ordinal))
        {
            if (!expectedGroups.TryGetValue(id, out var expectedGroup))
            {
                Add(result, ledger, id, "type", "type", "extra", null, Canonical(actualGroups[id]));
                continue;
            }
            if (!actualGroups.TryGetValue(id, out var actualGroup))
            {
                Add(result, ledger, id, "type", "type", "missing", Canonical(expectedGroup), null);
                continue;
            }
            CompareTypeGroup(result, ledger, id, expectedGroup, actualGroup);
        }

        result.Differences = result.Differences.OrderBy(difference => difference.LogicalId, StringComparer.Ordinal)
            .ThenBy(difference => difference.RecordKind, StringComparer.Ordinal).ThenBy(difference => difference.Facet, StringComparer.Ordinal).ToList();
        result.FirstDivergentLogicalIds = result.Differences.Where(difference => !difference.Intentional)
            .Select(difference => difference.LogicalId).Distinct(StringComparer.Ordinal).ToList();
        result.Summary.Differences = result.Differences.Count;
        result.Summary.IntentionalDifferences = result.Differences.Count(difference => difference.Intentional);
        result.Summary.UnexplainedDifferences = result.Differences.Count(difference => !difference.Intentional);
        result.Summary.ByDifferenceKind = CountBy(result.Differences.Select(difference => difference.DifferenceKind));
        result.Summary.ByRecordKind = CountBy(result.Differences.Select(difference => difference.RecordKind));
        result.Summary.ByFacet = CountBy(result.Differences.Select(difference => difference.Facet));
        return result;
    }

    private static void CompareTypeGroup(ProductSemanticDiff result, CompatibilityLedger ledger, string id, List<TypeRecord> expected, List<TypeRecord> actual)
    {
        CompareFacet(result, ledger, id, "type", "kind", expected.Select(type => type.Kind), actual.Select(type => type.Kind));
        CompareFacet(result, ledger, id, "type", "accessibility", expected.Select(type => type.Accessibility), actual.Select(type => type.Accessibility));
        CompareFacet(result, ledger, id, "type", "modifiers", expected.SelectMany(type => type.Modifiers), actual.SelectMany(type => type.Modifiers));
        CompareFacet(result, ledger, id, "type", "attributes", expected.SelectMany(type => type.Attributes).Select(NormalizeAttribute), actual.SelectMany(type => type.Attributes).Select(NormalizeAttribute));
        CompareFacet(result, ledger, id, "type", "supportVersions", expected.SelectMany(type => type.SupportVersions).Select(NormalizeCSharp), actual.SelectMany(type => type.SupportVersions).Select(NormalizeCSharp));
        CompareFacet(result, ledger, id, "type", "basesInterfaces", expected.SelectMany(type => new[] { type.BaseType }.Concat(type.Interfaces)).Select(NormalizeCSharp), actual.SelectMany(type => new[] { type.BaseType }.Concat(type.Interfaces)).Select(NormalizeCSharp));
        CompareFacet(result, ledger, id, "type", "docs", expected.SelectMany(type => type.Parts).Select(part => CanonicalDocumentation(part.Documentation)), actual.SelectMany(type => type.Parts).Select(part => CanonicalDocumentation(part.Documentation)));
        CompareFacet(result, ledger, id, "type", "partition", expected.SelectMany(type => type.Parts).Select(part => ContractIdentity.NormalizePartition(part.Source)), actual.SelectMany(type => type.Parts).Select(part => ContractIdentity.NormalizePartition(part.Source)));

        var expectedMembers = expected.SelectMany(type => type.Members).GroupBy(member => ComparisonMemberId(id, member), StringComparer.Ordinal).ToDictionary(group => group.Key, group => group.ToList(), StringComparer.Ordinal);
        var actualMembers = actual.SelectMany(type => type.Members).GroupBy(member => ComparisonMemberId(id, member), StringComparer.Ordinal).ToDictionary(group => group.Key, group => group.ToList(), StringComparer.Ordinal);
        foreach (var comparisonId in expectedMembers.Keys.Union(actualMembers.Keys, StringComparer.Ordinal).OrderBy(value => value, StringComparer.Ordinal))
        {
            if (!expectedMembers.TryGetValue(comparisonId, out var expectedGroup))
            {
                var extraGroup = actualMembers[comparisonId];
                Add(result, ledger, extraGroup[0].LogicalId, "member", "member", "extra", null, Canonical(extraGroup));
                continue;
            }
            if (!actualMembers.TryGetValue(comparisonId, out var actualGroup))
            {
                Add(result, ledger, expectedGroup[0].LogicalId, "member", "member", "missing", Canonical(expectedGroup), null);
                continue;
            }
            var memberId = expectedGroup[0].LogicalId;
            CompareFacet(result, ledger, memberId, "member", "signature", expectedGroup.Select(NormalizeMemberSignature), actualGroup.Select(NormalizeMemberSignature));
            CompareFacet(result, ledger, memberId, "member", "attributes", expectedGroup.SelectMany(member => member.Attributes).Select(NormalizeAttribute), actualGroup.SelectMany(member => member.Attributes).Select(NormalizeAttribute));
            CompareFacet(result, ledger, memberId, "member", "defaults", expectedGroup.Select(CanonicalDefaultValues), actualGroup.Select(CanonicalDefaultValues));
            CompareFacet(result, ledger, memberId, "member", "invocation", expectedGroup.SelectMany(member => member.InvocationText).Select(NormalizeInvocation), actualGroup.SelectMany(member => member.InvocationText).Select(NormalizeInvocation));
            CompareFacet(result, ledger, memberId, "member", "supportVersions", expectedGroup.SelectMany(member => member.SupportVersions).Select(NormalizeCSharp), actualGroup.SelectMany(member => member.SupportVersions).Select(NormalizeCSharp));
            CompareFacet(result, ledger, memberId, "member", "docs", expectedGroup.Select(member => CanonicalDocumentation(member.Documentation)), actualGroup.Select(member => CanonicalDocumentation(member.Documentation)));
            CompareFacet(result, ledger, memberId, "member", "partition", expectedGroup.Select(member => ContractIdentity.NormalizePartition(member.Source)), actualGroup.Select(member => ContractIdentity.NormalizePartition(member.Source)));
        }
    }

    private static void CompareFacet(ProductSemanticDiff result, CompatibilityLedger ledger, string id, string kind, string facet, IEnumerable<string> expected, IEnumerable<string> actual)
    {
        var expectedValue = CanonicalValues(expected);
        var actualValue = CanonicalValues(actual);
        if (!string.Equals(expectedValue, actualValue, StringComparison.Ordinal))
            Add(result, ledger, id, kind, facet, "changed", expectedValue, actualValue);
    }

    private static void Add(ProductSemanticDiff result, CompatibilityLedger ledger, string id, string kind, string facet, string differenceKind, string expected, string actual)
    {
        var intentional = ledger.Entries.FirstOrDefault(entry => entry.Status == "intentional" && entry.LogicalId == id && (entry.Facet == null || entry.Facet == facet));
        result.Differences.Add(new SemanticDifference
        {
            LogicalId = id,
            RecordKind = kind,
            Facet = facet,
            DifferenceKind = differenceKind,
            Expected = expected,
            Actual = actual,
            Intentional = intentional != null,
            LedgerRationale = intentional?.Rationale,
            LedgerProvenance = intentional?.Provenance
        });
    }

    private static string FindActualApiRoot(string actualTree, string api, bool onlyApi)
    {
        var candidates = new[]
        {
            Path.Combine(actualTree, api), Path.Combine(actualTree, "Source", api),
            Path.Combine(actualTree, api, "Generated"), Path.Combine(actualTree, "Source", api, "Generated")
        };
        foreach (var candidate in candidates)
            if (Directory.Exists(candidate) && Directory.EnumerateFiles(candidate, "*.cs", SearchOption.AllDirectories).Any())
                return candidate;
        if (onlyApi && Directory.EnumerateFiles(actualTree, "*.cs", SearchOption.AllDirectories).Any())
            return actualTree;
        throw new CommandLineException("Generated output tree has no C# source root for " + api + ".");
    }

    private static string ComparisonMemberId(string typeId, MemberRecord member) =>
        typeId + "::" + member.Kind + ":" + member.Name + NormalizeParameterIdentity(member.Parameters ?? string.Empty);

    private static string NormalizeParameterIdentity(string parameters)
    {
        var normalized = NormalizeCSharp(parameters);
        string previous;
        do
        {
            previous = normalized;
            normalized = Regex.Replace(normalized, @"(?<prefix>[(,\[])\s*\[[^\]]*\]\s*", "${prefix}");
        } while (!string.Equals(previous, normalized, StringComparison.Ordinal));
        return normalized;
    }

    private static string NormalizeMemberSignature(MemberRecord member)
    {
        var normalized = NormalizeCSharp(member.Signature);
        if (string.Equals(member.Kind, "event", StringComparison.Ordinal) &&
            normalized.EndsWith(" {", StringComparison.Ordinal))
            return normalized[..^2];
        if (string.Equals(member.Kind, "enumValue", StringComparison.Ordinal) &&
            normalized.EndsWith(",", StringComparison.Ordinal))
            return normalized[..^1].TrimEnd();
        return normalized;
    }

    private static string NormalizeInvocation(string value)
    {
        var normalized = NormalizeCSharp(value);
        if (normalized == null)
            return null;
        var tokens = new List<string>();
        for (var index = 0; index < normalized.Length;)
        {
            if (char.IsWhiteSpace(normalized[index]))
            {
                index++;
                continue;
            }
            if (index + 1 < normalized.Length && normalized[index] == '/' && normalized[index + 1] == '/')
                break;
            if (index + 1 < normalized.Length && normalized[index] == '/' && normalized[index + 1] == '*')
            {
                var commentEnd = normalized.IndexOf("*/", index + 2, StringComparison.Ordinal);
                index = commentEnd < 0 ? normalized.Length : commentEnd + 2;
                continue;
            }

            var start = index;
            if (normalized[index] is '\"' or '\'')
            {
                var quote = normalized[index++];
                var verbatim = quote == '\"' && start > 0 && normalized[start - 1] == '@';
                while (index < normalized.Length)
                {
                    if (normalized[index] != quote)
                    {
                        index++;
                        continue;
                    }
                    if (verbatim && index + 1 < normalized.Length && normalized[index + 1] == quote)
                    {
                        index += 2;
                        continue;
                    }
                    var slashes = 0;
                    for (var scan = index - 1; !verbatim && scan >= start && normalized[scan] == '\\'; scan--)
                        slashes++;
                    index++;
                    if (verbatim || (slashes & 1) == 0)
                        break;
                }
                tokens.Add(normalized[start..index]);
                continue;
            }
            if (char.IsLetter(normalized[index]) || normalized[index] == '_')
            {
                index++;
                while (index < normalized.Length &&
                       (char.IsLetterOrDigit(normalized[index]) || normalized[index] == '_'))
                    index++;
                tokens.Add(normalized[start..index]);
                continue;
            }
            if (char.IsDigit(normalized[index]))
            {
                index++;
                while (index < normalized.Length &&
                       (char.IsLetterOrDigit(normalized[index]) || normalized[index] is '_' or '.'))
                    index++;
                tokens.Add(normalized[start..index]);
                continue;
            }

            var operatorLength = CSharpOperatorLength(normalized, index);
            tokens.Add(normalized.Substring(index, operatorLength));
            index += operatorLength;
        }
        return string.Join(" ", RemoveRedundantEmptyStatements(tokens));
    }

    private static List<string> RemoveRedundantEmptyStatements(List<string> tokens)
    {
        var result = new List<string>(tokens.Count);
        var parenthesisDepth = 0;
        foreach (var token in tokens)
        {
            if (token == "(")
                parenthesisDepth++;
            else if (token == ")" && parenthesisDepth > 0)
                parenthesisDepth--;
            if (token == ";" && parenthesisDepth == 0 && result.Count > 0 && result[^1] == ";")
                continue;
            result.Add(token);
        }
        return result;
    }

    private static int CSharpOperatorLength(string value, int index)
    {
        if (index + 3 <= value.Length)
        {
            var three = value.Substring(index, 3);
            if (three is "<<=" or ">>=" or "??=")
                return 3;
        }
        if (index + 2 <= value.Length)
        {
            var two = value.Substring(index, 2);
            if (two is "=>" or "==" or "!=" or "<=" or ">=" or "++" or "--" or "&&" or "||" or
                "??" or "?." or "+=" or "-=" or "*=" or "/=" or "%=" or "&=" or "|=" or "^=" or
                "<<" or ">>" or "::" or "->")
                return 2;
        }
        return 1;
    }

    private static string CanonicalDefaultValues(MemberRecord member)
    {
        var normalized = member.DefaultValues.OrderBy(pair => pair.Key, StringComparer.Ordinal)
            .ToDictionary(pair => pair.Key, pair => NormalizeCSharp(pair.Value), StringComparer.Ordinal);
        return Canonical(normalized);
    }

    private static string NormalizeAttribute(string value) =>
        NormalizeAttributeNamespaceSegment(NormalizeCSharp(value));

    private static string NormalizeAttributeNamespaceSegment(string value)
    {
        var normalized = Regex.Replace(value,
            @"\b(?:System\.ComponentModel|System\.Runtime\.CompilerServices|System\.Runtime\.InteropServices|NetOffice\.Attributes)\.",
            string.Empty);
        normalized = Regex.Replace(normalized, @"\b([A-Za-z_]\w*)Attribute(?=\s*(?:\(|$))", "$1");
        return normalized;
    }

    private static string NormalizeCSharp(string value)
    {
        if (value == null)
            return null;
        var result = new System.Text.StringBuilder(value.Length);
        var segmentStart = 0;
        for (var index = 0; index < value.Length; index++)
        {
            var quote = value[index];
            if (quote != '\"' && quote != '\'')
                continue;
            result.Append(NormalizeCSharpSegment(value[segmentStart..index]));
            var literalStart = index;
            var verbatim = quote == '\"' && index > 0 && value[index - 1] == '@';
            index++;
            while (index < value.Length)
            {
                if (value[index] == quote)
                {
                    if (verbatim && index + 1 < value.Length && value[index + 1] == quote)
                    {
                        index += 2;
                        continue;
                    }
                    if (!verbatim && index > literalStart && value[index - 1] == '\\')
                    {
                        var slashes = 1;
                        for (var scan = index - 2; scan >= literalStart && value[scan] == '\\'; scan--)
                            slashes++;
                        if ((slashes & 1) == 1)
                        {
                            index++;
                            continue;
                        }
                    }
                    break;
                }
                index++;
            }
            var literalEnd = Math.Min(index + 1, value.Length);
            result.Append(value, literalStart, literalEnd - literalStart);
            segmentStart = literalEnd;
            index = literalEnd - 1;
        }
        result.Append(NormalizeCSharpSegment(value[segmentStart..]));
        var normalizedResult = result.ToString().Trim();
        return Regex.Replace(normalizedResult, @"\[(?<attribute>[^\[\]]+)\]",
            match => "[" + NormalizeAttributeNamespaceSegment(match.Groups["attribute"].Value) + "]");
    }

    private static string NormalizeCSharpSegment(string value)
    {
        var normalized = value.Replace("global::", string.Empty, StringComparison.Ordinal);
        normalized = Regex.Replace(normalized, @"\bNetRuntimeSystem\b", "System");
        var aliases = new Dictionary<string, string>(StringComparer.Ordinal)
        {
            ["Boolean"] = "bool", ["Byte"] = "byte", ["SByte"] = "sbyte",
            ["Int16"] = "short", ["UInt16"] = "ushort", ["Int32"] = "int",
            ["UInt32"] = "uint", ["Int64"] = "long", ["UInt64"] = "ulong",
            ["Single"] = "float", ["Double"] = "double", ["Decimal"] = "decimal",
            ["Char"] = "char", ["String"] = "string", ["Object"] = "object",
            ["Void"] = "void"
        };
        foreach (var alias in aliases)
            normalized = Regex.Replace(normalized, @"(?<![\w.])(?:System\.)?" + alias.Key + @"\b", alias.Value);
        normalized = Regex.Replace(normalized, @"(?<![\w.])(?:System\.)?Type\b", "System.Type");
        normalized = Regex.Replace(normalized, @"\s*,\s*", ", ");
        normalized = Regex.Replace(normalized, @"\s*\[\s*\]", "[]");
        normalized = Regex.Replace(normalized, @"\s*\?\s*", "?");
        normalized = Regex.Replace(normalized, @"\s*:\s*", " : ");
        normalized = Regex.Replace(normalized, @"(?<![=!<>])\s*=\s*(?![=>])", " = ");
        normalized = Regex.Replace(normalized, @"\(\s*", "(");
        normalized = Regex.Replace(normalized, @"\s*\)", ")");
        normalized = Regex.Replace(normalized, @"\s+", " ");
        return normalized;
    }

    private static string CanonicalDocumentation(DocumentationRecord documentation)
    {
        if (documentation == null)
            return "null";
        var normalized = new DocumentationRecord
        {
            ParseStatus = documentation.ParseStatus,
            ParseError = documentation.ParseError,
            Raw = NormalizeXmlDocumentation(documentation.Raw),
            Summary = documentation.Summary,
            Remarks = documentation.Remarks,
            Returns = documentation.Returns,
            Value = documentation.Value,
            Parameters = documentation.Parameters.OrderBy(pair => pair.Key, StringComparer.Ordinal)
                .ToDictionary(pair => pair.Key, pair => pair.Value, StringComparer.Ordinal),
            TypeParameters = documentation.TypeParameters.OrderBy(pair => pair.Key, StringComparer.Ordinal)
                .ToDictionary(pair => pair.Key, pair => pair.Value, StringComparer.Ordinal),
            Exceptions = documentation.Exceptions.OrderBy(value => value, StringComparer.Ordinal).ToList()
        };
        return Canonical(normalized);
    }

    private static string NormalizeXmlDocumentation(string raw)
    {
        if (raw == null)
            return null;
        var result = new System.Text.StringBuilder(raw.Length);
        var index = 0;
        while (index < raw.Length)
        {
            if (raw[index] != '<')
            {
                var nextTag = raw.IndexOf('<', index);
                if (nextTag < 0)
                    nextTag = raw.Length;
                var text = Regex.Replace(raw[index..nextTag], @"\s+", " ").Trim();
                if (text.Length > 0)
                    result.Append(text);
                index = nextTag;
                continue;
            }

            var end = FindMarkupEnd(raw, index);
            if (end < 0)
            {
                result.Append(Regex.Replace(raw[index..], @"\s+", " ").Trim());
                break;
            }
            result.Append(NormalizeMarkupTag(raw[index..(end + 1)]));
            index = end + 1;
        }
        return result.ToString();
    }

    private static int FindMarkupEnd(string value, int start)
    {
        var quote = '\0';
        for (var index = start + 1; index < value.Length; index++)
        {
            var current = value[index];
            if (quote != '\0')
            {
                if (current == quote)
                    quote = '\0';
                continue;
            }
            if (current is '\"' or '\'')
                quote = current;
            else if (current == '>')
                return index;
        }
        return -1;
    }

    private static string NormalizeMarkupTag(string tag)
    {
        var result = new System.Text.StringBuilder(tag.Length);
        var quote = '\0';
        var pendingSpace = false;
        foreach (var current in tag)
        {
            if (quote != '\0')
            {
                result.Append(current);
                if (current == quote)
                    quote = '\0';
                continue;
            }
            if (current is '\"' or '\'')
            {
                if (pendingSpace && result.Length > 0 && result[^1] is not '<' and not '=')
                    result.Append(' ');
                pendingSpace = false;
                quote = current;
                result.Append(current);
                continue;
            }
            if (char.IsWhiteSpace(current))
            {
                pendingSpace = true;
                continue;
            }
            if (pendingSpace && result.Length > 0 && result[^1] is not '<' and not '=' and not '/' &&
                current is not '>' and not '=' and not '/')
                result.Append(' ');
            pendingSpace = false;
            result.Append(current);
        }
        return result.ToString();
    }

    private static Dictionary<string, int> CountBy(IEnumerable<string> values) => values
        .GroupBy(value => value, StringComparer.Ordinal).OrderBy(group => group.Key, StringComparer.Ordinal)
        .ToDictionary(group => group.Key, group => group.Count(), StringComparer.Ordinal);

    private static Dictionary<string, int> SumCounts(IEnumerable<Dictionary<string, int>> counts) => counts
        .SelectMany(count => count)
        .GroupBy(pair => pair.Key, StringComparer.Ordinal).OrderBy(group => group.Key, StringComparer.Ordinal)
        .ToDictionary(group => group.Key, group => group.Sum(pair => pair.Value), StringComparer.Ordinal);

    private static string CanonicalValues(IEnumerable<string> values) => string.Join("\n", values.Where(value => value != null).Select(value => value).Distinct(StringComparer.Ordinal).OrderBy(value => value, StringComparer.Ordinal));
    private static string Canonical(object value) => value == null ? "null" : JsonSerializer.Serialize(value, JsonFile.SerializerOptions);
}

internal sealed class SemanticDiffReport
{
    public string SchemaVersion { get; set; } = "1.0";
    public string ContractKind { get; set; } = "NetOffice.WrapperSemanticDiff";
    public string NormalizationProfile { get; set; } = "csharp-semantic/v1";
    public SemanticDiffSummary Summary { get; set; } = new();
    public List<ProductSemanticDiff> Products { get; set; } = new();
}

internal sealed class ProductSemanticDiff
{
    public string SchemaVersion { get; set; } = "1.0";
    public string ContractKind { get; set; } = "NetOffice.WrapperProductSemanticDiff";
    public string NormalizationProfile { get; set; } = "csharp-semantic/v1";
    public string Api { get; set; } = string.Empty;
    public SemanticDiffSummary Summary { get; set; } = new();
    public List<string> FirstDivergentLogicalIds { get; set; } = new();
    public List<SemanticDifference> Differences { get; set; } = new();
}

internal sealed class SemanticDiffSummary
{
    public int Products { get; set; }
    public int Differences { get; set; }
    public int IntentionalDifferences { get; set; }
    public int UnexplainedDifferences { get; set; }
    public Dictionary<string, int> ByDifferenceKind { get; set; } = new(StringComparer.Ordinal);
    public Dictionary<string, int> ByRecordKind { get; set; } = new(StringComparer.Ordinal);
    public Dictionary<string, int> ByFacet { get; set; } = new(StringComparer.Ordinal);
}

internal sealed class SemanticDifference
{
    public string LogicalId { get; set; } = string.Empty;
    public string RecordKind { get; set; } = string.Empty;
    public string Facet { get; set; } = string.Empty;
    public string DifferenceKind { get; set; } = string.Empty;
    public string Expected { get; set; }
    public string Actual { get; set; }
    public bool Intentional { get; set; }
    public string LedgerRationale { get; set; }
    public string LedgerProvenance { get; set; }
}
