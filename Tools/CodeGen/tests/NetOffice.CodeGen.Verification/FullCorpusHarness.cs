using System.Diagnostics;
using System.Security.Cryptography;
using System.Text;
using System.Text.Json;
using System.Text.RegularExpressions;
using NetOffice.CodeGen.Data;
using NetOffice.CodeGen.Emit;
using NetOffice.CodeGen.Projection;
using NetOffice.CodeGen.Storage;

namespace NetOffice.CodeGen.Verification;

internal sealed record FullCorpusOptions(
    bool Enabled,
    string? DataPath,
    string SourcePath,
    string ContractsPath,
    string? PolicyPath,
    string? ReportPath)
{
    public static FullCorpusOptions Parse(string[] args)
    {
        var repositoryRoot = FindRepositoryRoot();
        var enabled = args.Contains("--full-corpus", StringComparer.Ordinal);
        var data = Value(args, "--full-data") ?? Environment.GetEnvironmentVariable("NETOFFICE_DATA_V2");
        var source = Value(args, "--full-source") ?? Path.Combine(repositoryRoot, "Source");
        var contracts = Value(args, "--full-contracts") ?? Path.Combine(repositoryRoot, "Tools", "CodeGen", "contracts", "wrapper");
        return new FullCorpusOptions(enabled, data, source, contracts, Value(args, "--full-policy"), Value(args, "--full-report"));
    }

    private static string? Value(IReadOnlyList<string> args, string name)
    {
        for (var index = 0; index < args.Count; index++)
            if (string.Equals(args[index], name, StringComparison.Ordinal))
                return index + 1 < args.Count ? args[index + 1] : throw new ArgumentException(name + " requires a value.");
        return null;
    }

    private static string FindRepositoryRoot()
    {
        for (var candidate = new DirectoryInfo(AppContext.BaseDirectory); candidate is not null; candidate = candidate.Parent)
            if (Directory.Exists(Path.Combine(candidate.FullName, "Tools", "CodeGen")) && Directory.Exists(Path.Combine(candidate.FullName, "Source")))
                return candidate.FullName;
        var current = Path.GetFullPath(Environment.CurrentDirectory);
        if (Directory.Exists(Path.Combine(current, "NetOffice", "Tools", "CodeGen"))) return Path.Combine(current, "NetOffice");
        throw new DirectoryNotFoundException("Unable to locate the NetOffice repository root.");
    }
}

internal static class FullCorpusHarness
{
    private const int ExpectedTypes = 4368;
    private const int ExpectedMembers = 40590;
    private const int ExpectedValues = 15409;
    private const int ExpectedWrapperProducts = 12;
    private static readonly UTF8Encoding Utf8NoBom = new(false);
    private static readonly Regex BuildDiagnosticPattern = new(
        @"^(?:(?<file>.+?)(?:\((?<line>\d+),(?<column>\d+)\))?\s*:\s*)?(?<severity>error|warning)\s+(?<code>[A-Za-z]+\d+):\s*(?<message>.*?)(?:\s+\[(?<project>[^\]]+)\])?$",
        RegexOptions.Compiled | RegexOptions.CultureInvariant | RegexOptions.IgnoreCase);

    public static FullCorpusReport Run(FullCorpusOptions options)
    {
        var report = new FullCorpusReport
        {
            DataPath = options.DataPath ?? string.Empty,
            SourcePath = Path.GetFullPath(options.SourcePath),
            ContractsPath = Path.GetFullPath(options.ContractsPath)
        };
        var workRoot = Path.Combine(Path.GetTempPath(), "netoffice-codegen-full-verification-" + Guid.NewGuid().ToString("N"));
        Directory.CreateDirectory(workRoot);
        try
        {
            Execute(options, workRoot, report);
        }
        catch (Exception error)
        {
            report.Diagnostics.Add(new FullCorpusDiagnostic("verification", "harness.unhandled", "", "", null, null, error.ToString()));
        }
        finally
        {
            try { if (Directory.Exists(workRoot)) Directory.Delete(workRoot, true); }
            catch (Exception error) { report.Diagnostics.Add(new FullCorpusDiagnostic("verification", "harness.cleanup", workRoot, "", null, null, error.Message)); }
            FinalizeReport(report);
        }
        return report;
    }

    private static void Execute(FullCorpusOptions options, string workRoot, FullCorpusReport report)
    {
        if (string.IsNullOrWhiteSpace(options.DataPath))
        {
            report.Diagnostics.Add(new FullCorpusDiagnostic("data-v2", "input.missing", "", "", null, null, "--full-data or NETOFFICE_DATA_V2 is required for --full-corpus."));
            return;
        }
        var graphPath = File.Exists(options.DataPath) ? Path.GetFullPath(options.DataPath) : Path.Combine(Path.GetFullPath(options.DataPath), "graph.json");
        if (!File.Exists(graphPath))
        {
            report.Diagnostics.Add(new FullCorpusDiagnostic("data-v2", "input.missing", graphPath, "", null, null, "The canonical full graph does not exist."));
            return;
        }

        DataGraph graph;
        try
        {
            graph = CanonicalJson.Parse(File.ReadAllText(graphPath));
            var validation = DataGraphValidator.Validate(graph, graph.Digest);
            foreach (var issue in validation.Issues)
                report.Diagnostics.Add(new FullCorpusDiagnostic("data-v2", issue.Code, issue.Path ?? graphPath, "", null, null, issue.Message));
            report.GraphDigest = graph.Digest;
            report.Coverage = new FullCorpusCoverage(graph.Types.Count, 0, graph.Members.Count, 0, graph.Values.Count, 0);
            report.InputProducts = graph.Projects.Select(static item => item.Name).Distinct(StringComparer.Ordinal).Order(StringComparer.Ordinal).ToArray();
            report.Gates["converted-graph-valid"] = validation.IsValid;
            report.Gates["canonical-corpus-size"] = graph.Types.Count == ExpectedTypes && graph.Members.Count == ExpectedMembers && graph.Values.Count == ExpectedValues;
            var sourceProducts = DiscoverWrapperProjects(report.SourcePath).Select(static item => item.Product).ToArray();
            report.Gates["all-products"] = sourceProducts.Length == ExpectedWrapperProducts &&
                sourceProducts.All(product => report.InputProducts.Contains(product, StringComparer.OrdinalIgnoreCase));
            if (!report.Gates["all-products"])
                report.Diagnostics.Add(new FullCorpusDiagnostic("verification", "coverage.products", report.SourcePath, "", null, null,
                    $"Expected {ExpectedWrapperProducts} generated wrapper projects represented in the graph; found {sourceProducts.Length}."));
            if (!report.Gates["canonical-corpus-size"])
                report.Diagnostics.Add(new FullCorpusDiagnostic("data-v2", "coverage.canonical", graphPath, "", null, null,
                    $"Expected {ExpectedTypes}/{ExpectedMembers}/{ExpectedValues} types/members/values; found {graph.Types.Count}/{graph.Members.Count}/{graph.Values.Count}."));
        }
        catch (Exception error)
        {
            report.Diagnostics.Add(new FullCorpusDiagnostic("data-v2", "input.invalid", graphPath, "", null, null, error.Message));
            return;
        }

        var contractFiles = ContractFiles(options.ContractsPath);
        report.ContractFiles = contractFiles.Select(static path => Path.GetFileName(path) ?? path).Order(StringComparer.Ordinal).ToArray();
        var contracts = new List<(string Path, WrapperContract Contract)>();
        foreach (var path in contractFiles)
        {
            try { contracts.Add((path, WrapperContract.Read(path))); }
            catch (Exception error) { report.Diagnostics.Add(new FullCorpusDiagnostic("parity-contract", "contract.invalid", path, "", null, null, error.Message)); }
        }
        if (contracts.Count == 0)
        {
            report.Diagnostics.Add(new FullCorpusDiagnostic("parity-contract", "contract.missing", options.ContractsPath, "", null, null, "No wrapper contracts were found."));
            return;
        }

        var policyPath = options.PolicyPath;
        ProjectionPolicy policy;
        try
        {
            if (string.IsNullOrWhiteSpace(policyPath))
            {
                policy = ProjectionPolicy.Default;
                policyPath = Path.Combine(workRoot, "projection-policy.json");
                File.WriteAllText(policyPath, JsonSerializer.Serialize(policy, JsonOptions) + "\n", Utf8NoBom);
            }
            else
            {
                policyPath = Path.GetFullPath(policyPath);
                policy = ProjectionPolicy.Read(policyPath);
            }
            policy.ValidateOrThrow();
        }
        catch (Exception error)
        {
            report.Diagnostics.Add(new FullCorpusDiagnostic("projection", "policy.invalid", policyPath ?? "", "", null, null, error.Message));
            return;
        }

        ProjectionResult? projection = null;
        try
        {
            var primary = contracts.Count == 1 ? contracts[0].Contract : EmptyContract();
            projection = ProjectionEngine.Project(graph, primary, policy);
            report.Coverage = new FullCorpusCoverage(
                projection.Coverage.DataTypes, projection.Coverage.ProjectedTypes,
                projection.Coverage.DataMembers, projection.Coverage.ProjectedDataMembers,
                projection.Coverage.DataValues, projection.Coverage.ProjectedValues);
            var projectionProducts = DiscoverWrapperProjects(report.SourcePath).Select(static item => item.Product).ToHashSet(StringComparer.OrdinalIgnoreCase);
            report.ExpectedFiles = projection.Files.Count(file => file.Types.Any(type => projectionProducts.Contains(type.Product)));
            report.Gates["projection-coverage"] =
                projection.Coverage.DataTypes == ExpectedTypes && projection.Coverage.ProjectedTypes == ExpectedTypes &&
                projection.Coverage.DataMembers == ExpectedMembers && projection.Coverage.ProjectedDataMembers == ExpectedMembers &&
                projection.Coverage.DataValues == ExpectedValues && projection.Coverage.ProjectedValues == ExpectedValues;
            if (!report.Gates["projection-coverage"])
                report.Diagnostics.Add(new FullCorpusDiagnostic("projection", "coverage.incomplete", "", "", null, null,
                    $"Expected projected coverage {ExpectedTypes}/{ExpectedMembers}/{ExpectedValues}; found {projection.Coverage.ProjectedTypes}/{projection.Coverage.ProjectedDataMembers}/{projection.Coverage.ProjectedValues}."));
        }
        catch (ProjectionValidationException error)
        {
            foreach (var issue in error.Issues)
                report.Diagnostics.Add(new FullCorpusDiagnostic("projection", issue.Code, issue.Path, "", null, null, issue.Message));
        }
        catch (Exception error)
        {
            report.Diagnostics.Add(new FullCorpusDiagnostic("projection", "projection.failed", "", "", null, null, error.Message));
        }
        if (projection is null) return;
        var generatedProducts = DiscoverWrapperProjects(report.SourcePath).Select(static item => item.Product).Order(StringComparer.Ordinal).ToArray();

        var output = Path.Combine(workRoot, "generated-first");
        var repeatedOutput = Path.Combine(workRoot, "generated-second");
        var firstReport = Path.Combine(workRoot, "generate-first.json");
        var secondReport = Path.Combine(workRoot, "generate-second.json");
        var common = new[]
        {
            "--locked", "--no-cache", "--isolated-output", "--docs-profile", "baseline", "--data", graphPath,
            "--contract", Path.GetFullPath(options.ContractsPath), "--policy", policyPath!,
            "--source", report.SourcePath, "--projects", "all"
        };
        var first = RunCli("generate", common.Concat(new[] { "--output", output, "--report", firstReport }), TimeSpan.FromMinutes(20));
        var second = RunCli("generate", common.Concat(new[] { "--output", repeatedOutput, "--report", secondReport }), TimeSpan.FromMinutes(20));
        report.Generation = new FullCorpusGeneration(
            first.ExitCode, second.ExitCode, NormalizeCommand(common, output), firstReport, secondReport,
            first.Stdout, first.Stderr, second.Stdout, second.Stderr);
        report.Gates["generation-succeeded"] = first.ExitCode == 0 && second.ExitCode == 0;
        if (first.ExitCode != 0)
            report.Diagnostics.Add(new FullCorpusDiagnostic(GenerationOwner(first.Stderr, first.Stdout), "generate.failed", firstReport, "", null, null, FirstUsefulLine(first.Stderr, first.Stdout)));
        if (second.ExitCode != 0)
            report.Diagnostics.Add(new FullCorpusDiagnostic(GenerationOwner(second.Stderr, second.Stdout), "generate.repeat-failed", secondReport, "", null, null, FirstUsefulLine(second.Stderr, second.Stdout)));

        InspectGeneratedOutput(output, repeatedOutput, firstReport, secondReport, generatedProducts, report);
        VerifySemanticParity(output, options.ContractsPath, generatedProducts, workRoot, report);
        BuildGeneratedLibraries(output, report, workRoot);
    }

    private static void InspectGeneratedOutput(
        string output,
        string repeatedOutput,
        string firstReportPath,
        string secondReportPath,
        IReadOnlyCollection<string> generatedProducts,
        FullCorpusReport report)
    {
        CodegenGenerationReport first;
        CodegenGenerationReport second;
        try
        {
            first = ReadCodegenReport(firstReportPath);
            second = ReadCodegenReport(secondReportPath);
        }
        catch (Exception error)
        {
            report.Gates["complete-output"] = false;
            report.Gates["csharp73-bom-crlf-d9"] = false;
            report.Gates["no-source-wrapper-copy"] = false;
            report.Gates["deterministic-output"] = false;
            report.Gates["complete-manifest"] = false;
            report.Diagnostics.Add(new FullCorpusDiagnostic("orchestration", "output.report-invalid", firstReportPath, "", null, null, error.Message));
            return;
        }

        var expectedProducts = generatedProducts.Order(StringComparer.Ordinal).ToArray();
        var emittedPaths = first.EmittedPaths.Order(StringComparer.Ordinal).ToArray();
        var repeatedPaths = second.EmittedPaths.Order(StringComparer.Ordinal).ToArray();
        var actualPaths = Directory.Exists(output)
            ? Directory.EnumerateFiles(output, "*.cs", SearchOption.AllDirectories)
                .Select(path => Path.GetRelativePath(output, path).Replace('\\', '/'))
                .Where(static path => path.Contains("/Generated/", StringComparison.Ordinal))
                .Order(StringComparer.Ordinal)
                .ToArray()
            : Array.Empty<string>();
        var missing = emittedPaths.Except(actualPaths, StringComparer.Ordinal).ToArray();
        var extra = actualPaths.Except(emittedPaths, StringComparer.Ordinal).ToArray();
        report.ExpectedFiles = first.ProjectedFiles;
        report.ActualFiles = actualPaths.Length;
        report.AdditionalFiles = extra.Length;
        foreach (var path in missing)
            report.Diagnostics.Add(new FullCorpusDiagnostic("orchestration", "output.missing", path, "", null, null, "An emitted path from codegen-report-v3 does not exist."));
        foreach (var path in extra)
            report.Diagnostics.Add(new FullCorpusDiagnostic("orchestration", "output.unreported", path, "", null, null, "A generated C# file is absent from codegen-report-v3 emittedPaths."));

        var reportShapeValid =
            first.SchemaVersion == "codegen-report-v3" && second.SchemaVersion == "codegen-report-v3" &&
            first.IsolatedOutput && second.IsolatedOutput && first.ExitCode == 0 && second.ExitCode == 0 &&
            first.InputTypes == ExpectedTypes && first.InputMembers == ExpectedMembers && first.InputValues == ExpectedValues &&
            first.SelectedProducts == ExpectedWrapperProducts && first.ProjectedFiles == emittedPaths.Length &&
            first.Products.Order(StringComparer.Ordinal).SequenceEqual(expectedProducts, StringComparer.Ordinal) &&
            emittedPaths.SequenceEqual(repeatedPaths, StringComparer.Ordinal);
        report.Gates["complete-output"] = reportShapeValid && emittedPaths.Length > 77 && missing.Length == 0 && extra.Length == 0;
        if (!reportShapeValid)
            report.Diagnostics.Add(new FullCorpusDiagnostic("orchestration", "output.report-shape", firstReportPath, "", null, null,
                "codegen-report-v3 did not describe the complete isolated 12-product corpus."));

        var validator = new BasicSourceValidator();
        var invalid = 0;
        foreach (var relative in emittedPaths)
        {
            var path = ResolveOutputPath(output, relative);
            var valid = File.Exists(path);
            if (!valid) continue;
            var bytes = File.ReadAllBytes(path);
            var hasBom = bytes.Length >= 3 && bytes[0] == 0xef && bytes[1] == 0xbb && bytes[2] == 0xbf;
            if (!hasBom)
            {
                valid = false;
                report.Diagnostics.Add(new FullCorpusDiagnostic("emission", "output.bom", relative, "", null, null, "Generated C# does not start with the UTF-8 BOM."));
            }
            var crlf = HasOnlyCrLfWithFinalNewline(bytes, hasBom ? 3 : 0);
            if (!crlf)
            {
                valid = false;
                report.Diagnostics.Add(new FullCorpusDiagnostic("emission", "output.crlf", relative, "", null, null, "Generated C# is not CRLF-only with a final newline."));
            }
            string text;
            try
            {
                text = new UTF8Encoding(false, true).GetString(bytes, hasBom ? 3 : 0, bytes.Length - (hasBom ? 3 : 0));
            }
            catch (Exception error)
            {
                valid = false;
                text = "";
                report.Diagnostics.Add(new FullCorpusDiagnostic("emission", "output.utf8", relative, "", null, null, error.Message));
            }
            var markerCount = Count(text, OwnershipMarker.D9);
            if (markerCount != 1)
            {
                valid = false;
                report.Diagnostics.Add(new FullCorpusDiagnostic("emission", "output.d9", relative, "", null, null, $"Expected one D9 ownership marker; found {markerCount}."));
            }
            try { validator.Validate(relative, text); }
            catch (Exception error)
            {
                valid = false;
                report.Diagnostics.Add(new FullCorpusDiagnostic("emission", "output.csharp73", relative, "", null, null, error.Message));
            }
            if (!valid) invalid++;
        }
        report.InvalidInvariantFiles = invalid;
        report.Gates["csharp73-bom-crlf-d9"] = emittedPaths.Length != 0 && invalid == 0;

        var allowedCopies = new HashSet<string>(new[] { "build-metadata", "manual-companion", "runtime-companion" }, StringComparer.Ordinal);
        var invalidCopies = first.CopiedPaths.Where(copy =>
            !allowedCopies.Contains(copy.Classification) ||
            emittedPaths.Contains(copy.Path, StringComparer.Ordinal) ||
            copy.Path.Contains("/Generated/", StringComparison.Ordinal)).ToArray();
        foreach (var copy in invalidCopies)
            report.Diagnostics.Add(new FullCorpusDiagnostic("orchestration", "output.source-copy", copy.Path, "", null, null,
                $"Isolated output copied a disallowed source path classified as '{copy.Classification}'."));
        report.Gates["no-source-wrapper-copy"] = invalidCopies.Length == 0;

        var firstTree = Directory.Exists(output) ? TreeDigest(output, includeManifest: true) : "";
        var secondTree = Directory.Exists(repeatedOutput) ? TreeDigest(repeatedOutput, includeManifest: true) : "";
        report.OutputDigest = firstTree;
        report.RepeatedOutputDigest = secondTree;
        report.Gates["deterministic-output"] =
            firstTree.Length == 64 && string.Equals(firstTree, secondTree, StringComparison.Ordinal) &&
            !string.IsNullOrWhiteSpace(first.OutputTreeHash) &&
            string.Equals(first.OutputTreeHash, second.OutputTreeHash, StringComparison.Ordinal);
        if (!report.Gates["deterministic-output"])
            report.Diagnostics.Add(new FullCorpusDiagnostic("emission", "output.nondeterministic", "", "", null, null,
                $"First tree {firstTree}/{first.OutputTreeHash}; repeated tree {secondTree}/{second.OutputTreeHash}."));

        InspectManifest(output, emittedPaths, report);
    }

    private static void InspectManifest(string output, IReadOnlyCollection<string> generatedPaths, FullCorpusReport report)
    {
        var manifestPath = Path.Combine(output, ".codegen", "ownership.json");
        if (!File.Exists(manifestPath))
        {
            report.Gates["complete-manifest"] = false;
            report.Diagnostics.Add(new FullCorpusDiagnostic("storage", "manifest.missing", manifestPath, "", null, null, "Generation did not create an ownership manifest."));
            return;
        }
        try
        {
            var manifest = OwnershipManifest.Parse(File.ReadAllBytes(manifestPath));
            report.ManifestFiles = manifest.Files.Count;
            var listed = manifest.Files.Select(static item => item.Path).Order(StringComparer.Ordinal).ToArray();
            var complete = listed.SequenceEqual(generatedPaths.Order(StringComparer.Ordinal), StringComparer.Ordinal);
            foreach (var file in manifest.Files)
            {
                var path = Path.Combine(output, file.Path.Replace('/', Path.DirectorySeparatorChar));
                if (!File.Exists(path) || new FileInfo(path).Length != file.Length || !string.Equals(FileDigest(path), file.Sha256, StringComparison.Ordinal))
                {
                    complete = false;
                    report.Diagnostics.Add(new FullCorpusDiagnostic("storage", "manifest.entry", file.Path, "", null, null, "Manifest digest or length does not match generated bytes."));
                }
            }
            report.Gates["complete-manifest"] = complete && manifest.Files.Count > 77;
            if (!complete)
                report.Diagnostics.Add(new FullCorpusDiagnostic("storage", "manifest.incomplete", manifestPath, "", null, null, "Manifest paths do not exactly equal all emitted managed paths."));
        }
        catch (Exception error)
        {
            report.Gates["complete-manifest"] = false;
            report.Diagnostics.Add(new FullCorpusDiagnostic("storage", "manifest.invalid", manifestPath, "", null, null, error.Message));
        }
    }

    private static void VerifySemanticParity(
        string output,
        string contractsPath,
        IReadOnlyCollection<string> generatedProducts,
        string workRoot,
        FullCorpusReport report)
    {
        var semanticReportPath = Path.Combine(workRoot, "semantic-diff.json");
        var extractor = FindContractExtractorAssembly(report.SourcePath);
        if (extractor is null)
        {
            report.Gates["semantic-parity-shape"] = false;
            report.Gates["semantic-parity"] = false;
            report.Diagnostics.Add(new FullCorpusDiagnostic("parity-contract", "semantic.tool-missing", "", "", null, null,
                "The built NetOffice.CodeGen.ContractExtractor assembly was not found."));
            return;
        }

        var result = RunProcess("dotnet", new[]
        {
            extractor, "compare", "--expected", Path.GetFullPath(contractsPath),
            "--actual", output, "--report", semanticReportPath
        }, Path.GetDirectoryName(report.SourcePath)!, TimeSpan.FromMinutes(10));
        if (!File.Exists(semanticReportPath))
        {
            report.Gates["semantic-parity-shape"] = false;
            report.Gates["semantic-parity"] = false;
            report.Diagnostics.Add(new FullCorpusDiagnostic("parity-contract", "semantic.report-missing", semanticReportPath, "", null, null,
                FirstUsefulLine(result.Stderr, result.Stdout)));
            return;
        }

        try
        {
            using var document = JsonDocument.Parse(File.ReadAllText(semanticReportPath));
            var root = document.RootElement;
            var profile = root.GetProperty("NormalizationProfile").GetString() ?? "";
            var kind = root.GetProperty("ContractKind").GetString() ?? "";
            var summary = root.GetProperty("Summary");
            var unexplained = summary.GetProperty("UnexplainedDifferences").GetInt32();
            var facets = summary.GetProperty("ByFacet").EnumerateObject()
                .OrderBy(static item => item.Name, StringComparer.Ordinal)
                .ToDictionary(static item => item.Name, static item => item.Value.GetInt32(), StringComparer.Ordinal);
            var products = new List<SemanticParityProduct>();
            foreach (var product in root.GetProperty("Products").EnumerateArray())
            {
                var api = product.GetProperty("Api").GetString() ?? "";
                var productSummary = product.GetProperty("Summary");
                var productUnexplained = productSummary.GetProperty("UnexplainedDifferences").GetInt32();
                products.Add(new SemanticParityProduct(api, productUnexplained == 0 ? "matched" : "divergent", 0, 0, productUnexplained));
                foreach (var difference in product.GetProperty("Differences").EnumerateArray())
                {
                    if (difference.TryGetProperty("Intentional", out var intentional) && intentional.ValueKind == JsonValueKind.True) continue;
                    var logicalId = difference.GetProperty("LogicalId").GetString() ?? "";
                    var facet = difference.GetProperty("Facet").GetString() ?? "unknown";
                    var differenceKind = difference.GetProperty("DifferenceKind").GetString() ?? "unknown";
                    var expected = difference.GetProperty("Expected").ValueKind == JsonValueKind.Null ? "<missing>" : difference.GetProperty("Expected").GetString() ?? "";
                    var actual = difference.GetProperty("Actual").ValueKind == JsonValueKind.Null ? "<missing>" : difference.GetProperty("Actual").GetString() ?? "";
                    report.Diagnostics.Add(new FullCorpusDiagnostic("projection", "semantic." + facet, semanticReportPath, api, null, null,
                        $"{logicalId}: {differenceKind}; expected '{expected}', actual '{actual}'."));
                }
            }
            var expectedProducts = generatedProducts.Order(StringComparer.Ordinal).ToArray();
            var actualProducts = products.Select(static item => item.Product).Order(StringComparer.Ordinal).ToArray();
            var shapeValid =
                kind == "NetOffice.WrapperSemanticDiff" &&
                profile == "csharp-semantic/v1" &&
                summary.GetProperty("Products").GetInt32() == ExpectedWrapperProducts &&
                products.Count == ExpectedWrapperProducts &&
                actualProducts.SequenceEqual(expectedProducts, StringComparer.Ordinal) &&
                root.GetProperty("Products").EnumerateArray().All(product =>
                    product.GetProperty("NormalizationProfile").GetString() == "csharp-semantic/v1");
            report.SemanticParity = new SemanticParityReport(
                "semantic-parity/v1", profile, products, facets, unexplained,
                unexplained == 0 ? "none" : "projection");
            report.Gates["semantic-parity-shape"] = shapeValid;
            report.Gates["semantic-parity"] = result.ExitCode == 0 && unexplained == 0;
            if (!shapeValid)
                report.Diagnostics.Add(new FullCorpusDiagnostic("parity-contract", "semantic.report-shape", semanticReportPath, "", null, null,
                    "ContractExtractor did not produce the required 12-product csharp-semantic/v1 report."));
            if (result.ExitCode != 0 && unexplained == 0)
                report.Diagnostics.Add(new FullCorpusDiagnostic("parity-contract", "semantic.compare-failed", semanticReportPath, "", null, null,
                    FirstUsefulLine(result.Stderr, result.Stdout)));
        }
        catch (Exception error)
        {
            report.Gates["semantic-parity-shape"] = false;
            report.Gates["semantic-parity"] = false;
            report.Diagnostics.Add(new FullCorpusDiagnostic("parity-contract", "semantic.report-invalid", semanticReportPath, "", null, null, error.Message));
        }
    }

    private static void BuildGeneratedLibraries(string output, FullCorpusReport report, string workRoot)
    {
        var projectPaths = Directory.Exists(Path.Combine(output, "Source"))
            ? Directory.EnumerateFiles(Path.Combine(output, "Source"), "*.csproj", SearchOption.AllDirectories)
                .Where(path => string.Equals(Path.GetFileName(path), "NetOffice.csproj", StringComparison.OrdinalIgnoreCase) ||
                    Path.GetFileName(path).EndsWith("Api.csproj", StringComparison.OrdinalIgnoreCase))
                .Order(StringComparer.Ordinal)
                .ToArray()
            : Array.Empty<string>();
        if (projectPaths.Length != ExpectedWrapperProducts + 1)
        {
            report.GeneratedLibraries = Array.Empty<GeneratedLibraryBuild>();
            report.Gates["isolated-generated-library-builds"] = false;
            report.Diagnostics.Add(new FullCorpusDiagnostic("orchestration", "build.projects-missing", output, "", null, null,
                $"Expected {ExpectedWrapperProducts + 1} isolated runtime/wrapper projects; found {projectPaths.Length}."));
            return;
        }
        var solutionRoot = Path.Combine(workRoot, "generated-solution");
        Directory.CreateDirectory(solutionRoot);
        var create = RunProcess("dotnet", new[]
        {
            "new", "sln", "--name", "NetOffice.Generated", "--format", "sln", "--output", solutionRoot
        }, workRoot, TimeSpan.FromMinutes(2));
        var solution = Path.Combine(solutionRoot, "NetOffice.Generated.sln");
        var add = create.ExitCode == 0
            ? RunProcess("dotnet", new[] { "sln", solution, "add" }.Concat(projectPaths), solutionRoot, TimeSpan.FromMinutes(2))
            : new ProcessResult(1, "", "Solution creation failed.", false);
        if (create.ExitCode != 0 || add.ExitCode != 0 || !File.Exists(solution))
        {
            report.GeneratedLibraries = Array.Empty<GeneratedLibraryBuild>();
            report.Gates["isolated-generated-library-builds"] = false;
            report.Diagnostics.Add(new FullCorpusDiagnostic("orchestration", "build.solution-create", solution, "", null, null,
                FirstUsefulLine(create.Stderr, create.Stdout, add.Stderr, add.Stdout)));
            return;
        }

        var artifacts = Path.Combine(workRoot, "build-artifacts");
        var restore = RunProcess("dotnet", new[]
        {
            "restore", solution, "--locked-mode", "--nologo",
            "--artifacts-path", artifacts,
            "-p:RestorePackagesWithLockFile=true",
            "-p:RestoreLockedMode=true"
        }, output, TimeSpan.FromMinutes(10));
        var build = restore.ExitCode == 0
            ? RunProcess("dotnet", new[]
            {
                "build", solution, "-c", "Release", "--no-restore", "--nologo",
                "--artifacts-path", artifacts,
                "-p:RestorePackagesWithLockFile=true",
                "-p:RestoreLockedMode=true"
            }, output, TimeSpan.FromMinutes(20))
            : new ProcessResult(1, "", "Build skipped because locked restore failed.", false);

        var uniqueDiagnostics = new HashSet<FullCorpusDiagnostic>();
        foreach (var diagnostic in ParseBuildDiagnostics("all", restore.Stdout, restore.Stderr, build.Stdout, build.Stderr))
            if (uniqueDiagnostics.Add(diagnostic)) report.Diagnostics.Add(diagnostic);
        if (restore.ExitCode != 0 && !uniqueDiagnostics.Any(static item => item.Code.StartsWith("NU", StringComparison.Ordinal) || item.Code.StartsWith("MSB", StringComparison.Ordinal)))
        {
            var diagnostic = new FullCorpusDiagnostic("orchestration", "restore.failed", solution, "all", null, null, FirstUsefulLine(restore.Stderr, restore.Stdout));
            if (uniqueDiagnostics.Add(diagnostic)) report.Diagnostics.Add(diagnostic);
        }
        if (build.ExitCode != 0 && restore.ExitCode == 0 && !uniqueDiagnostics.Any(static item => item.Severity == "error"))
        {
            var diagnostic = new FullCorpusDiagnostic("orchestration", "build.failed", solution, "all", null, null, FirstUsefulLine(build.Stderr, build.Stdout));
            if (uniqueDiagnostics.Add(diagnostic)) report.Diagnostics.Add(diagnostic);
        }

        var exitCode = restore.ExitCode == 0 ? build.ExitCode : restore.ExitCode;
        report.GeneratedLibraries = new[] { new GeneratedLibraryBuild("all", solution, exitCode, uniqueDiagnostics.Count) };
        report.Gates["isolated-generated-library-builds"] = restore.ExitCode == 0 && build.ExitCode == 0;
    }

    private static CodegenGenerationReport ReadCodegenReport(string path)
    {
        using var document = JsonDocument.Parse(File.ReadAllText(path));
        var root = document.RootElement;
        var counts = root.GetProperty("entityCounts");
        var emitted = root.GetProperty("emittedPaths").EnumerateArray().Select(static item => item.GetString() ?? "").ToArray();
        var copied = root.GetProperty("copiedPaths").EnumerateArray()
            .Select(static item => new CodegenCopiedPath(
                item.GetProperty("path").GetString() ?? "",
                item.GetProperty("classification").GetString() ?? ""))
            .ToArray();
        return new CodegenGenerationReport(
            root.GetProperty("schemaVersion").GetString() ?? "",
            root.GetProperty("isolatedOutput").GetBoolean(),
            root.GetProperty("exitCode").GetInt32(),
            counts.GetProperty("inputTypes").GetInt32(),
            counts.GetProperty("inputMembers").GetInt32(),
            counts.GetProperty("inputValues").GetInt32(),
            counts.GetProperty("selectedProducts").GetInt32(),
            counts.GetProperty("projectedFiles").GetInt32(),
            root.GetProperty("products").EnumerateArray().Select(static item => item.GetString() ?? "").ToArray(),
            emitted,
            copied,
            root.GetProperty("outputTreeHash").GetString() ?? "");
    }

    private static string ResolveOutputPath(string output, string relative)
    {
        var root = Path.GetFullPath(output) + Path.DirectorySeparatorChar;
        var path = Path.GetFullPath(Path.Combine(root, relative.Replace('/', Path.DirectorySeparatorChar)));
        if (!path.StartsWith(root, StringComparison.OrdinalIgnoreCase))
            throw new InvalidDataException("Reported output path escapes the isolated output root: " + relative);
        return path;
    }

    private static bool HasOnlyCrLfWithFinalNewline(byte[] bytes, int offset)
    {
        if (bytes.Length - offset < 2 || bytes[^2] != 0x0d || bytes[^1] != 0x0a) return false;
        for (var index = offset; index < bytes.Length; index++)
        {
            if (bytes[index] == 0x0a && (index == offset || bytes[index - 1] != 0x0d)) return false;
            if (bytes[index] == 0x0d && (index + 1 >= bytes.Length || bytes[index + 1] != 0x0a)) return false;
        }
        return true;
    }

    private static string? FindContractExtractorAssembly(string sourcePath)
    {
        var repositoryRoot = Path.GetDirectoryName(Path.GetFullPath(sourcePath));
        if (repositoryRoot is null) return null;
        var projectRoot = Path.Combine(repositoryRoot, "Tools", "CodeGen", "tools", "NetOffice.CodeGen.ContractExtractor");
        var configuration = new DirectoryInfo(AppContext.BaseDirectory).Parent?.Name;
        if (!string.IsNullOrWhiteSpace(configuration))
        {
            var preferred = Path.Combine(projectRoot, "bin", configuration);
            if (Directory.Exists(preferred))
            {
                var match = Directory.EnumerateFiles(preferred, "NetOffice.CodeGen.ContractExtractor.dll", SearchOption.AllDirectories)
                    .FirstOrDefault(path => !path.Contains(Path.DirectorySeparatorChar + "ref" + Path.DirectorySeparatorChar, StringComparison.OrdinalIgnoreCase));
                if (match is not null) return match;
            }
        }
        var bin = Path.Combine(projectRoot, "bin");
        return Directory.Exists(bin)
            ? Directory.EnumerateFiles(bin, "NetOffice.CodeGen.ContractExtractor.dll", SearchOption.AllDirectories)
                .Where(path => !path.Contains(Path.DirectorySeparatorChar + "ref" + Path.DirectorySeparatorChar, StringComparison.OrdinalIgnoreCase))
                .OrderByDescending(static path => File.GetLastWriteTimeUtc(path))
                .FirstOrDefault()
            : null;
    }

    internal static string[] CopyManualCompanions(string product, string sourceRoot, string contractsRoot, string generatedProjectDirectory)
    {
        var classificationPath = Path.Combine(contractsRoot, product + ".classification.json");
        if (!File.Exists(classificationPath)) return Array.Empty<string>();
        using var document = JsonDocument.Parse(File.ReadAllText(classificationPath));
        if (!document.RootElement.TryGetProperty("Files", out var files) || files.ValueKind != JsonValueKind.Array) return Array.Empty<string>();
        var sourceProductRoot = Path.GetFullPath(Path.Combine(sourceRoot, product)) + Path.DirectorySeparatorChar;
        var companionRoot = Path.Combine(generatedProjectDirectory, "Companions");
        var copied = new List<string>();
        foreach (var file in files.EnumerateArray())
        {
            if (!file.TryGetProperty("Ownership", out var ownership) || !string.Equals(ownership.GetString(), "manual", StringComparison.OrdinalIgnoreCase)) continue;
            if (!file.TryGetProperty("RequiredForIsolatedBuild", out var required) || required.ValueKind != JsonValueKind.True) continue;
            if (!file.TryGetProperty("Path", out var pathValue) || string.IsNullOrWhiteSpace(pathValue.GetString())) continue;
            var relative = pathValue.GetString()!.Replace('/', Path.DirectorySeparatorChar);
            if (!relative.EndsWith(".cs", StringComparison.OrdinalIgnoreCase)) continue;
            var source = Path.GetFullPath(Path.Combine(sourceProductRoot, relative));
            if (!source.StartsWith(sourceProductRoot, StringComparison.OrdinalIgnoreCase) || !File.Exists(source)) continue;
            var destination = Path.Combine(companionRoot, relative);
            Directory.CreateDirectory(Path.GetDirectoryName(destination)!);
            File.Copy(source, destination, overwrite: false);
            copied.Add(destination);
        }
        return copied.Order(StringComparer.Ordinal).ToArray();
    }


    private static IEnumerable<FullCorpusDiagnostic> ParseBuildDiagnostics(string product, params string[] streams)
    {
        var seen = new HashSet<string>(StringComparer.Ordinal);
        foreach (var line in streams.SelectMany(static stream => stream.Split(new[] { "\r\n", "\n" }, StringSplitOptions.RemoveEmptyEntries)))
        {
            var match = BuildDiagnosticPattern.Match(line.Trim());
            if (!match.Success) continue;
            var code = match.Groups["code"].Value;
            var file = match.Groups["file"].Value;
            var message = match.Groups["message"].Value.Trim();
            var diagnosticProduct = ProductFromDiagnostic(file, match.Groups["project"].Value, product);
            var key = string.Join("|", diagnosticProduct, file, match.Groups["line"].Value, match.Groups["column"].Value, code, message);
            if (!seen.Add(key)) continue;
            int? lineNumber = int.TryParse(match.Groups["line"].Value, out var parsedLine) ? parsedLine : null;
            int? column = int.TryParse(match.Groups["column"].Value, out var parsedColumn) ? parsedColumn : null;
            yield return new FullCorpusDiagnostic(CompileOwner(code, message), code, file, diagnosticProduct, lineNumber, column, message)
            {
                Severity = match.Groups["severity"].Value.ToLowerInvariant()
            };
        }
    }

    private static string GenerationOwner(params string[] output)
    {
        var message = string.Join("\n", output);
        if (message.Contains("contract.", StringComparison.OrdinalIgnoreCase) || message.Contains("Wrapper Contract", StringComparison.OrdinalIgnoreCase)) return "parity-contract";
        if (message.Contains("digest.", StringComparison.OrdinalIgnoreCase) || message.Contains("data.", StringComparison.OrdinalIgnoreCase)) return "data-v2";
        if (message.Contains("policy.", StringComparison.OrdinalIgnoreCase) || message.Contains("projection", StringComparison.OrdinalIgnoreCase)) return "projection";
        if (message.Contains("manifest", StringComparison.OrdinalIgnoreCase)) return "storage";
        if (message.Contains("C# 7.3", StringComparison.OrdinalIgnoreCase) || message.Contains("emit", StringComparison.OrdinalIgnoreCase)) return "emission";
        return "orchestration";
    }

    private static string ProductFromDiagnostic(string file, string project, string fallback)
    {
        var parts = (file + "/" + project).Replace('\\', '/').Split('/', StringSplitOptions.RemoveEmptyEntries);
        for (var index = 0; index + 1 < parts.Length; index++)
            if (string.Equals(parts[index], "Source", StringComparison.OrdinalIgnoreCase))
                return parts[index + 1];
        return fallback;
    }

    private static string CompileOwner(string code, string message)
    {
        if (code.StartsWith("NU", StringComparison.Ordinal) || code.StartsWith("MSB", StringComparison.Ordinal) || code is "CS0006" or "CS0012" or "CS0518" or "CS1705") return "orchestration";
        if (code == "CS1591") return "documentation";
        if (code is "CS1001" or "CS1002" or "CS1003" or "CS1009" or "CS1010" or "CS1022" or "CS1031" or "CS1056" or "CS1513" or "CS1514" or "CS1519" or "CS1525" or "CS8124") return "emission";
        if (code.StartsWith("CS", StringComparison.Ordinal)) return "projection";
        return message.Contains("manifest", StringComparison.OrdinalIgnoreCase) ? "storage" : "orchestration";
    }

    private static WrapperContract EmptyContract() => new()
    {
        SchemaVersion = "1.0",
        ContractKind = "NetOffice.WrapperContract",
        Generator = new ContractGenerator { Name = "NetOffice.CodeGen.Verification", Version = "1" },
        Source = new ContractSource { Root = "full-corpus", Api = "__full_corpus_data_only__", SourceRoots = new[] { "Data v2" } },
        Types = Array.Empty<ContractType>()
    };

    private static string[] ContractFiles(string path)
    {
        if (File.Exists(path)) return new[] { Path.GetFullPath(path) };
        if (!Directory.Exists(path)) return Array.Empty<string>();
        return Directory.EnumerateFiles(path, "*.wrapper-contract.json", SearchOption.TopDirectoryOnly).Order(StringComparer.Ordinal).ToArray();
    }

    private static WrapperProject[] DiscoverWrapperProjects(string sourceRoot)
    {
        if (!Directory.Exists(sourceRoot)) return Array.Empty<WrapperProject>();
        return Directory.EnumerateDirectories(sourceRoot)
            .Select(directory => new { Directory = directory, Projects = Directory.EnumerateFiles(directory, "*Api.csproj", SearchOption.TopDirectoryOnly).ToArray() })
            .Where(static item => item.Projects.Length == 1 && !item.Directory.EndsWith("Office.Extensions", StringComparison.OrdinalIgnoreCase))
            .Select(static item => new WrapperProject(new DirectoryInfo(item.Directory).Name, item.Projects[0]))
            .OrderBy(static item => item.Product, StringComparer.Ordinal)
            .ToArray();
    }

    private static ProcessResult RunCli(string command, IEnumerable<string> arguments, TimeSpan timeout)
    {
        var cli = Path.Combine(AppContext.BaseDirectory, "NetOffice.CodeGen.Cli.dll");
        if (!File.Exists(cli)) return new ProcessResult(1, "", "Built CLI not found: " + cli, false);
        return RunProcess("dotnet", new[] { cli, command }.Concat(arguments), AppContext.BaseDirectory, timeout);
    }

    private static ProcessResult RunProcess(string fileName, IEnumerable<string> arguments, string workingDirectory, TimeSpan timeout)
    {
        var info = new ProcessStartInfo(fileName)
        {
            RedirectStandardOutput = true,
            RedirectStandardError = true,
            UseShellExecute = false,
            WorkingDirectory = workingDirectory,
            CreateNoWindow = true
        };
        foreach (var argument in arguments) info.ArgumentList.Add(argument);
        using var process = Process.Start(info) ?? throw new InvalidOperationException("Unable to start " + fileName + ".");
        var stdout = process.StandardOutput.ReadToEndAsync();
        var stderr = process.StandardError.ReadToEndAsync();
        var exited = process.WaitForExit((int)Math.Min(int.MaxValue, timeout.TotalMilliseconds));
        if (!exited)
        {
            process.Kill(entireProcessTree: true);
            process.WaitForExit();
        }
        Task.WaitAll(stdout, stderr);
        return new ProcessResult(exited ? process.ExitCode : 1, stdout.Result, stderr.Result, !exited);
    }

    private static string NormalizeCommand(IEnumerable<string> common, string output)
        => "dotnet <NetOffice.CodeGen.Cli.dll> generate " + string.Join(" ", common.Select(static value => value.Contains(' ') ? "<path>" : value)) + " --output " + output;

    private static int Count(string value, string needle)
    {
        var count = 0;
        for (var index = 0; (index = value.IndexOf(needle, index, StringComparison.Ordinal)) >= 0; index += needle.Length) count++;
        return count;
    }


    private static string TreeDigest(string root, bool includeManifest)
    {
        var files = Directory.EnumerateFiles(root, "*", SearchOption.AllDirectories)
            .Where(path => includeManifest || !path.Contains(Path.DirectorySeparatorChar + ".codegen" + Path.DirectorySeparatorChar, StringComparison.Ordinal))
            .Select(path => Path.GetRelativePath(root, path).Replace('\\', '/'))
            .Order(StringComparer.Ordinal)
            .ToArray();
        return Hex(SHA256.HashData(Encoding.UTF8.GetBytes(string.Join("\n", files.Select(path => path + ":" + FileDigest(Path.Combine(root, path.Replace('/', Path.DirectorySeparatorChar))))) + "\n")));
    }

    private static string FileDigest(string path) => Hex(SHA256.HashData(File.ReadAllBytes(path)));
    private static string Hex(byte[] bytes) => Convert.ToHexString(bytes).ToLowerInvariant();
    private static string FirstUsefulLine(params string[] values) => values.SelectMany(static value => value.Split(new[] { "\r\n", "\n" }, StringSplitOptions.RemoveEmptyEntries)).FirstOrDefault() ?? "Command failed without output.";
    private static readonly JsonSerializerOptions JsonOptions = new(JsonSerializerDefaults.Web) { WriteIndented = true };

    private static void FinalizeReport(FullCorpusReport report)
    {
        var order = new[] { "data-v2", "projection", "documentation", "emission", "storage", "orchestration", "parity-contract", "verification" };
        report.DiagnosticsByOwner = report.Diagnostics
            .GroupBy(static item => item.Owner, StringComparer.Ordinal)
            .OrderBy(group => Array.IndexOf(order, group.Key) is var index && index >= 0 ? index : int.MaxValue)
            .ToDictionary(static group => group.Key, static group => (IReadOnlyList<FullCorpusDiagnostic>)group.OrderBy(static item => item.Product, StringComparer.Ordinal).ThenBy(static item => item.Path, StringComparer.Ordinal).ThenBy(static item => item.Line).ThenBy(static item => item.Code, StringComparer.Ordinal).ToArray(), StringComparer.Ordinal);
        var divergentOwners = report.Diagnostics.Where(static diagnostic => diagnostic.Severity == "error").Select(static diagnostic => diagnostic.Owner).ToHashSet(StringComparer.Ordinal);
        report.FirstDivergentOwner = order.FirstOrDefault(divergentOwners.Contains)
            ?? divergentOwners.Order(StringComparer.Ordinal).FirstOrDefault()
            ?? "none";
        var compileDiagnostics = report.Diagnostics.Where(static diagnostic =>
            diagnostic.Code.StartsWith("CS", StringComparison.Ordinal) ||
            diagnostic.Code.StartsWith("NU", StringComparison.Ordinal) ||
            diagnostic.Code.StartsWith("MSB", StringComparison.Ordinal) ||
            diagnostic.Code.StartsWith("build.", StringComparison.Ordinal) ||
            diagnostic.Code.StartsWith("restore.", StringComparison.Ordinal)).ToArray();
        var compileByOwner = compileDiagnostics
            .GroupBy(static item => item.Owner, StringComparer.Ordinal)
            .ToDictionary(static group => group.Key, static group => (IReadOnlyList<FullCorpusDiagnostic>)group.ToArray(), StringComparer.Ordinal);
        var compileErrorOwners = compileDiagnostics.Where(static diagnostic => diagnostic.Severity == "error").Select(static diagnostic => diagnostic.Owner).ToHashSet(StringComparer.Ordinal);
        var firstCompileOwner = order.FirstOrDefault(compileErrorOwners.Contains)
            ?? compileErrorOwners.Order(StringComparer.Ordinal).FirstOrDefault()
            ?? "none";
        report.Compilation = new GeneratedCompilationReport(
            "generated-library-compilation/v1",
            report.GeneratedLibraries.All(static library => library.ExitCode == 0) && report.GeneratedLibraries.Count != 0 ? "passed" : "failed",
            firstCompileOwner,
            report.GeneratedLibraries,
            compileDiagnostics.Length,
            compileByOwner);
        report.Outcome = report.Gates.Count != 0 && report.Gates.Values.All(static value => value) ? "passed" : "failed";
    }

    private sealed record CodegenGenerationReport(
        string SchemaVersion,
        bool IsolatedOutput,
        int ExitCode,
        int InputTypes,
        int InputMembers,
        int InputValues,
        int SelectedProducts,
        int ProjectedFiles,
        IReadOnlyList<string> Products,
        IReadOnlyList<string> EmittedPaths,
        IReadOnlyList<CodegenCopiedPath> CopiedPaths,
        string OutputTreeHash);
    private sealed record CodegenCopiedPath(string Path, string Classification);
    private sealed record WrapperProject(string Product, string ProjectPath);
    private sealed record ProcessResult(int ExitCode, string Stdout, string Stderr, bool TimedOut);
}

internal sealed class FullCorpusReport
{
    public string SchemaVersion { get; init; } = "netoffice-codegen-full-corpus/v1";
    public string Outcome { get; set; } = "failed";
    public string FirstDivergentOwner { get; set; } = "none";
    public string DataPath { get; init; } = "";
    public string SourcePath { get; init; } = "";
    public string ContractsPath { get; init; } = "";
    public string GraphDigest { get; set; } = "";
    public IReadOnlyList<string> InputProducts { get; set; } = Array.Empty<string>();
    public IReadOnlyList<string> ContractFiles { get; set; } = Array.Empty<string>();
    public FullCorpusCoverage Coverage { get; set; } = new(0, 0, 0, 0, 0, 0);
    public int ExpectedFiles { get; set; }
    public int AdditionalFiles { get; set; }
    public int ActualFiles { get; set; }
    public int InvalidInvariantFiles { get; set; }
    public int ManifestFiles { get; set; }
    public string OutputDigest { get; set; } = "";
    public string RepeatedOutputDigest { get; set; } = "";
    public FullCorpusGeneration? Generation { get; set; }
    public SemanticParityReport? SemanticParity { get; set; }
    public GeneratedCompilationReport? Compilation { get; set; }
    public IReadOnlyList<GeneratedLibraryBuild> GeneratedLibraries { get; set; } = Array.Empty<GeneratedLibraryBuild>();
    public SortedDictionary<string, bool> Gates { get; } = new(StringComparer.Ordinal);
    [System.Text.Json.Serialization.JsonIgnore]
    public List<FullCorpusDiagnostic> Diagnostics { get; } = new();
    public IReadOnlyDictionary<string, IReadOnlyList<FullCorpusDiagnostic>> DiagnosticsByOwner { get; set; } = new Dictionary<string, IReadOnlyList<FullCorpusDiagnostic>>();
    public bool Gate(string name) => Gates.TryGetValue(name, out var passed) && passed;
    public string ToJson() => JsonSerializer.Serialize(this, new JsonSerializerOptions { WriteIndented = true, PropertyNamingPolicy = JsonNamingPolicy.CamelCase, DictionaryKeyPolicy = JsonNamingPolicy.CamelCase });
}

internal sealed record FullCorpusCoverage(
    int DataTypes,
    int ProjectedTypes,
    int DataMembers,
    int ProjectedDataMembers,
    int DataValues,
    int ProjectedValues);

internal sealed record FullCorpusGeneration(
    int FirstExitCode,
    int SecondExitCode,
    string Command,
    string FirstReport,
    string SecondReport,
    string FirstStdout,
    string FirstStderr,
    string SecondStdout,
    string SecondStderr);

internal sealed record FullCorpusDiagnostic(
    string Owner,
    string Code,
    string Path,
    string Product,
    int? Line,
    int? Column,
    string Message)
{
    public string Severity { get; init; } = "error";
}

internal sealed record GeneratedCompilationReport(
    string SchemaVersion,
    string Outcome,
    string FirstDivergentOwner,
    IReadOnlyList<GeneratedLibraryBuild> Libraries,
    int DiagnosticCount,
    IReadOnlyDictionary<string, IReadOnlyList<FullCorpusDiagnostic>> DiagnosticsByOwner);

internal sealed record SemanticParityReport(
    string SchemaVersion,
    string NormalizationProfile,
    IReadOnlyList<SemanticParityProduct> Products,
    IReadOnlyDictionary<string, int> Facets,
    int DiagnosticCount,
    string FirstDivergentOwner);

internal sealed record SemanticParityProduct(
    string Product,
    string Status,
    int ExpectedTypes,
    int ActualTypes,
    int DiagnosticCount);

internal sealed record GeneratedLibraryBuild(
    string Product,
    string ProjectPath,
    int ExitCode,
    int DiagnosticCount);
