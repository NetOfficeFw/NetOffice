using System.Collections.Immutable;
using System.Diagnostics;
using System.Security.Cryptography;
using System.Text;
using System.Text.Json;
using System.Text.Json.Serialization;
using NetOffice.CodeGen.Data;
using NetOffice.CodeGen.Docs;
using NetOffice.CodeGen.Emit;
using NetOffice.CodeGen.Storage;
using NetOffice.CodeGen.TypeLib;

namespace NetOffice.CodeGen.Verification;

internal static class Program
{
    internal static object? EndToEndEvidence { get; private set; }
    internal static FullCorpusReport? FullCorpusEvidence { get; private set; }

    private static int Main(string[] args)
    {
        var fixtureRoot = Path.Combine(AppContext.BaseDirectory, "fixtures");
        var fullOptions = FullCorpusOptions.Parse(args);
        var checks = new List<(string Id, string Stage, Action Run)>
        {
            ("data-schema-and-digest", "schema", () => VerifyDataSchemaAndDigest(fixtureRoot)),
            ("data-fixture-diagnostics", "schema", () => VerifyDataFixtures(fixtureRoot)),
            ("typelib-diagnostics", "typelib", () => VerifyTypeLibFixtures(fixtureRoot)),
            ("output-invariants", "emission", VerifyOutputInvariants),
            ("manual-companion-isolation", "compilation", VerifyManualCompanionIsolation),
            ("manifest-safety", "storage", VerifyManifestSafety),
            ("documentation-attribution", "documentation", () => VerifyDocumentation(fixtureRoot)),
            ("fault-injection-rollback", "storage", VerifyRollback),
            ("e2e-cli-pipeline", "orchestration", () => VerifyEndToEnd(fixtureRoot)),
        };
        if (fullOptions.Enabled)
        {
            var fullCorpus = new Lazy<FullCorpusReport>(() => FullCorpusEvidence = FullCorpusHarness.Run(fullOptions));
            checks.Add(("full-corpus-projection", "projection", () => RequireFullGates(fullCorpus.Value, "converted-graph-valid", "canonical-corpus-size", "all-products", "projection-coverage")));
            checks.Add(("full-corpus-generation", "orchestration", () => RequireFullGates(fullCorpus.Value, "generation-succeeded", "complete-output")));
            checks.Add(("full-corpus-output-invariants", "emission", () => RequireFullGates(fullCorpus.Value, "csharp73-bom-crlf-d9", "no-source-wrapper-copy")));
            checks.Add(("full-corpus-determinism-manifest", "storage", () => RequireFullGates(fullCorpus.Value, "deterministic-output", "complete-manifest")));
            checks.Add(("full-corpus-semantic-parity-shape", "parity-contract", () => RequireFullGates(fullCorpus.Value, "semantic-parity-shape")));
            checks.Add(("full-corpus-semantic-parity", "projection", () => RequireFullGates(fullCorpus.Value, "semantic-parity")));
            checks.Add(("full-corpus-isolated-build", "compilation", () => RequireFullGates(fullCorpus.Value, "isolated-generated-library-builds")));
        }

        var report = VerificationEngine.Run(checks);
        var reportPath = ReadReportPath(args) ?? Path.Combine(Path.GetTempPath(), "netoffice-codegen-verification-report.json");
        Directory.CreateDirectory(Path.GetDirectoryName(Path.GetFullPath(reportPath))!);
        File.WriteAllText(reportPath, report.ToJson(), new UTF8Encoding(false));
        if (fullOptions.Enabled && FullCorpusEvidence is not null)
        {
            var fullReportPath = fullOptions.ReportPath ?? Path.Combine(Path.GetDirectoryName(Path.GetFullPath(reportPath))!, Path.GetFileNameWithoutExtension(reportPath) + ".full-corpus.json");
            Directory.CreateDirectory(Path.GetDirectoryName(Path.GetFullPath(fullReportPath))!);
            File.WriteAllText(fullReportPath, FullCorpusEvidence.ToJson(), new UTF8Encoding(false));
            Console.WriteLine($"Full-corpus report: {fullReportPath}");
        }
        Console.WriteLine($"Verification report: {reportPath}");
        Console.WriteLine(report.Passed ? "Verification passed." : $"Verification failed at stage '{report.FirstDivergentStage}'.");
        return report.Passed ? 0 : 1;
    }

    private static void RequireFullGates(FullCorpusReport report, params string[] gates)
    {
        var failed = gates.Where(gate => !report.Gate(gate)).ToArray();
        if (failed.Length == 0) return;
        var diagnostics = report.Diagnostics
            .Where(diagnostic => diagnostic.Owner == report.FirstDivergentOwner)
            .Take(5)
            .Select(diagnostic => $"{diagnostic.Code}: {diagnostic.Message}");
        throw new InvalidOperationException($"Full-corpus gates failed: {string.Join(", ", failed)}. First divergent owner: {report.FirstDivergentOwner}. {string.Join(" | ", diagnostics)}");
    }

    private static string? ReadReportPath(IEnumerable<string> args)
    {
        var values = args.ToArray();
        for (var index = 0; index + 1 < values.Length; index++)
            if (string.Equals(values[index], "--report", StringComparison.Ordinal))
                return values[index + 1];
        return null;
    }

    private static void VerifyDataSchemaAndDigest(string fixtureRoot)
    {
        var zero = new string('0', 64);
        var source = new DataSource { Kind = "canonical-input", Path = "fixture.xml", Sha256 = zero, Revision = "verification" };
        var provenance = new Provenance { SourcePath = source.Path, SourceSha256 = source.Sha256, SourceRevision = source.Revision, Location = "/fixture" };
        var graph = new DataGraph
        {
            Source = source,
            Libraries = new[] { new DataLibrary { LogicalId = "library:fixture", Name = "Fixture", Guid = "00000000-0000-0000-0000-000000000201", Version = "1.0", Provenance = provenance } },
            Types = new[] { new DataType { LogicalId = "type:fixture", LibraryId = "library:fixture", Name = "IFixture", Kind = "interface", SourceKey = "fixture", Provenance = provenance } },
            Members = new[] { new DataMember { LogicalId = "member:fixture", TypeId = "type:fixture", Name = "Value", Kind = "property", SourceKey = "value", ReturnType = "string", Provenance = provenance } }
        };
        var digest = CanonicalJson.ComputeDigest(graph);
        graph = graph with { Digest = digest };
        var valid = DataGraphValidator.Validate(graph, digest);
        Assert(valid.IsValid, "a canonical graph with its computed digest must validate");

        var badDigest = DataGraphValidator.Validate(graph with { Digest = zero }, digest);
        Assert(badDigest.Issues.Any(x => x.Code == "digest.mismatch"), "a changed graph digest must be rejected");
        var badSchema = DataGraphValidator.Validate(graph with { SchemaVersion = "2.999" }, digest);
        Assert(badSchema.Issues.Any(x => x.Code == "schema.version"), "an unknown data schema must be rejected");
        Assert(!string.IsNullOrWhiteSpace(DataSchema.Version) && !string.IsNullOrWhiteSpace(DataSchema.Serialization), "the data schema must expose version and serialization pins");
    }

    private static void VerifyDataFixtures(string fixtureRoot)
    {
        var missing = CanonicalJson.Parse(File.ReadAllText(Fixture(fixtureRoot, "data-missing-reference.json")));
        var missingResult = DataGraphValidator.Validate(missing);
        Assert(missingResult.Issues.Any(x => x.Code == "reference.missing"), "missing logical references must be reported");

        var unknown = CanonicalJson.Parse(File.ReadAllText(Fixture(fixtureRoot, "data-unknown-schema.json")));
        var unknownResult = DataGraphValidator.Validate(unknown);
        Assert(unknownResult.Issues.Any(x => x.Code == "schema.version"), "unknown data schema fixtures must report schema.version");

        var ambiguity = CanonicalJson.Parse(File.ReadAllText(Fixture(fixtureRoot, "data-ambiguity.json")));
        var ambiguityWithDigest = ambiguity with { Digest = CanonicalJson.ComputeDigest(ambiguity) };
        var ambiguityResult = DataGraphValidator.Validate(ambiguityWithDigest, ambiguityWithDigest.Digest);
        Assert(ambiguityResult.IsValid, "an explicit unresolved ambiguity is valid input, not an implicit removal");
        Assert(ambiguity.Ambiguities.Count == 1 && ambiguity.Ambiguities[0].Candidates.Count == 2 && ambiguity.Ambiguities[0].Resolution is null, "ambiguity fixtures must retain all candidates until resolved");
    }

    private static void VerifyTypeLibFixtures(string fixtureRoot)
    {
        var missing = ReadObservation(Fixture(fixtureRoot, "typelib/missing-reference.json"));
        var missingGraph = TypeLibDependencyGraph.Build(new[] { missing });
        Assert(!missingGraph.IsValid && missingGraph.UnresolvedReferences.Length == 1, "missing typelib references must remain unresolved");

        var existing = ReadObservation(Fixture(fixtureRoot, "typelib/conflict-existing.json"));
        var incoming = ReadObservation(Fixture(fixtureRoot, "typelib/conflict-incoming.json"));
        var conflict = TypeLibMerger.Merge(new[] { existing }, new[] { incoming });
        Assert(conflict.Conflicts.Any(x => x.Kind == TypeLibMergeConflictKind.DuplicateIdentity), "divergent observations with one identity must conflict");

        var staleIncoming = ReadObservation(Fixture(fixtureRoot, "typelib/stale-incoming.json"));
        var stale = TypeLibMerger.Merge(
            new[] { existing },
            new[] { staleIncoming },
            new TypeLibMergeRequest(new Dictionary<TypeLibIdentity, string> { [existing.Identity] = new string('f', 64) }));
        Assert(stale.Conflicts.Any(x => x.Kind == TypeLibMergeConflictKind.StaleInput), "compare-and-swap base mismatches must be stale conflicts");

        var unknown = ReadObservation(Fixture(fixtureRoot, "typelib/unknown-schema.json"));
        var unknownResult = TypeLibMerger.Merge(Array.Empty<TypeLibObservation>(), new[] { unknown });
        Assert(unknownResult.Conflicts.Any(x => x.Kind == TypeLibMergeConflictKind.UnknownSchema), "unknown typelib schemas must be rejected");

        var owner = ReadObservation(Fixture(fixtureRoot, "typelib/ambiguous-owner.json"));
        var candidates = ReadObservationArray(Fixture(fixtureRoot, "typelib/ambiguous-candidates.json"));
        var ambiguous = TypeLibDependencyGraph.Build(new[] { owner }.Concat(candidates));
        Assert(ambiguous.UnresolvedReferences.Any(x => x.Reason.Contains("multiple", StringComparison.OrdinalIgnoreCase)), "an unversioned reference with multiple candidates must be ambiguous");
    }

    private static void VerifyOutputInvariants()
    {
        var model = new WrapperFile
        {
            RelativePath = "Excel/Generated/Invariant.cs",
            Namespace = "Fixture",
            Usings = new List<string> { "System" },
            Types = new List<WrapperType>
            {
                new WrapperType
                {
                    Name = "Invariant",
                    Documentation = new WrapperDocumentation { Summary = "A & B", Remarks = "C# 7.3 source" },
                    Members = new List<WrapperMember>
                    {
                        new WrapperMember { Name = "Value", Kind = "property", Type = "string", Documentation = new WrapperDocumentation { Summary = "A value" } },
                        new WrapperMember { Name = "Call", Kind = "method", Type = "void", Parameters = new List<WrapperParameter> { new WrapperParameter { Name = "value", Type = "string" } } }
                    }
                }
            }
        };
        var emitted = new CSharpEmitter().Emit(model);
        Assert(emitted.Bytes.Length > 3 && emitted.Bytes[0] == 0xef && emitted.Bytes[1] == 0xbb && emitted.Bytes[2] == 0xbf, "emitted C# must have a UTF-8 BOM");
        Assert(emitted.Text.Contains(OwnershipMarker.D9, StringComparison.Ordinal), "emitted C# must contain the stable ownership marker");
        Assert(!emitted.Text.Contains("<auto-generated", StringComparison.OrdinalIgnoreCase), "the ownership marker must not be the auto-generated marker");
        Assert(emitted.Text.EndsWith("\r\n", StringComparison.Ordinal), "emitted C# must end in a final CRLF");
        Assert(!emitted.Text.Replace("\r\n", string.Empty, StringComparison.Ordinal).Contains('\n'), "emitted C# must not contain lone LF line endings");
        Assert(emitted.Text.Contains("namespace Fixture\r\n{", StringComparison.Ordinal), "emitted source must use C# 7.3-compatible block namespaces");
        Assert(emitted.Text.Contains("<summary>", StringComparison.Ordinal) && emitted.Text.Contains("A &amp; B", StringComparison.Ordinal), "XML documentation must be escaped and retained");
    }

    private static void VerifyManualCompanionIsolation()
    {
        var root = NewTempRoot("companions");
        try
        {
            var source = Path.Combine(root, "source", "Fixture");
            var contracts = Path.Combine(root, "contracts");
            var generatedProject = Path.Combine(root, "generated-project");
            Directory.CreateDirectory(Path.Combine(source, "Manual"));
            Directory.CreateDirectory(Path.Combine(source, "DispatchInterfaces"));
            Directory.CreateDirectory(contracts);
            File.WriteAllText(Path.Combine(source, "Manual", "Helper.cs"), "internal sealed class Helper {}");
            File.WriteAllText(Path.Combine(source, "DispatchInterfaces", "Wrapper.cs"), "public sealed class Wrapper {}");
            File.WriteAllText(Path.Combine(contracts, "Fixture.classification.json"),
                "{\"Files\":[{\"Path\":\"Manual/Helper.cs\",\"Ownership\":\"manual\",\"RequiredForIsolatedBuild\":true},{\"Path\":\"DispatchInterfaces/Wrapper.cs\",\"Ownership\":\"wrapper-generated\",\"RequiredForIsolatedBuild\":true}]}");

            var copied = FullCorpusHarness.CopyManualCompanions("Fixture", Path.Combine(root, "source"), contracts, generatedProject);
            Assert(copied.Length == 1 && copied[0].EndsWith(Path.Combine("Manual", "Helper.cs"), StringComparison.Ordinal), "isolated builds must copy only classified manual companions");
            Assert(!Directory.EnumerateFiles(generatedProject, "*.cs", SearchOption.AllDirectories).Any(path => path.EndsWith("Wrapper.cs", StringComparison.Ordinal)), "isolated builds must never copy source wrappers");
        }
        finally { DeleteTempRoot(root); }
    }

    private static void VerifyManifestSafety()
    {
        var root = NewTempRoot("manifest");
        try
        {
            var emitted = EmitFixture("Stable");
            var initial = GenerationPlanBuilder.Build(root, new[] { emitted }, null, new PlanOptions { GeneratorVersion = "verification" });
            initial.Apply();
            var manifestPath = Path.Combine(root, ".codegen", "ownership.json");
            var manifestBytes = File.ReadAllBytes(manifestPath);
            var manifest = OwnershipManifest.Parse(manifestBytes);
            Assert(manifest.Files.Count == 1 && manifest.Files[0].Path == emitted.RelativePath, "the ownership manifest must contain the generated path and digest");

            var noOp = GenerationPlanBuilder.Build(root, new[] { emitted }, manifest, new PlanOptions { GeneratorVersion = "verification" });
            Assert(!noOp.HasChanges, "a byte-identical managed tree must produce no operations");

            var managedPath = Path.Combine(root, emitted.RelativePath.Replace('/', Path.DirectorySeparatorChar));
            File.AppendAllText(managedPath, "drift", Encoding.UTF8);
            AssertThrows<HashMismatchException>(() => GenerationPlanBuilder.Build(root, new[] { emitted }, manifest), "managed drift must abort before a write plan is applied");
            File.WriteAllBytes(managedPath, emitted.Bytes);

            var unmanagedManifest = OwnershipManifest.Create("verification", new[] { new ManifestFile { Path = "Excel/Stable.cs", Sha256 = emitted.Sha256, Length = emitted.Bytes.LongLength } });
            AssertThrows<InvalidOperationException>(() => GenerationPlanBuilder.Build(root, Array.Empty<EmittedFile>(), unmanagedManifest), "deletion outside a Generated root must be refused");

            var before = File.ReadAllBytes(manifestPath);
            var checkOnly = GenerationPlanBuilder.Build(root, new[] { EmitFixture("Changed") }, manifest, new PlanOptions { CheckOnly = true });
            AssertThrows<InvalidOperationException>(checkOnly.Apply, "check-only plans must never write files");
            Assert(File.ReadAllBytes(manifestPath).SequenceEqual(before), "check-only plans must preserve the manifest");
        }
        finally { DeleteTempRoot(root); }
    }

    private static void VerifyDocumentation(string fixtureRoot)
    {
        var docsRoot = Fixture(fixtureRoot, "docs");
        var result = DocumentationSync.Sync(
            new[] { "excel" },
            new DocsOptions(docsRoot, Path.Combine(docsRoot, "pin.json"), DocsProfile.Vba, Locked: true));
        Assert(result.Pin is not null && result.Pin.Repository.Equals("MicrosoftDocs/VBA-Docs", StringComparison.OrdinalIgnoreCase), "locked VBA docs must retain the approved repository pin");
        Assert(result.Mappings.Count == 1 && result.Mappings[0].Kind == MappingKind.Exact, "an exact article match must be recorded");
        Assert(result.Mappings[0].CanonicalUrl!.StartsWith("https://learn.microsoft.com/", StringComparison.Ordinal), "mapped docs must use canonical Learn URLs");
        Assert(result.XmlDocs["excel"].Contains("<summary>", StringComparison.Ordinal) && result.XmlDocs["excel"].Contains("<param name=\"Name\">", StringComparison.Ordinal), "constrained Markdown conversion must retain summary and parameter documentation");

        var ledgerPath = Path.Combine(docsRoot, "package-ledger.json");
        using var ledger = JsonDocument.Parse(File.ReadAllText(ledgerPath));
        var root = ledger.RootElement;
        Assert(root.GetProperty("schemaVersion").GetInt32() == 1, "package ledger schema version must be explicit");
        var packageIds = root.GetProperty("packages").EnumerateArray().Select(x => x.GetString() ?? string.Empty).Order(StringComparer.Ordinal).ToArray();
        var expectedDigest = Hex(SHA256.HashData(Encoding.UTF8.GetBytes(string.Join("\n", packageIds) + "\n")));
        Assert(string.Equals(root.GetProperty("digest").GetString(), expectedDigest, StringComparison.Ordinal), "package ledger digest must cover the complete sorted package set");

        var notice = File.ReadAllText(Path.Combine(docsRoot, "THIRD-PARTY-NOTICES.md"));
        foreach (var required in new[] { "Microsoft Corporation", "MicrosoftDocs/VBA-Docs", "CC BY 4.0", "0123456789abcdef0123456789abcdef01234567", "constrained C# XML documentation", "do not endorse or warrant" })
            Assert(notice.Contains(required, StringComparison.OrdinalIgnoreCase), "third-party notice is missing required attribution text: " + required);
    }


    private static void VerifyEndToEnd(string fixtureRoot)
    {
        var root = NewTempRoot("e2e");
        try
        {
            var data = Path.Combine(root, "data");
            var source = Path.Combine(root, "source");
            var output = Path.Combine(root, "output");
            var second = Path.Combine(root, "second");
            Directory.CreateDirectory(data);
            Directory.CreateDirectory(source);
            File.Copy(Fixture(fixtureRoot, "e2e/data/graph.json"), Path.Combine(data, "graph.json"));
            File.Copy(Fixture(fixtureRoot, "e2e/contract.json"), Path.Combine(root, "contract.json"));
            File.Copy(Fixture(fixtureRoot, "e2e/policy.json"), Path.Combine(root, "policy.json"));
            File.WriteAllText(Path.Combine(source, "source-only-sentinel.txt"), "must never be copied");

            var common = new[] { "--locked", "--no-cache", "--data", data, "--contract", Path.Combine(root, "contract.json"), "--policy", Path.Combine(root, "policy.json"), "--source", source };
            var first = RunCli("generate", common.Concat(new[] { "--output", output, "--report", Path.Combine(root, "generate.json") }));
            Assert(first.ExitCode == 0, "the real CLI generate command must succeed: " + first.Stderr);
            var generatedFiles = Directory.EnumerateFiles(output, "Widget.cs", SearchOption.AllDirectories).ToArray();
            Assert(generatedFiles.Length == 1, "projection must emit exactly one expected type file");
            var generated = generatedFiles[0];
            var generatedRelative = Path.GetRelativePath(output, generated).Replace('\\', '/');
            using var generationReport = JsonDocument.Parse(File.ReadAllText(Path.Combine(root, "generate.json")));
            var generationRoot = generationReport.RootElement;
            var emittedPaths = generationRoot.GetProperty("emittedPaths").EnumerateArray().Select(item => item.GetString() ?? "").ToArray();
            Assert(generationRoot.GetProperty("schemaVersion").GetString() == "codegen-report-v3", "CLI output report must use codegen-report-v3");
            Assert(emittedPaths.SequenceEqual(new[] { generatedRelative }, StringComparer.Ordinal), "codegen-report-v3 emittedPaths must exactly describe the projected output");
            Assert(!RelativeFiles(output).Any(path => path.EndsWith("source-only-sentinel.txt", StringComparison.Ordinal)), "generate must not copy source-only files");
            Assert(RelativeFiles(output).SequenceEqual(emittedPaths, StringComparer.Ordinal), "only codegen-report-v3 emitted paths may be emitted");
            var bytes = File.ReadAllBytes(generated);
            var text = Encoding.UTF8.GetString(bytes);
            Assert(bytes.AsSpan().StartsWith(new byte[] { 0xef, 0xbb, 0xbf }), "generated output must have a UTF-8 BOM");
            Assert(text.Contains(OwnershipMarker.D9, StringComparison.Ordinal), "generated output must contain D9 ownership");
            Assert(text.EndsWith("\r\n", StringComparison.Ordinal) && !text.Replace("\r\n", "", StringComparison.Ordinal).Contains('\n'), "generated output must use CRLF and a final newline");
            Assert(text.Contains("namespace Fixture\r\n{", StringComparison.Ordinal), "generated output must validate as C# 7.3 block-namespace syntax");

            var secondRun = RunCli("generate", common.Concat(new[] { "--output", second, "--report", Path.Combine(root, "second.json") }));
            Assert(secondRun.ExitCode == 0, "the repeated real CLI generation must succeed");
            Assert(TreeDigest(output) == TreeDigest(second), "repeated locked generation must be byte deterministic");

            var beforeCheck = TreeDigest(output);
            var cleanCheck = RunCli("generate", common.Concat(new[] { "--output", output, "--check", "--report", Path.Combine(root, "check.json") }));
            Assert(cleanCheck.ExitCode == 0, "a clean --check must exit zero");
            Assert(TreeDigest(output) == beforeCheck, "--check must not change output bytes");

            var expected = Path.Combine(root, "expected");
            CopyTree(output, expected);
            var generatedDirectory = Path.GetDirectoryName(generated)!;
            var expectedGenerated = Path.Combine(expected, generatedRelative);
            File.WriteAllBytes(generated, File.ReadAllBytes(generated).Concat(Encoding.UTF8.GetBytes("// drift\r\n")).ToArray());
            File.WriteAllText(Path.Combine(generatedDirectory, "Added.cs"), "added\r\n");
            File.Delete(expectedGenerated);
            File.WriteAllText(Path.Combine(Path.GetDirectoryName(expectedGenerated)!, "Deleted.cs"), "deleted\r\n");
            var diff = RunCli("diff", new[] { "--expected", expected, "--actual", output, "--report", Path.Combine(root, "diff.json") });
            var diffOutput = diff.Stdout + diff.Stderr;
            Assert(diff.ExitCode == 2 && diffOutput.Contains("Widget.cs", StringComparison.Ordinal) && diffOutput.Contains("Added.cs", StringComparison.Ordinal) && diffOutput.Contains("Deleted.cs", StringComparison.Ordinal), "diff must report changed, added, and deleted paths with exit 2");

            var invalid = Path.Combine(root, "invalid");
            CopyTree(data, invalid);
            File.WriteAllText(Path.Combine(invalid, "graph.json"), File.ReadAllText(Path.Combine(invalid, "graph.json")).Replace("\"2.0\"", "\"2.invalid\"", StringComparison.Ordinal));
            var invalidResult = RunCli("generate", common.Concat(new[] { "--data", invalid, "--output", Path.Combine(root, "invalid-output") }));
            Assert(invalidResult.ExitCode == 1, "an invalid graph must exit 1");
            var inputDigest = Hex(SHA256.HashData(Encoding.UTF8.GetBytes(string.Join("\n", new[] { FileDigest(Path.Combine(data, "graph.json")), FileDigest(Path.Combine(root, "contract.json")), FileDigest(Path.Combine(root, "policy.json")) }) + "\n")));
            var outputDigest = TreeDigest(second);
            EndToEndEvidence = new { command = $"dotnet {CliDisplayPath()} generate --locked --no-cache --data <data> --contract <contract> --policy <policy> --source <source> --output <output>", inputDigest, outputDigest, outputReportSchema = "codegen-report-v3", csharpValidation = "CSharp7_3-block-namespace", firstDivergentStage = "none" };
            Console.WriteLine($"e2e command=dotnet {CliDisplayPath()} generate --locked --no-cache --data <data> --contract <contract> --policy <policy> --source <source> --output <output>; inputDigest={inputDigest}; outputDigest={outputDigest}; csharp=7.3; firstDivergentStage=none");
        }
        finally { DeleteTempRoot(root); }
    }

    private static CliResult RunCli(string command, IEnumerable<string> arguments)
    {
        var cli = Path.Combine(AppContext.BaseDirectory, "NetOffice.CodeGen.Cli.dll");
        Assert(File.Exists(cli), "verification must invoke the built CLI application: " + cli);
        var info = new ProcessStartInfo("dotnet") { RedirectStandardOutput = true, RedirectStandardError = true, UseShellExecute = false };
        info.ArgumentList.Add(cli);
        info.ArgumentList.Add(command);
        foreach (var argument in arguments) info.ArgumentList.Add(argument);
        using var process = Process.Start(info) ?? throw new InvalidOperationException("unable to start dotnet CLI");
        var stdout = process.StandardOutput.ReadToEnd();
        var stderr = process.StandardError.ReadToEnd();
        process.WaitForExit();
        return new CliResult(process.ExitCode, stdout, stderr);
    }

    private static string CliDisplayPath() => Path.Combine(AppContext.BaseDirectory, "NetOffice.CodeGen.Cli.dll");

    private static string[] RelativeFiles(string root) => Directory.EnumerateFiles(root, "*", SearchOption.AllDirectories)
        .Where(x => !x.Contains(Path.DirectorySeparatorChar + ".codegen" + Path.DirectorySeparatorChar, StringComparison.Ordinal))
        .Select(x => Path.GetRelativePath(root, x).Replace(Path.DirectorySeparatorChar, '/'))
        .Order(StringComparer.Ordinal).ToArray();

    private static string TreeDigest(string root) => Hex(SHA256.HashData(Encoding.UTF8.GetBytes(string.Join("\n", RelativeFiles(root).Select(path => path + ":" + FileDigest(Path.Combine(root, path.Replace('/', Path.DirectorySeparatorChar))))) + "\n")));


    private static void CopyTree(string source, string destination)
    {
        foreach (var file in Directory.EnumerateFiles(source, "*", SearchOption.AllDirectories))
        {
            var relative = Path.GetRelativePath(source, file);
            var target = Path.Combine(destination, relative);
            Directory.CreateDirectory(Path.GetDirectoryName(target)!);
            File.Copy(file, target);
        }
    }
    private static string FileDigest(string path) => Hex(SHA256.HashData(File.ReadAllBytes(path)));

    private sealed record CliResult(int ExitCode, string Stdout, string Stderr);
    private static void VerifyRollback()
    {
        var root = NewTempRoot("rollback");
        try
        {
            var initialFile = EmitFixture("Initial");
            GenerationPlanBuilder.Build(root, new[] { initialFile }, null, new PlanOptions { GeneratorVersion = "verification" }).Apply();
            var manifestPath = Path.Combine(root, ".codegen", "ownership.json");
            var managedPath = Path.Combine(root, initialFile.RelativePath.Replace('/', Path.DirectorySeparatorChar));
            var oldBytes = File.ReadAllBytes(managedPath);
            var oldManifest = File.ReadAllBytes(manifestPath);
            var changed = EmitFixture("Changed");
            var plan = GenerationPlanBuilder.Build(root, new[] { changed }, OwnershipManifest.Parse(oldManifest), new PlanOptions
            {
                GeneratorVersion = "verification",
                BeforeOperation = (index, _) => { if (index == 1) throw new InvalidOperationException("verification fault"); }
            });
            var failure = AssertThrows<InvalidOperationException>(plan.Apply, "fault injection must fail the transaction");
            Assert(failure.Message == "verification fault", "fault injection must preserve the injected failure");
            Assert(File.ReadAllBytes(managedPath).SequenceEqual(oldBytes), "rollback must restore the prior generated bytes");
            Assert(File.ReadAllBytes(manifestPath).SequenceEqual(oldManifest), "rollback must restore the prior manifest bytes");
        }
        finally { DeleteTempRoot(root); }
    }

    private static EmittedFile EmitFixture(string value)
        => new CSharpEmitter().Emit(new WrapperFile
        {
            RelativePath = "Excel/Generated/Stable.cs",
            Namespace = "Fixture",
            Types = new List<WrapperType> { new WrapperType { Name = "Stable", Members = new List<WrapperMember> { new WrapperMember { Name = "Value", Kind = "property", Type = "string", Documentation = new WrapperDocumentation { Summary = value } } } } }
        });

    private static TypeLibObservation ReadObservation(string path) => TypeLibObservationSerializer.Deserialize(File.ReadAllBytes(path));

    private static TypeLibObservation[] ReadObservationArray(string path)
    {
        using var document = JsonDocument.Parse(File.ReadAllText(path));
        return document.RootElement.EnumerateArray().Select(item => TypeLibObservationSerializer.Deserialize(Encoding.UTF8.GetBytes(item.GetRawText()))).ToArray();
    }

    private static string Fixture(string root, string relative) => Path.Combine(root, relative.Replace('/', Path.DirectorySeparatorChar));

    private static string NewTempRoot(string name)
    {
        var root = Path.Combine(Path.GetTempPath(), "netoffice-codegen-verification-" + name + "-" + Guid.NewGuid().ToString("N"));
        Directory.CreateDirectory(root);
        return root;
    }

    private static void DeleteTempRoot(string root)
    {
        if (Directory.Exists(root)) Directory.Delete(root, true);
    }

    private static string Hex(byte[] value) => Convert.ToHexString(value).ToLowerInvariant();

    private static void Assert(bool condition, string message)
    {
        if (!condition) throw new InvalidOperationException(message);
    }

    private static T AssertThrows<T>(Action action, string message) where T : Exception
    {
        try
        {
            action();
        }
        catch (T error)
        {
            return error;
        }
        catch (Exception error)
        {
            throw new InvalidOperationException(message + $" (unexpected {error.GetType().Name})", error);
        }
        throw new InvalidOperationException(message);
    }
}

internal static class VerificationEngine
{
    public static VerificationReport Run(IEnumerable<(string Id, string Stage, Action Run)> checks)
    {
        var results = new List<VerificationCheck>();
        foreach (var check in checks)
        {
            try
            {
                check.Run();
                results.Add(new VerificationCheck { Id = check.Id, Stage = check.Stage, Passed = true, Message = "ok" });
            }
            catch (Exception error)
            {
                results.Add(new VerificationCheck { Id = check.Id, Stage = check.Stage, Passed = false, Message = error.Message });
            }
        }
        var first = results.FirstOrDefault(x => !x.Passed)?.Stage ?? "none";
        return new VerificationReport { Outcome = results.All(x => x.Passed) ? "passed" : "failed", FirstDivergentStage = first, Checks = results };
    }
}

internal sealed class VerificationReport
{
    [JsonPropertyName("schemaVersion")] public string SchemaVersion { get; init; } = "netoffice-codegen-verification/v2";
    [JsonPropertyName("outcome")] public string Outcome { get; init; } = "failed";
    [JsonPropertyName("firstDivergentStage")] public string FirstDivergentStage { get; init; } = "none";
    [JsonPropertyName("firstDivergentOwner")] public string FirstDivergentOwner => Program.FullCorpusEvidence?.FirstDivergentOwner ?? "none";
    [JsonPropertyName("checks")] public IReadOnlyList<VerificationCheck> Checks { get; init; } = Array.Empty<VerificationCheck>();
    [JsonPropertyName("evidence")] public object? Evidence => Program.EndToEndEvidence;
    [JsonPropertyName("fullCorpus")] public FullCorpusReport? FullCorpus => Program.FullCorpusEvidence;
    [JsonIgnore] public bool Passed => Outcome == "passed";
    public string ToJson() => JsonSerializer.Serialize(this, new JsonSerializerOptions { WriteIndented = true, PropertyNamingPolicy = JsonNamingPolicy.CamelCase, DictionaryKeyPolicy = JsonNamingPolicy.CamelCase, DefaultIgnoreCondition = JsonIgnoreCondition.WhenWritingNull });
}

internal sealed class VerificationCheck
{
    [JsonPropertyName("id")] public string Id { get; init; } = "";
    [JsonPropertyName("stage")] public string Stage { get; init; } = "";
    [JsonPropertyName("passed")] public bool Passed { get; init; }
    [JsonPropertyName("message")] public string Message { get; init; } = "";
}
