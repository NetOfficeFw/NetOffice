using System.Text.Json;
using System.Xml.Linq;
using NetOffice.CodeGen.Docs;

var fixtureRoot = Path.Combine(AppContext.BaseDirectory, "fixtures");
var temporaryRoot = Path.Combine(Path.GetTempPath(), "netoffice-docs-tests-" + Guid.NewGuid().ToString("N"));
Directory.CreateDirectory(temporaryRoot);
try
{
    ExactBindingRenamesBaselineParameters();
    AmbiguousBindingFailsAfterWritingReports();
    UnmatchedBindingUsesDeterministicFallback();
    MalformedBaselineRawIsPreserved();
    PinnedVbaProfileRemainsOfflineAndDoesNotRenameParameters();
    CompleteAvailableCorpusBinds();
    Console.WriteLine("Documentation tests passed.");
}
finally
{
    Directory.Delete(temporaryRoot, true);
}

void ExactBindingRenamesBaselineParameters()
{
    var contract = Path.Combine(fixtureRoot, "baseline-exact.json");
    var reports = Path.Combine(temporaryRoot, "exact-reports");
    var targets = new[]
    {
        DocumentationTarget.ForType("type-data-id", "docs:type-data-id", "type-generated-id", "Type", "Fixture"),
        DocumentationTarget.ForType("part-data-id", "docs:part-data-id", "type-generated-id", "Type", "Fixture", "Interfaces/Type.cs"),
        DocumentationTarget.ForType("nested-part-data-id", "docs:nested-part-data-id", "Fixture.Type", "Type", "Fixture.Interfaces", "Interfaces/Nested.Type.cs"),
        DocumentationTarget.ForMember("member-data-id", "docs:member-data-id", "type-generated-id", "run", "Method",
            "public global::System.Int32 run(global::System.String text, global::System.Int32 repeatCount)", new[] { "text", "repeatCount" }, "Type", "Fixture"),
        DocumentationTarget.ForType("split-type-data-id", "docs:split-type", "Fixture.Type", "Type_", "Fixture"),
        DocumentationTarget.ForMember("split-constructor-data-id", "docs:split-constructor", "Fixture.Type", "Type_", "constructor",
            "public Type_() : base() {", Array.Empty<string>(), "Type_", "Fixture")
    };
    var options = new DocsOptions("", Profile: DocsProfile.Baseline, BaselineContractPath: contract, ReportDirectory: reports);
    var first = DocumentationSync.Sync(targets, options);
    var firstReport = File.ReadAllBytes(Path.Combine(reports, "mapping-ledger.json"));
    var second = DocumentationSync.Sync(targets.Reverse(), options);
    Assert(first.Digest == second.Digest, "baseline digest must be independent of target input order");
    Assert(firstReport.SequenceEqual(File.ReadAllBytes(Path.Combine(reports, "mapping-ledger.json"))), "mapping report bytes must be deterministic");
    var session = new BaselineDocumentationSession(options with { BaselineContractPath = fixtureRoot, ReportDirectory = null });
    var batchOne = session.BindBatch(targets.Where((_, index) => index % 2 == 0), contract);
    var batchTwo = session.BindBatch(targets.Where((_, index) => index % 2 != 0).Reverse(), contract);
    var combined = session.Complete();
    Assert(batchOne.Mappings.Count + batchTwo.Mappings.Count == targets.Length, "baseline session did not return every batch mapping");
    Assert(combined.Digest == first.Digest
        && combined.Mappings.Count == first.Mappings.Count
        && combined.Documents.All(document => first.Documents.TryGetValue(document.Key, out var expected)
            && expected.Documentation.RawXml == document.Value.Documentation.RawXml),
        "batched baseline binding did not reproduce the one-shot canonical result");
    var summarySession = new BaselineDocumentationSession(
        options with { BaselineContractPath = fixtureRoot, ReportDirectory = null }, retainDocuments: false);
    summarySession.BindBatch(targets.Reverse(), contract);
    var summary = summarySession.Complete();
    Assert(summary.Digest == first.Digest && summary.Mappings.Count == first.Mappings.Count && summary.Documents.Count == 0,
        "summary-only batched binding did not preserve the canonical digest without retaining documents");
    var typeXml = first.Documents["docs:type-data-id"].Documentation;
    Assert(typeXml.Summary == "Current source summary." && typeXml.Remarks == "Current source remarks.", "type summary and remarks were not retained");
    Assert(first.Documents["docs:part-data-id"].Documentation.Summary == "Primary type part.", "multipart type documentation did not select the source-matching part");
    Assert(first.Documents["docs:nested-part-data-id"].Documentation.Summary == "Nested namespace part.", "namespace-specific multipart type documentation selected the canonical type instead");
    var member = first.Documents["docs:member-data-id"].Documentation;
    Assert(member.Summary == "Runs the fixture." && member.Returns == "The result." && member.Remarks == "Current support remarks.", "complete member XML documentation was not retained");
    Assert(member.Parameters["text"] == "Input text." && member.Parameters["repeatCount"] == "Repeat count.", "baseline parameters were not reconciled by signature position");
    Assert(!member.RawXml.Contains("name=\"input\"", StringComparison.Ordinal) && !member.RawXml.Contains("name=\"count\"", StringComparison.Ordinal), "source parameter names leaked after reconciliation");
    Assert(first.Documents["docs:split-type"].Documentation.Summary == "Hidden split type."
        && first.Documents["docs:split-constructor"].Documentation.Summary == "Hidden stub .ctor",
        "split contract type documentation did not override its primary canonical logical ID");
    Assert(File.Exists(Path.Combine(reports, "mapping-ledger.json"))
        && File.Exists(Path.Combine(reports, "unmatched-report.json"))
        && File.Exists(Path.Combine(reports, "ambiguous-report.json")), "deterministic documentation reports were not written");
}

void AmbiguousBindingFailsAfterWritingReports()
{
    var reports = Path.Combine(temporaryRoot, "ambiguous-reports");
    try
    {
        DocumentationSync.Sync(new[] { "Fixture.Type/Run" }, new DocsOptions("", Profile: DocsProfile.Baseline,
            BaselineContractPath: Path.Combine(fixtureRoot, "baseline-ambiguous.json"), ReportDirectory: reports));
        throw new Exception("ambiguous baseline mapping was accepted");
    }
    catch (DocumentationMappingException error)
    {
        Assert(error.Result.Mappings.Single().Kind == MappingKind.Ambiguous, "ambiguous result did not identify the mapping kind");
        Assert(error.Result.Mappings.Single().Candidates?.Count == 2, "ambiguous result did not retain deterministic candidates");
    }
    using var report = JsonDocument.Parse(File.ReadAllBytes(Path.Combine(reports, "ambiguous-report.json")));
    Assert(report.RootElement.GetProperty("count").GetInt32() == 1, "ambiguous report was not produced before failure");
}

void UnmatchedBindingUsesDeterministicFallback()
{
    var options = new DocsOptions("", Profile: DocsProfile.Baseline, BaselineContractPath: Path.Combine(fixtureRoot, "baseline-unmatched.json"));
    var target = DocumentationTarget.ForMember("missing-data-id", "docs:missing", "Fixture.Type", "Missing", "method", "public void Missing()", Array.Empty<string>());
    var first = DocumentationSync.Sync(new[] { target }, options);
    var second = DocumentationSync.Sync(new[] { target }, options);
    Assert(first.Mappings.Single().Kind == MappingKind.Unmatched, "missing baseline mapping was not recorded");
    Assert(first.Digest == second.Digest && first.XmlDocs["docs:missing"] == second.XmlDocs["docs:missing"], "fallback documentation was not deterministic");
}

void MalformedBaselineRawIsPreserved()
{
    var contract = Path.Combine(fixtureRoot, "baseline-malformed.xml.json");
    var result = DocumentationSync.Sync(
        new[] { DocumentationTarget.ForType("malformed-type", "docs:malformed", "Fixture.Type", "Type") },
        new DocsOptions("", Profile: DocsProfile.Baseline, BaselineContractPath: contract));
    var documentation = result.Documents["docs:malformed"].Documentation;
    Assert(result.Mappings.Single().Kind == MappingKind.Exact, "malformed contract documentation did not retain its exact mapping");
    Assert(documentation.RawXml == "<summary>unterminated" && documentation.ParseStatus == "invalid"
        && documentation.ParseError == "Malformed XML documentation.", "malformed contract Raw XML and parse status were not preserved");
}

void PinnedVbaProfileRemainsOfflineAndDoesNotRenameParameters()
{
    var root = Path.Combine(temporaryRoot, "vba");
    Directory.CreateDirectory(root);
    File.WriteAllText(Path.Combine(root, "pin.json"), "{\"repository\":\"MicrosoftDocs/VBA-Docs\",\"commit\":\"0123456789abcdef0123456789abcdef01234567\",\"license\":\"CC BY 4.0\"}");
    File.WriteAllText(Path.Combine(root, "run.md"), string.Join('\n', new[] { "# Run", "", "Runs the command.", "", "- `value`: The value.", "" }));
    var target = DocumentationTarget.ForMember("data-run", "docs:run", "Fixture.Type", "Run", "method", "public void Run(string value)", new[] { "value" });
    var result = DocumentationSync.Sync(new[] { target }, new DocsOptions(root, Path.Combine(root, "pin.json"), DocsProfile.Vba, Locked: true));
    Assert(result.Pin?.Commit == "0123456789abcdef0123456789abcdef01234567" && result.Mappings.Single().Kind == MappingKind.Exact, "VBA binding did not retain its offline pin");
    Assert(result.Documents["docs:run"].Documentation.Parameters["value"] == "The value.", "VBA parameter documentation changed");

    var renamed = target with { EmittedParameterNames = new[] { "renamed" } };
    try
    {
        DocumentationSync.Sync(new[] { renamed }, new DocsOptions(root, Path.Combine(root, "pin.json"), DocsProfile.Vba, Locked: true));
        throw new Exception("VBA parameter name was positionally invented");
    }
    catch (InvalidDataException) { }
}

void CompleteAvailableCorpusBinds()
{
    var repository = FindRepositoryRoot();
    var contractDirectory = Path.Combine(repository, "Tools", "CodeGen", "contracts", "wrapper");
    var daoContract = Path.Combine(contractDirectory, "DAO.wrapper-contract.json");
    var corpus = BaselineContractCorpus.Load(contractDirectory);
    var targets = corpus.CreateValidationTargets();
    var reports = Path.Combine(temporaryRoot, "corpus-reports");
    var result = DocumentationSync.Sync(targets, new DocsOptions("", Profile: DocsProfile.Baseline, BaselineContractPath: contractDirectory, ReportDirectory: reports));
    Assert(targets.Count == corpus.TypeCount + corpus.PartCount + corpus.MemberCount, "corpus validation did not enumerate every type, part, and member");
    Assert(result.Mappings.Count == targets.Count && result.Mappings.All(static x => x.Kind is MappingKind.Exact or MappingKind.Unmatched), "corpus baseline binding did not process every available declaration");
    var byBinding = targets.ToDictionary(static x => x.BindingKey, StringComparer.Ordinal);
    foreach (var bound in result.Documents.Values)
    {
        var target = byBinding[bound.BindingKey];
        if (target.Kind == DocumentationTargetKind.Type || string.Equals(bound.Documentation.ParseStatus, "invalid", StringComparison.OrdinalIgnoreCase)) continue;
        var emitted = target.EmittedParameterNames?.ToHashSet(StringComparer.Ordinal) ?? new HashSet<string>(StringComparer.Ordinal);
        Assert(bound.Documentation.Parameters.Keys.All(emitted.Contains), $"'{target.LogicalId}' emitted XML for an absent parameter");
    }

    using var dao = JsonDocument.Parse(File.ReadAllBytes(daoContract));
    var database = dao.RootElement.GetProperty("Types").EnumerateArray().Single(x => x.GetProperty("LogicalId").GetString() == "NetOffice.DAOApi.Database");
    var databaseTypeTarget = DocumentationTarget.ForType("type-dao-database", "docs:dao:database", "NetOffice.DAOApi.Database", "Database");
    var execute = database.GetProperty("Members").EnumerateArray().Single(x => x.GetProperty("Signature").GetString() == "public void Execute(string query, object options) {");
    var executeTarget = DocumentationTarget.ForMember("member-dao-database-execute", "docs:dao:database:execute", "NetOffice.DAOApi.Database", "Execute", "method",
        execute.GetProperty("Signature").GetString()!, new[] { "commandText", "executeOptions" });
    var representative = DocumentationSync.Sync(new[] { databaseTypeTarget, executeTarget }, new DocsOptions("", Profile: DocsProfile.Baseline, BaselineContractPath: daoContract));
    Assert(representative.Documents[databaseTypeTarget.BindingKey].Documentation.Summary == "DispatchInterface Database SupportByVersion DAO, 3.6,12.0", "DAO Database summary was not retained");
    var executeDocs = representative.Documents[executeTarget.BindingKey].Documentation;
    Assert(executeDocs.Summary == "SupportByVersion DAO 3.6, 12.0", "DAO Database method support summary was not retained");
    Assert(executeDocs.Parameters["commandText"] == "string query" && executeDocs.Parameters["executeOptions"] == "optional object options", "DAO Database parameter text was not retained while names were reconciled");
    Console.WriteLine($"Baseline corpus: {corpus.TypeCount} types, {corpus.PartCount} parts, {corpus.MemberCount} members, {result.Mappings.Count(static x => x.Kind == MappingKind.Exact)} mapped, {result.Mappings.Count(static x => x.Kind == MappingKind.Unmatched)} unmatched, digest {result.Digest}.");
}

string FindRepositoryRoot()
{
    var current = new DirectoryInfo(Directory.GetCurrentDirectory());
    while (current is not null)
    {
        if (File.Exists(Path.Combine(current.FullName, "Tools", "CodeGen", "contracts", "wrapper", "DAO.wrapper-contract.json"))) return current.FullName;
        current = current.Parent;
    }
    throw new DirectoryNotFoundException("Unable to locate the NetOffice repository root.");
}

static void Assert(bool condition, string message)
{
    if (!condition) throw new Exception(message);
}
