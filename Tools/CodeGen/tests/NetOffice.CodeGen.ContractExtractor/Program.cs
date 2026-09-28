// Copyright (c) 2026 NetOffice contributors
// SPDX-License-Identifier: MIT

using NetOffice.CodeGen.ContractExtractor;

namespace NetOffice.CodeGen.ContractExtractor.Tests;

internal static class Program
{
    public static int Main()
    {
        var fixtureRoot = Path.Combine(AppContext.BaseDirectory, "fixtures", "Source");
        var fixtureApiRoot = Path.Combine(fixtureRoot, "Fixture");
        var first = Extractor.Extract(fixtureRoot, "Fixture", fixtureApiRoot);
        var second = Extractor.Extract(fixtureRoot, "Fixture", fixtureApiRoot);
        AssertEqual(JsonFile.Serialize(first), JsonFile.Serialize(second), "Extraction must be byte deterministic.");

        var generated = Single(first.Types, "Fixture.Generated");
        AssertSequence(new[] { "BaseGenerated" }, new[] { generated.BaseType }, "Base type was not extracted.");
        AssertSequence(new[] { "IFoo" }, generated.Interfaces, "Interface was not extracted.");
        AssertSequence(new[] { "\"Fixture\", 1, 2" }, generated.SupportVersions, "Type support versions were not extracted.");
        AssertEqual("Generated wrapper.", generated.Documentation.Summary, "Baseline type documentation was not extracted.");

        var invoke = generated.Members.Single(member => member.Name == "Invoke");
        AssertEqual("4", invoke.DefaultValues["value"], "Parameter default was not extracted.");
        Assert(invoke.InvocationText.Any(text => text.Contains("ExecuteStringMethodGet", StringComparison.Ordinal)), "Invocation facet was not extracted.");
        AssertSequence(new[] { "\"Fixture\", 2" }, invoke.SupportVersions, "Member support versions were not extracted.");
        AssertEqual("The value.", invoke.Documentation.Parameters["value"], "Parameter documentation was not extracted.");
        AssertSequence(new[] { "value" }, invoke.Documentation.Parameters.Keys, "Private-member documentation leaked to a public member.");

        var partial = Single(first.Types, "Fixture.PartialWrapper");
        AssertEqual(2, partial.Parts.Count, "Partial declarations were not merged into one contract type.");
        AssertSequence(new[] { "One", "Two" }, partial.Members.Select(member => member.Name), "Partial members were not merged deterministically.");
        AssertSequence(new[] { "SupportByVersion(\"Fixture\", 1)" }, partial.Attributes, "Multiline attribute metadata was not extracted.");
        AssertSequence(
            new[] { "First partial declaration.", "Second partial declaration." },
            partial.Parts.Select(part => part.Documentation.Summary),
            "Partial declaration documentation was not preserved.");
        AssertEqual(0, first.Unknowns.Count, "Projectable contract must not contain unknown records.");
        AssertEqual(0, first.Ambiguities.Count, "Projectable contract must not contain ambiguity records.");
        Assert(first.Types.Any(type => type.Kind == "delegate" && type.Name == "GeneratedEventHandler"), "Delegate identity was not extracted.");

        var classification = Records.CreateClassification(first);
        AssertEqual("wrapper-generated", classification.Files.Single(file => file.Path == "DispatchInterfaces/Generated.cs").Ownership, "Generated file ownership is wrong.");
        AssertEqual("companion", classification.Files.Single(file => file.Path == "Tools/Companion.cs").Ownership, "Companion file ownership is wrong.");
        AssertEqual("runtime", classification.Files.Single(file => file.Path == "Native/Runtime.cs").Ownership, "Runtime file ownership is wrong.");
        AssertEqual("manual", classification.Files.Single(file => file.Path == "DeclarationKinds.cs").Ownership, "Manual file ownership is wrong.");

        var temp = Path.Combine(Path.GetTempPath(), "netoffice-contract-extractor-tests-" + Guid.NewGuid().ToString("N"));
        try
        {
            var contracts = Path.Combine(temp, "contracts");
            var actual = Path.Combine(temp, "actual");
            Directory.CreateDirectory(contracts);
            CopyDirectory(fixtureRoot, actual);
            JsonFile.Write(Path.Combine(contracts, "Fixture.wrapper-contract.json"), first);
            JsonFile.Write(Path.Combine(contracts, "Fixture.classification.json"), classification);
            var ledger = Records.CreateLedger(first);
            var ledgerPath = Path.Combine(contracts, "Fixture.compatibility-ledger.json");
            JsonFile.Write(ledgerPath, ledger);

            var clean = SemanticComparer.Compare(contracts, actual, new[] { "Fixture" });
            AssertEqual(0, clean.Summary.UnexplainedDifferences, "Identical source must compare cleanly.");

            var declarationPath = Path.Combine(actual, "Fixture", "DispatchInterfaces", "Generated.cs");
            var declarations = File.ReadAllText(declarationPath);
            var qualifiedConstructor = declarations.Replace(
                "[System.ComponentModel.EditorBrowsable(System.ComponentModel.EditorBrowsableState.Never)]",
                "[EditorBrowsable(EditorBrowsableState.Never)]",
                StringComparison.Ordinal);
            File.WriteAllText(declarationPath, qualifiedConstructor);
            var constructorQualification = SemanticComparer.Compare(contracts, actual, new[] { "Fixture" });
            AssertEqual(0, constructorQualification.Summary.UnexplainedDifferences,
                "Constructor attribute namespace qualification must compare cleanly.");

            var customEvent = qualifiedConstructor.Replace(
                "public event System.EventHandler Changed;",
                "public event System.EventHandler Changed\n        {\n            add { Attach(value); }\n            remove { Detach(value); }\n        }",
                StringComparison.Ordinal);
            File.WriteAllText(declarationPath, customEvent);
            Assert(customEvent.Contains("Attach(value)", StringComparison.Ordinal),
                "Custom event fixture mutation was not applied.");
            var eventContract = Extractor.Extract(actual, "Fixture", Path.Combine(actual, "Fixture"));
            var changedEvent = eventContract.Types.SelectMany(type => type.Members)
                .Single(member => member.LogicalId == "Fixture.Generated::event:Changed");
            Assert(changedEvent.InvocationText.Count > 0,
                "Custom event accessors were not retained in the invocation facet.");
            var eventShape = SemanticComparer.Compare(contracts, actual, new[] { "Fixture" });
            Assert(!eventShape.Products[0].Differences.Any(difference =>
                    difference.LogicalId.EndsWith("::event:Changed", StringComparison.Ordinal) &&
                    difference.Facet == "signature"),
                "Field-like and custom-accessor event declarations must have the same public signature.");
            Assert(eventShape.Products[0].Differences.Any(difference =>
                    difference.LogicalId.EndsWith("::event:Changed", StringComparison.Ordinal) &&
                    difference.Facet == "invocation"),
                "Custom event accessor behavior must remain an invocation difference: " +
                string.Join("; ", eventShape.Products[0].Differences.Select(difference =>
                    difference.LogicalId + "/" + difference.Facet + " expected=" + difference.Expected +
                    " actual=" + difference.Actual)));

            File.WriteAllText(declarationPath, qualifiedConstructor.Replace(
                "EditorBrowsableState.Never", "EditorBrowsableState.Always", StringComparison.Ordinal));
            var constructorAttributeArgument = SemanticComparer.Compare(contracts, actual, new[] { "Fixture" });
            Assert(constructorAttributeArgument.Products[0].Differences.Any(difference =>
                    difference.LogicalId.EndsWith("::constructor:Generated()", StringComparison.Ordinal) &&
                    difference.Facet == "attributes"),
                "Constructor attribute arguments must remain semantic differences.");

            File.WriteAllText(declarationPath, qualifiedConstructor.Replace(
                "        [EditorBrowsable(EditorBrowsableState.Never)]\n", string.Empty, StringComparison.Ordinal));
            var missingConstructorAttribute = SemanticComparer.Compare(contracts, actual, new[] { "Fixture" });
            Assert(missingConstructorAttribute.Products[0].Differences.Any(difference =>
                    difference.LogicalId.EndsWith("::constructor:Generated()", StringComparison.Ordinal) &&
                    difference.Facet == "attributes"),
                "Constructor attribute presence must remain a semantic difference.");
            File.WriteAllText(declarationPath, declarations);

            var changedPath = Path.Combine(actual, "Fixture", "DispatchInterfaces", "Generated.cs");
            var equivalent = File.ReadAllText(changedPath)
                .Replace("[SupportByVersion(", "[global::NetOffice.Attributes.SupportByVersionAttribute(", StringComparison.Ordinal)
                .Replace("public string Invoke(int value = 4)", "public global::System.String Invoke(global::System.Int32 value = 4)", StringComparison.Ordinal)
                .Replace("new object[]{ value }", "new object[] { value }", StringComparison.Ordinal)
                .Replace("ExecuteStringMethodGet(", "ExecuteStringMethodGet (", StringComparison.Ordinal)
                .Replace("if(false==Factory.Settings.Enabled)", "if ( false == Factory.Settings.Enabled )", StringComparison.Ordinal)
                .Replace("});;", "});", StringComparison.Ordinal)
                .Replace("Value = 128", "Value = 128,", StringComparison.Ordinal)
                .Replace("/// <summary>Generated wrapper.</summary>", "/// <summary>\n    /// Generated wrapper.\n    /// </summary>", StringComparison.Ordinal);
            File.WriteAllText(changedPath, equivalent);
            var equivalentResult = SemanticComparer.Compare(contracts, actual, new[] { "Fixture" });
            AssertEqual(0, equivalentResult.Summary.UnexplainedDifferences, "Equivalent C# token trivia, enum comma, qualification, aliases, attributes, and XML trivia must compare cleanly: " + string.Join("; ", equivalentResult.Products[0].Differences.Select(difference => difference.LogicalId + "/" + difference.Facet + " expected=" + difference.Expected + " actual=" + difference.Actual)));

            File.WriteAllText(changedPath, equivalent.Replace("\"Invoke\"", "\"InvokeChanged\"", StringComparison.Ordinal));
            var invocationDrift = SemanticComparer.Compare(contracts, actual, new[] { "Fixture" });
            Assert(invocationDrift.Products[0].Differences.Any(difference => difference.Facet == "invocation"), "Invocation literals must remain semantic differences.");

            File.WriteAllText(changedPath, equivalent.Replace("new object[] { value }", "new object[] { Convert(value) }", StringComparison.Ordinal));
            var conversionDrift = SemanticComparer.Compare(contracts, actual, new[] { "Fixture" });
            Assert(conversionDrift.Products[0].Differences.Any(difference => difference.Facet == "invocation"), "Invocation conversions must remain semantic differences.");

            File.WriteAllText(changedPath, equivalent.Replace("this, \"Invoke\"", "\"Invoke\", this", StringComparison.Ordinal));
            var argumentOrderDrift = SemanticComparer.Compare(contracts, actual, new[] { "Fixture" });
            Assert(argumentOrderDrift.Products[0].Differences.Any(difference => difference.Facet == "invocation"), "Invocation argument order must remain a semantic difference.");

            File.WriteAllText(changedPath, equivalent.Replace("if ( false == Factory.Settings.Enabled ) return null;", "if ( false == Factory.Settings.Enabled )", StringComparison.Ordinal));
            var statementDrift = SemanticComparer.Compare(contracts, actual, new[] { "Fixture" });
            Assert(statementDrift.Products[0].Differences.Any(difference => difference.Facet == "invocation"), "Added or removed invocation statements must remain semantic differences.");

            File.WriteAllText(changedPath, equivalent.Replace("for(;;)", "for(;)", StringComparison.Ordinal));
            var forDelimiterDrift = SemanticComparer.Compare(contracts, actual, new[] { "Fixture" });
            Assert(forDelimiterDrift.Products[0].Differences.Any(difference => difference.Facet == "invocation"), "Required for-loop semicolon delimiters must remain distinct tokens.");

            File.WriteAllText(changedPath, equivalent.Replace("ExecuteStringMethodGet", "ExecuteObjectMethodGet", StringComparison.Ordinal));
            var helperDrift = SemanticComparer.Compare(contracts, actual, new[] { "Fixture" });
            Assert(helperDrift.Products[0].Differences.Any(difference => difference.Facet == "invocation"), "Invocation helper changes must remain semantic differences.");

            File.WriteAllText(changedPath, equivalent.Replace("The value.", "Changed value documentation.", StringComparison.Ordinal));
            var documentationDrift = SemanticComparer.Compare(contracts, actual, new[] { "Fixture" });
            Assert(documentationDrift.Products[0].Differences.Any(difference => difference.Facet == "docs"), "Documentation text must remain a semantic difference.");

            File.WriteAllText(changedPath, equivalent.Replace("Int32 value = 4", "Int32 renamed = 4", StringComparison.Ordinal));
            var parameterDrift = SemanticComparer.Compare(contracts, actual, new[] { "Fixture" });
            Assert(parameterDrift.Products[0].Differences.Any(difference => difference.Facet == "member"), "Parameter names must remain part of member identity.");

            File.WriteAllText(changedPath, equivalent);
            var movedPath = Path.Combine(Path.GetDirectoryName(changedPath)!, "GeneratedMoved.cs");
            File.Move(changedPath, movedPath);
            var partitionDrift = SemanticComparer.Compare(contracts, actual, new[] { "Fixture" });
            Assert(partitionDrift.Products[0].Differences.Any(difference => difference.Facet == "partition"), "File partitions must remain semantic differences.");
            File.Move(movedPath, changedPath);

            var changed = equivalent.Replace("\"Fixture\", 1, 2", "\"Fixture\", 1", StringComparison.Ordinal);
            File.WriteAllText(changedPath, changed);
            var drift = SemanticComparer.Compare(contracts, actual, new[] { "Fixture" });
            Assert(drift.Summary.UnexplainedDifferences > 0, "Semantic drift must be reported.");
            AssertEqual("Fixture.Generated", drift.Products[0].FirstDivergentLogicalIds[0], "First divergent logical ID is wrong.");
            Assert(drift.Products[0].Differences.Any(difference => difference.Facet == "supportVersions"), "Support-version drift facet is missing.");
            AssertEqual(2, drift.Summary.ByDifferenceKind["changed"], "Changed-difference aggregate is wrong.");
            AssertEqual(2, drift.Summary.ByRecordKind["type"], "Type-difference aggregate is wrong.");

            ledger.Entries.Add(new LedgerEntry
            {
                LogicalId = "Fixture.Generated",
                RecordKind = "type",
                Source = "DispatchInterfaces/Generated.cs",
                Status = "intentional",
                ExpectedMatchCount = 1,
                Facet = "supportVersions",
                Rationale = "Approved fixture difference.",
                Provenance = "fixture:test",
                ApprovedBy = "test"
            });
            JsonFile.Write(ledgerPath, ledger);
            var approved = SemanticComparer.Compare(contracts, actual, new[] { "Fixture" });
            AssertEqual(1, approved.Summary.IntentionalDifferences, "Intentional ledger difference must remain explicit.");
            var approvedDifference = approved.Products[0].Differences.Single(difference => difference.Facet == "supportVersions");
            Assert(approvedDifference.Intentional, "Intentional ledger match was not marked.");
            AssertEqual("Approved fixture difference.", approvedDifference.LedgerRationale, "Intentional rationale was not preserved.");
        }
        finally
        {
            if (Directory.Exists(temp))
                Directory.Delete(temp, recursive: true);
        }

        Console.WriteLine("Contract extractor behavioral tests passed.");
        return 0;
    }

    private static TypeRecord Single(IEnumerable<TypeRecord> records, string logicalId)
    {
        var matches = records.Where(record => record.LogicalId == logicalId).ToList();
        if (matches.Count != 1)
            throw new InvalidOperationException("Expected one " + logicalId + "; found: " + string.Join(", ", records.Select(record => record.LogicalId)));
        return matches[0];
    }
    private static void CopyDirectory(string source, string destination)
    {
        foreach (var directory in Directory.EnumerateDirectories(source, "*", SearchOption.AllDirectories))
            Directory.CreateDirectory(Path.Combine(destination, Path.GetRelativePath(source, directory)));
        foreach (var file in Directory.EnumerateFiles(source, "*", SearchOption.AllDirectories))
        {
            var target = Path.Combine(destination, Path.GetRelativePath(source, file));
            Directory.CreateDirectory(Path.GetDirectoryName(target)!);
            File.Copy(file, target);
        }
    }

    private static void Assert(bool condition, string message)
    {
        if (!condition)
            throw new InvalidOperationException(message);
    }

    private static void AssertEqual<T>(T expected, T actual, string message)
    {
        if (!EqualityComparer<T>.Default.Equals(expected, actual))
            throw new InvalidOperationException(message + " Expected: " + expected + "; actual: " + actual + ".");
    }

    private static void AssertSequence(IEnumerable<string> expected, IEnumerable<string> actual, string message)
    {
        if (!expected.SequenceEqual(actual, StringComparer.Ordinal))
            throw new InvalidOperationException(message + " Expected: " + string.Join(", ", expected) + "; actual: " + string.Join(", ", actual) + ".");
    }
}
