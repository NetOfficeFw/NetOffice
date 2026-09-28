# NetOffice.CodeGen.Projection

This standalone `net8.0` project is the pure projection stage. It references only `NetOffice.CodeGen.Data`; it does not load Office, COM type libraries, the network, or source files.

```csharp
var policy = ProjectionPolicy.Read("policy/fixtures/success.json");
var graph = CanonicalJson.Read("data-v2.json");
var contract = WrapperContract.Read("contracts/wrapper/DAO.wrapper-contract.json");
var result = ProjectionEngine.Project(graph, contract, policy, new ProjectionOptions {
    ExpectedDataDigest = graph.Digest,
    ExpectedPolicyDigest = policy.Digest
});
File.WriteAllText("projection.json", result.ToJson());
```

The API validates the pinned Data v2 schema, graph digest, policy schema/digest, and contract schema before doing any work. Unknown schema versions, stale digests, unresolved ambiguity/unknown records, missing references, and stale typed overrides fail with `ProjectionValidationException`; the projector never guesses a replacement.

The deterministic `ProjectionResult` is enumerated exclusively from Data v2. The Wrapper Contract can enrich a matching entity with approved parity metadata but cannot add or hide entities. Projection covers every library/type/member/value; namespace and current file categories, entity kinds, inheritance, optional-parameter overloads, support observations, attributes, invocation evidence and argument order, return conversion, property access direction, collection/event/enumerator/indexer capabilities, constructor plans, runtime requirements, duplicates, and documentation keys are retained as typed records. `Coverage` reports the one-to-one Data inputs, and projection fails rather than silently dropping an unsupported kind or entity. Explain traces identify the rule and evidence used at each stage.

Set `ProjectionOptions.IncludeTraces` to `false` for generation paths that do not serve `explain`; this prevents trace construction rather than allocating and discarding the evidence afterward. The default remains `true` for direct projection and explain consumers.

Build this project directly with:

```text
dotnet restore NetOffice.CodeGen.Projection.csproj
dotnet build NetOffice.CodeGen.Projection.csproj -c Release --no-restore
```

Run the product-neutral representative fixtures with:

```text
dotnet run --project Tests/NetOffice.CodeGen.Projection.Tests.csproj -c Release
```

Passing a canonical full graph and an approved API contract adds the corpus coverage smoke run:

```text
dotnet run --project Tests/NetOffice.CodeGen.Projection.Tests.csproj -c Release -- <graph.json> <wrapper-contract.json>
```
