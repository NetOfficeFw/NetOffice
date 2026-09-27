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

The deterministic `ProjectionResult` contains WrapperFile-like records partitioned by policy, inheritance closure, duplicate/overload groups, sanitized C# names/signatures, parsed support versions, invocation plans, collection/event/enumerator/indexer capabilities, runtime requirements, stable documentation binding keys, typed overrides, and explain traces. All collections are sorted by stable IDs or names and JSON uses invariant compact serialization.

Build this project directly with:

```text
dotnet restore NetOffice.CodeGen.Projection.csproj
dotnet build NetOffice.CodeGen.Projection.csproj -c Release --no-restore
```
