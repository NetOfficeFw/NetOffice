# NetOffice.CodeGen.TypeLib

This net10.0 library imports COM type libraries into immutable, versioned raw observations.

## Offline and live paths

Fixture tests and locked generation inject `INativeTypeLibApi` and an `ITypeLibDescriptorDecoder`; no Office installation or COM activation is required. The live implementation is `WindowsNativeTypeLibApi`, which calls `LoadTypeLibEx` with `RegistryImportKind.None` (`REGKIND_NONE`) and releases the returned COM object in a `finally` block. `ComTypeLibDescriptorDecoder` releases every `ITypeLib`/`ITypeInfo` descriptor through `IComDescriptorReleaser`, including on decoding failures.

## Data contract

`TypeLibObservation` and its nested records are immutable and carry schema `typelib-observation/v1`. `TypeLibObservationSerializer` emits compact canonical UTF-8 JSON: arrays and metadata keys are sorted ordinally and SHA-256 is computed over those bytes. `CapturedAtUtc` is diagnostic provenance and intentionally does not enter canonical bytes.

`TypeLibDependencyGraph.Build` resolves exact versions, reports omitted/ambiguous/unresolved references explicitly, and never invents a dependency. `TypeLibMerger.Merge` is deterministic, rejects unknown schemas and stale compare-and-swap bases, emits conflicts for divergent duplicate identities, and emits `TypeLibRemovalCandidate` records for absent imports rather than deleting anything.

`TypeLibProvenanceValidator` validates source digests and Current Channel metadata. The accepted Current Channel GUID is `492350f6-3a01-4f97-b9c0-c7c6ddf67d60`; build, SKU, locale, and architecture are required when channel metadata is present or required by policy.
