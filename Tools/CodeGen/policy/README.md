# Projection policy

`projection-policy/v1` contains only product-neutral rules for the pure Projection stage. The policy pins the Data v2 and Wrapper Contract schema versions, file partitioning, C# identifier rules, explicit typed overrides, documentation-key prefix, and runtime capability mappings.

`ProjectionPolicy.ValidateOrThrow` rejects an unknown policy schema, a malformed or stale digest, duplicate override selectors, and unsafe output roots. `ProjectionEngine.Project` validates the DataGraph with the pinned schema and digest before projection; unresolved data/contract ambiguity and unknown records are errors, never inferred defaults.

## Determinism

Projection arrays are sorted by logical ID, path, or canonical name. `ProjectionResult.ToJson()` uses compact UTF-8-independent JSON with invariant ordering. Explain traces record each matching, naming, inheritance, overload, invocation, documentation, and partition rule.

## Overrides

Overrides must identify a type or member and declare `expectedMatches`. A missing selector is a stale override and fails projection. Supported properties are `csharpName`, `namespace`, `signature`, `docsBindingKey`, `invocationOperation`, and `runtimeCapability`. No product or type-name exceptions belong in this file.

The fixtures demonstrate valid policy loading plus missing-reference, conflict, stale-override, and unknown-schema rejection cases. The fixture files intentionally use deterministic digests only for the valid case.
