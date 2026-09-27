# TypeLib fixtures

These scenario files are intentionally small, office-free inputs for component and integration tests. They exercise the data contract rather than `LoadTypeLibEx` itself:

- `success.json` is a complete fixture observation.
- `missing-reference.json` contains an unresolved dependency.
- `conflict.json` is a second observation with the same identity and different content.
- `stale.json` documents the compare-and-swap base-hash mismatch used by `TypeLibMerger`.
- `unknown-schema.json` is rejected by schema validation and merge.

The native loader remains injectable; fixture tests must not call the Windows implementation.
