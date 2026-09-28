# NetOffice.CodeGen.Cli

The CLI composes canonical Data v2, projection policy, Wrapper Contracts, offline documentation binding, typed emission, and ownership-safe storage. Use `--help` for the complete command surface.

`generate` accepts a graph file or a directory containing `graph.json` and requires an explicit projection policy; the canonical production policy is `Tools/CodeGen/policy/projection-policy.json`. `--projects` accepts a deterministic comma-separated product list or `all`; `all` selects every non-ignored NetOffice product and leaves external/ignored typelib dependencies available only for reference resolution. A Wrapper Contract may be a single file or a directory of `*.wrapper-contract.json` files. `--docs-profile baseline` binds current-source XML documentation from those contracts and rejects ambiguous mappings.

Normal generation writes each product beneath `<Product>/Generated/` using the projection's contract-owned partition path. `--isolated-output --source Source` writes beneath `Source/<Product>/Generated/` and copies only explicitly reported build metadata, NetOffice runtime files, and non-generated/manual companions. Existing generated wrapper implementations are excluded. Reports distinguish `emittedPaths` from classified `copiedPaths` and include input hashes, entity counts, the intended output-tree hash, elapsed time, and peak managed memory.

`--from-data`/`--exploratory` is an unlocked data-only mode. Locked generation requires an explicit contract for every selected product. Locked generation, `diff`, and `explain` do not access Office or the network. `generate --check` computes drift without changing output bytes, timestamps, ownership manifest, or cache. Generation is currently cacheless even when `--no-cache` is omitted.

Exit codes are `0` for success/clean, `1` for validation or execution failure, `2` for detected generation/diff drift, and `130` for cancellation. Commands that are not yet composed fail explicitly rather than reporting success.
