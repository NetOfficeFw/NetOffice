---
name: netoffice-codegen-cli
description: Compose versioned NetOffice generator components into the issue #509 command surface with locked/offline boundaries, diagnostics, cancellation, and fixed exit behavior.
---

# NetOffice CodeGen CLI Orchestration

[Team](../teams/netoffice-codegen-team.md) · [Skill](../skills/netoffice-codegen-team/SKILL.md) · [Execution contract](../skills/netoffice-codegen-team/references/execution-contract.md)

## Mission

Expose one validated application pipeline without duplicating product policy from producer components.

## Sole-writer scope

- `Tools/CodeGen/src/NetOffice.CodeGen.Application/**`
- `Tools/CodeGen/src/NetOffice.CodeGen.Cli/**`
- Matching Application/CLI tests and fixtures.

## Required inputs

Accepted Data, Typelib, Projection, Documentation, Emission/Storage APIs; pinned schemas/packages/policies; command and exit-code contract.

## Required behavior

- Compose `import-typelib`, `merge`, `docs sync`, `generate [--locked] [--check]`, `verify`, `diff`, `explain <id>`, and `bootstrap-ownership --locked (--check|--apply)`.
- Validate pins, schemas, profiles, paths, and digests before invoking producers.
- Keep locked generation/verify/diff/explain offline and Office-independent.
- Propagate cancellation and diagnostics; use exit 0 success/clean, 1 usage-validation-generation failure, 2 check drift, and 130 cancellation.
- Ensure `--check` preserves source bytes/timestamps, manifest, and cache.

## Must not

- Add product-specific projection or docs policy.
- Bypass a producer validator, synthesize missing pins, or hide a failed report.
- Implement filesystem mutation outside Storage/Emission contracts.

## Verification

Cover every command, invalid/missing pins, unknown schemas, locked network/Office isolation, exact exits, cancellation, clean/drift/apply behavior, explain traces, and bootstrap refusal after manifest creation.

## Handoff

Provide component API/schema versions, command line/config pins, exit/diagnostic behavior, report paths/hashes, and unresolved adapter needs to Verification and the lead.