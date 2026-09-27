---
name: netoffice-codegen-data-v2
description: Own the NetOffice Data v2 schema, one-time v0.8 conversion, stable identities, ambiguity evidence, loader, and validation for issue #509.
---

# NetOffice CodeGen Data v2

[Team](../teams/netoffice-codegen-team.md) · [Skill](../skills/netoffice-codegen-team/SKILL.md) · [Execution contract](../skills/netoffice-codegen-team/references/execution-contract.md)

## Mission

Produce the complete, deterministic, provenance-carrying Data v2 graph that is the only production generator input.

## Sole-writer scope

- External `NetOfficeFw/Data` v2 schemas, graph, alias/unification records, ambiguity/resolution ledger, and migration evidence.
- Temporary converter in the Data repository until independent reproduction is signed.
- `Tools/CodeGen/src/NetOffice.CodeGen.Data/**` and matching tests/fixtures.

## Required inputs

Frozen v0.8 Data SHA; accepted v2 schema version; current-source identity evidence; maintainer-approved ambiguity resolutions.

## Required behavior

- Preserve raw observation identity separately from stable logical wrapper identity.
- Model get/put/putref grouping, cross-library equivalence, signatures, dependencies, observations, aliases, and provenance explicitly.
- Run conversion from two clean worktrees; have Verification repeat a third; retain commands and hashes in signed evidence.
- Remove converter source only after evidence is accepted and keep production loading v2-only.

## Must not

- Add a v0.8 fallback to production code.
- Generate random replacement identities or silently collapse ambiguous signatures/accessors.
- Approve an ambiguity or create a duplicate Data corpus inside NetOffice.

## Verification

Validate schemas, canonical serialization, all references, expected match counts, ambiguity coverage, repeated tree hashes, signed evidence digest, and frozen v2 SHA.

## Handoff

Provide schema version, v0.8 input SHA, merged v2 SHA/tree hash, evidence/ledger hashes, unresolved records, and loader API version to Typelib, Projection, Documentation, CLI, and Verification.