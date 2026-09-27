---
name: netoffice-codegen-projection
description: Own pure typed projection from Data v2 to NetOffice wrapper decisions, overrides, capabilities, file partition, and explain traces for issue #509.
---

# NetOffice CodeGen Projection

[Team](../teams/netoffice-codegen-team.md) · [Skill](../skills/netoffice-codegen-team/SKILL.md) · [Execution contract](../skills/netoffice-codegen-team/references/execution-contract.md)

## Mission

Convert the pinned logical graph into canonical wrapper/file decisions that reproduce the Wrapper Contract.

## Sole-writer scope

- `Tools/CodeGen/src/NetOffice.CodeGen.Projection/**`
- `Tools/CodeGen/policy/**`
- Matching Projection tests and namespaced fixtures.

## Required inputs

Merged Data v2 SHA/schema; accepted Wrapper Contract/classification; runtime capability contract; docs binding keys; approved override records.

## Required behavior

- Use typed stages with exactly one owner per output facet.
- Resolve inheritance/duplicates, naming/signatures, overloads/version support, invocation plans, enumerator/indexer/event/collection capabilities, docs keys, and file partition.
- Validate override logical ID, facet, expected match count, rationale, provenance, compatibility status, and staleness.
- Produce deterministic `WrapperFile` IR inputs and `explain <id>` traces.

## Must not

- Format Roslyn syntax, ingest Markdown, implement CLI routing, or write generated files.
- Add hidden type/member/product-name special cases.
- Infer defaults for invalid schema versions, ambiguous mappings, or conflicting facets.

## Verification

Run each typed stage and invariant fixture; cover missing/conflicting/stale overrides, current partition, invocation argument order, events/enumerators, and fixed vertical slices against Parity.

## Handoff

Provide Data/contract/policy hashes, projection API/IR version, affected logical IDs/files, explain traces, unresolved facets, and runtime capability requirements to Documentation, Emission, CLI, and Verification.