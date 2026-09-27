---
name: netoffice-codegen-parity
description: Extract and govern the current-source Wrapper Contract, compatibility ledger, ownership classification, and zero-difference gates for NetOffice issue #509.
---

# NetOffice CodeGen Parity Contract

[Team](../teams/netoffice-codegen-team.md) · [Skill](../skills/netoffice-codegen-team/SKILL.md) · [Execution contract](../skills/netoffice-codegen-team/references/execution-contract.md)

## Mission

Turn the frozen checked-in wrapper source into the versioned compatibility oracle used by Projection and Verification.

## Sole-writer scope

- `Tools/CodeGen/tools/NetOffice.CodeGen.ContractExtractor/**`
- `Tools/CodeGen/contracts/wrapper/**`
- Matching contract-extractor tests and namespaced fixtures.

## Required inputs

Frozen NetOffice commit; all twelve current API source trees; current runtime helper APIs; approved compatibility classifications.

## Required behavior

- Extract public/protected API, attributes, defaults, interface maps, invocation contracts, version support, events, enumerators, docs baseline, and file/type partition.
- Classify every current file/type part as generated, companion, obsolete, or unresolved.
- Record intentional behavior as typed compatibility-ledger entries with logical ID, facet, expected count, rationale, and provenance.
- Review Projection output and report the first unexplained contract difference.

## Must not

- Treat the legacy Ms-PL generator or released assemblies as the oracle.
- Copy legacy code/templates/resources.
- Change an expected contract to make output pass without approved evidence.
- Permit cutover with an unresolved ownership classification.

## Verification

Run extractor determinism and representative method/property/event/inheritance fixtures; compare every slice and the full corpus; preserve machine-readable zero/unexplained-difference reports.

## Handoff

Provide the contract/schema version, frozen source SHA, ownership classification hash, compatibility-ledger hash, affected logical IDs, and acceptance criteria to Projection and Verification.