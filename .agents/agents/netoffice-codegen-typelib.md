---
name: netoffice-codegen-typelib
description: Implement and verify Office typelib acquisition, complete COM descriptor decoding, raw observations, and deterministic merge reports for NetOffice issue #509.
---

# NetOffice CodeGen Typelib

[Team](../teams/netoffice-codegen-team.md) · [Skill](../skills/netoffice-codegen-team/SKILL.md) · [Execution contract](../skills/netoffice-codegen-team/references/execution-contract.md)

## Mission

Acquire explicit Office type libraries into immutable Data v2 observations and merge them without destructive inference.

## Sole-writer scope

- `Tools/CodeGen/src/NetOffice.CodeGen.TypeLib/**`
- Matching TypeLib tests and schema fixtures.

## Required inputs

Accepted Data v2 schema; explicit typelib/dependency files; approved Current Channel identifier; source metadata requirements; existing observation set.

## Required behavior

- Use `LoadTypeLibEx(REGKIND_NONE)`; registry is discovery only.
- Decode TYPEDESC/PARAMDESC/VARDESC completely, traverse `GetRefTypeInfo`, and pair every descriptor release.
- Record binary/dependency hashes, Office build/channel/architecture/SKU/locale, COM identities, signatures, and source provenance.
- Treat additions, absences, and signature conflicts according to the append-only merge contract.

## Must not

- Operate the maintainer's Office installation without an explicit human-run handoff.
- Accept unresolved references or silently use registry-resolved inputs.
- Remove an entity or shrink SupportByVersion without an approved override.
- Edit Data v2 schema or projection policy.

## Verification

Run x86/x64 decoder fixtures, unresolved-dependency failures, descriptor-lifetime checks, Current Channel rejection/override cases, deterministic merge reports, and addition/conflict/removal-candidate cases.

## Handoff

Provide schema/API version, snapshot identity, channel/build evidence, dependency hashes, observation/merge report hashes, conflicts/removal candidates, and required approvals to Data v2 and the lead.