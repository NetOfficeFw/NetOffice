---
name: netoffice-codegen-runtime
description: Implement the post-cutover NetOffice runtime collection capability required by issue #506 while preserving existing provider and COM ownership contracts.
---

# NetOffice Collection Runtime Capability

[Team](../teams/netoffice-codegen-team.md) · [Skill](../skills/netoffice-codegen-team/SKILL.md) · [Execution contract](../skills/netoffice-codegen-team/references/execution-contract.md)

## Mission

Provide the shared runtime behavior that generated eligible collection wrappers consume for #506.

## Sole-writer scope

- `Source/NetOffice/**` for #506 runtime capability and existing extension behavior.
- Role-owned tests only when they reside inside that project; shared solution/package wiring belongs to the lead and `Source/NetOffice.Tests/**` belongs to Collection Tests.

## Required inputs

Accepted generator capability contract; eligible/count-less provider inventory; existing `IEnumerableProvider<T>`, `ICOMObject.SyncRoot`, proxy ownership, and extension signatures; #506 acceptance contract.

## Required behavior

- Preserve `IEnumerableProvider<T>` source/binary compatibility.
- Keep public native `Count` types unchanged; support Int16 widening only through explicit `ICollection.Count`.
- Add the approved shared `CopyTo` helper with BCL-like validation and correct published/unpublished COM-wrapper lifetimes.
- Add the `ICollection` fast path to the existing parameterless NetOffice Count extension while preserving enumeration/disposal fallback and predicate behavior.

## Must not

- Edit generated wrappers, generator rules, analyzers, integration tests, solution/package files, or public Count signatures.
- Optimize predicate/transformed queries or use reflection/dynamic Count discovery.

## Verification

Use managed provider doubles for fast path, fallback, exceptions, Count-throws predicates, CopyTo validation/partial publication, and source-as-enumerator behavior. Request lead-owned wiring separately.

## Handoff

Provide runtime API/capability version, changed public surface, managed test evidence, generator requirement, shared-wiring request, and remaining COM-test needs to Projection, Analyzer, Collection Tests, and the lead.