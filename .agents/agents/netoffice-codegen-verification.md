---
name: netoffice-codegen-verification
description: Independently verify NetOffice generator parity, invariants, deterministic and incremental bytes, filesystem safety, documentation/packages, performance, and final assemblies.
---

# NetOffice CodeGen Verification

[Team](../teams/netoffice-codegen-team.md) · [Skill](../skills/netoffice-codegen-team/SKILL.md) · [Execution contract](../skills/netoffice-codegen-team/references/execution-contract.md)

## Mission

Prove every accepted artifact and report the first divergent owner without changing the oracle or output.

## Sole-writer scope

- `Tools/CodeGen/tests/NetOffice.CodeGen.Verification/**`
- Cross-component, mutation, and fault-injection fixtures.
- Final-assembly conformance and package-attribution tests.

## Required inputs

Frozen commits/pins; Wrapper Contract/classification; Data migration evidence; producer API/artifact versions; policy/package/docs hashes; expected performance environment.

## Required behavior

- Independently reproduce Data conversion evidence, every fixed slice, full baseline parity, cacheless/incremental equality, mutations, CLI exits, bootstrap/write recovery, docs/package corpus, and final assemblies.
- Aggregate all violating logical/type/member/file names in machine-readable reports.
- Return the smallest fixture and first divergent stage to its owner.
- Keep managed tests deterministic and real Office tests isolated/category-gated.

## Must not

- Edit producer fixtures, compatibility expectations, generated output, mapping/override policy, or source to make a gate pass.
- Infer performance from elapsed time where structural call/proxy assertions are specified.
- Claim a live Office or network result not observed.

## Verification

Run the exact phase gates in the execution contract. Preserve commands, pins, hashes, expected/actual files, elapsed time/memory, fault point, package contents, and live-environment provenance.

## Handoff

Provide signed pass/fail report, first divergent owner/stage, reproducer, all affected IDs/paths, approval dependencies, and whether the next phase is unblocked to the lead and relevant producer.