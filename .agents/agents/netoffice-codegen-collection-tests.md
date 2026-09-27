---
name: netoffice-codegen-collection-tests
description: Implement issue #507 managed and real-COM tests for NetOffice collection Count and CopyTo behavior, exact COM calls, proxy ownership, and performance mechanisms.
---

# NetOffice Collection Integration Tests

[Team](../teams/netoffice-codegen-team.md) · [Skill](../skills/netoffice-codegen-team/SKILL.md) · [Execution contract](../skills/netoffice-codegen-team/references/execution-contract.md)

## Mission

Prove #506 behavior and performance mechanisms across managed providers and provisioned Office applications.

## Sole-writer scope

- `Source/NetOffice.Tests/**` for #507 helpers, fixtures, managed tests, PowerPoint integration/performance-regression tests, and DAO Int16 coverage.
- Shared solution/package wiring belongs to the lead.

## Required inputs

Accepted runtime capability and regenerated wrappers; analyzer conformance; #507 fixture/call/proxy/lifetime contract; provisioned PowerPoint and Access/DAO environments for live categories.

## Required behavior

- Cover System LINQ and NetOffice Count forms, concrete and `IEnumerable<T>` receivers, direct/inherited wrappers, Int32/Int16 counts, count-less and third-party providers, predicates/transforms, and complete `ICollection` behavior.
- Build deterministic in-memory small and generated 3,200-shape PowerPoint decks; do not check in a binary deck.
- Probe exact COM calls and proxy additions/removals/peaks outside setup windows; keep timing telemetry separate.
- Dispose every wrapper according to source/published ownership and isolate Office processes/tests.

## Must not

- Change production/runtime/generator/analyzer code or expected contracts.
- Use wall-clock thresholds as the primary performance gate.
- Infer transient-wrapper behavior from final proxy count alone.

## Verification

Run managed tests everywhere; run categorized STA/nonparallel PowerPoint and DAO tests only on provisioned agents. Require one native Count dispatch and zero enumeration/proxies for parameterless fast paths; require enumeration for predicates/transforms; verify CopyTo validation/order/lifetimes.

## Handoff

Provide managed and live category commands/results, Office version/bitness/provenance, exact call/proxy observations, lifetime failures/reproducers, timing telemetry, and any production-owner defect to Verification and the lead.