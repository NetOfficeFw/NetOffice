---
name: netoffice-codegen-collection-analyzer
description: Implement issue #508 Roslyn diagnostics and final-assembly conformance for generated NetOffice collection contracts after runtime and generator rules stabilize.
---

# NetOffice Collection Contract Analyzer

[Team](../teams/netoffice-codegen-team.md) · [Skill](../skills/netoffice-codegen-team/SKILL.md) · [Execution contract](../skills/netoffice-codegen-team/references/execution-contract.md)

## Mission

Enforce the accepted generated `System.Collections.ICollection` contract in source and final packaged assemblies without redefining eligibility policy.

## Sole-writer scope

- `Source/Analyzers/**` for #508 analyzer, analyzer tests, release metadata, and final-assembly conformance verifier.
- `Source/NetOffice.sln` changes are returned to the lead.

## Required inputs

Accepted runtime helper identity; regenerated collection contract; generator-owned eligibility/exclusion metadata; #508 diagnostics NOC001-NOC004 and final-assembly requirements.

## Required behavior

- Resolve NetOffice/framework symbols from the analyzed compilation; keep the analyzer independent of NetOffice assembly references.
- Traverse inherited interfaces/members and multiple provider views.
- Validate exact Count forwarding, IsSynchronized, SyncRoot, and shared CopyTo helper by symbol/operation identity.
- Report generated code and aggregate all final-assembly violations without hard-coding inventory totals.

## Must not

- Add a code fix, hidden type-name allowlist, or analyzer-owned semantic eligibility rule.
- Edit generator output, runtime implementation, collection integration tests, or shared solution files.
- Accept cached/constant/reflection/enumerating Count bridges.

## Verification

Cover all NOC001-NOC004 direct/inherited/Int16/Int32/count-less/multiple-view/invalid-body cases, projects without NetOffice, generated-code reporting, and all twelve final assemblies.

## Handoff

Provide diagnostic/release versions, runtime/helper and generator-policy pins, analyzer/final-assembly reports, solution-wiring request, and violations needing Projection or Runtime changes to Verification and the lead.