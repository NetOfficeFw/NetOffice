---
name: netoffice-codegen-team
description: Governed team for NetOffice issue #509 covering Data v2, typelib acquisition, Roslyn wrapper generation, VBA documentation, deterministic ownership-safe regeneration, and linked collection issues #506-#508.
---

# NetOffice Code Generator Team

## Mission

Deliver issue [#509](https://github.com/NetOfficeFw/NetOffice/issues/509) under its binding D1-D11 decisions. Produce an independent MIT-licensed generator that reproduces the current wrapper contract from Data v2, emits deterministic C# 7.3, incorporates pinned VBA documentation, safely owns `Source/<Api>/Generated/`, and supports the linked collection capability/analyzer/tests.

Shared operating instructions live in the [team skill](../skills/netoffice-codegen-team/SKILL.md). Detailed gates live in the [execution contract](../skills/netoffice-codegen-team/references/execution-contract.md).

## Leadership and roster

Lead: [netoffice-codegen-lead](../agents/netoffice-codegen-lead.md)

### Core team

| Role | Definition | Activation | Primary result |
| --- | --- | --- | --- |
| Integration lead | [netoffice-codegen-lead](../agents/netoffice-codegen-lead.md) | Always | Versioned contracts, ordered integration, shared wiring |
| Parity contract | [netoffice-codegen-parity](../agents/netoffice-codegen-parity.md) | Contract wave and every gate | Current-source Wrapper Contract and ownership classification |
| Data v2 | [netoffice-codegen-data-v2](../agents/netoffice-codegen-data-v2.md) | Contract/data wave | Merged, pinned v2 graph and signed migration evidence |
| Typelib | [netoffice-codegen-typelib](../agents/netoffice-codegen-typelib.md) | After v2 schema | Raw observations and deterministic merge reports |
| Projection | [netoffice-codegen-projection](../agents/netoffice-codegen-projection.md) | After merged v2 SHA | Typed projected wrapper/file decisions and explain traces |
| Emission and safety | [netoffice-codegen-emission](../agents/netoffice-codegen-emission.md) | After WrapperFile fixtures | Roslyn bytes, incremental cache, transactional write plan |
| Documentation | [netoffice-codegen-docs](../agents/netoffice-codegen-docs.md) | After merged v2 SHA | Docs index/mapping/XML-doc digest and attribution |
| CLI orchestration | [netoffice-codegen-cli](../agents/netoffice-codegen-cli.md) | After producer APIs stabilize | Validated command pipeline and fixed exit behavior |
| Verification | [netoffice-codegen-verification](../agents/netoffice-codegen-verification.md) | Every phase gate | Independent reports and final conformance evidence |

Temporary read-only role: [netoffice-codegen-slice-coordinator](../agents/netoffice-codegen-slice-coordinator.md).

### Post-cutover linked-issue roles

| Role | Definition | Activation |
| --- | --- | --- |
| Runtime capability (#506) | [netoffice-codegen-runtime](../agents/netoffice-codegen-runtime.md) | After ownership cutover and accepted generator capability contract |
| Collection analyzer (#508) | [netoffice-codegen-collection-analyzer](../agents/netoffice-codegen-collection-analyzer.md) | After runtime capability and regenerated contracts |
| Collection integration tests (#507) | [netoffice-codegen-collection-tests](../agents/netoffice-codegen-collection-tests.md) | After runtime/analyzer interfaces stabilize |

## Dependency flow

```text
Frozen NetOffice/Data/VBA-Docs commits
        |
        +--> Parity Wrapper Contract -----------+
        |                                      |
        +--> Data v2 schema -> full conversion +--> pinned Data v2 SHA
                                                     |
                   +---------------------------------+-------------------+
                   |                  |              |                   |
                Typelib          Projection      Documentation       Emission
                   +------------------+--------------+-------------------+
                                                     |
                                               CLI composition
                                                     |
               DAO -> Excel -> PowerPoint -> Outlook/Access -> ADODB -> Office
                                                     |
                                      zero unexplained full parity
                                                     |
                              format -> path move -> ownership bootstrap
                                                     |
                     Current Channel import -> VBA docs product batches
                                                     |
                                  #506 -> regenerate -> #508 and #507
```

Projection, Documentation binding, CLI composition, and slices require the merged Data v2 SHA. Typelib and Emission may work earlier only against accepted schema fixtures. At most Typelib, Projection, Documentation, and Emission run concurrently.

## Sole-writer ownership

| Owner | Sole-writer paths/artifacts |
| --- | --- |
| Lead | `.agents/**`; `Tools/CodeGen/NetOffice.CodeGen.sln`; shared `Directory.Build.*`, locks and local-tool manifests; `Source/NetOffice.sln`; `Source/NetOffice.props`; CI; integration branches; verified generated/manifests commits |
| Parity | `Tools/CodeGen/tools/NetOffice.CodeGen.ContractExtractor/**`; `Tools/CodeGen/contracts/wrapper/**`; matching tests |
| Data v2 | External `NetOfficeFw/Data` v2 schema/graph/migration evidence; `Tools/CodeGen/src/NetOffice.CodeGen.Data/**`; matching tests |
| Typelib | `Tools/CodeGen/src/NetOffice.CodeGen.TypeLib/**`; matching tests |
| Projection | `Tools/CodeGen/src/NetOffice.CodeGen.Projection/**`; `Tools/CodeGen/policy/**`; matching tests |
| Emission | `Tools/CodeGen/src/NetOffice.CodeGen.Emit/**`; `Tools/CodeGen/src/NetOffice.CodeGen.Storage/**`; matching tests; manifest schema |
| Documentation | `Tools/CodeGen/src/NetOffice.CodeGen.Docs/**`; matching tests; `Tools/CodeGen/docs/*.json`; `THIRD-PARTY-NOTICES.md` |
| CLI | `Tools/CodeGen/src/NetOffice.CodeGen.Application/**`; `Tools/CodeGen/src/NetOffice.CodeGen.Cli/**`; matching tests |
| Verification | `Tools/CodeGen/tests/NetOffice.CodeGen.Verification/**`; cross-component/fault fixtures; conformance and package tests |
| Slice coordinator | No working-tree path; read-only reports and routing |
| Runtime | `Source/NetOffice/**` for #506 |
| Analyzer | `Source/Analyzers/**` for #508 |
| Collection tests | `Source/NetOffice.Tests/**` for #507 |

A cross-owner change is split into producer contract/API, consumer update, then lead-owned shared wiring. Consumers never edit producer fixtures. Generated files are never hand-patched.

## Review pairs

- Parity reviews Projection against the Wrapper Contract.
- Data v2 reviews Typelib's schema/identity use.
- Emission reviews Documentation AST consumption.
- CLI reviews every producer command adapter.
- Verification independently signs every phase/slice gate.
- Lead resolves only cross-contract integration; it does not invent domain policy.

## Fixed slice workflow

Slice order: DAO; Excel Range/WorksheetFunction; PowerPoint events/collections; Outlook/Access event conflicts; ADODB duplicates/inheritance; Office Adjustments.

For each slice:

1. Accept and pin Parity classification/contract.
2. Projection and Documentation submit independent artifacts pinned to that contract and Data v2 SHA. Before cutover use docs profile `baseline` only.
3. Emission produces a pinned dry-run byte/diff artifact.
4. CLI composes accepted versions without bypassing validators.
5. Verification independently reproduces and signs pins, profile, command, hashes, and differences.
6. Lead integrates only when unexplained differences are zero.

Intentional differences require an approved compatibility-ledger record with logical ID, facet, expected match count, rationale, provenance, and compatibility status before producer changes.

## Human approval boundary

Only the maintainer approves:

- Data and documentation ambiguities.
- Removal overrides and compatibility classifications.
- NetOffice, Data v2, and VBA-Docs pins.
- Microsoft 365 Current Channel acquisition evidence.
- CC BY 4.0 attribution and affected-package ledger.

Agents report blocked records and continue independent records. They never manufacture approvals or observed Office evidence.

## Branch and handoff contract

Branches: `agent/509-<workstream>-<slice>`. A branch changes one owner's paths plus namespaced fixtures. Cross-owner changes use dependency-ordered branches/PRs.

Each handoff records:

- producer role and accepted artifact/schema version;
- exact input commits and policy/package hashes;
- logical IDs and source/provenance locations;
- exact command and observed result;
- expected and actual affected files/hashes;
- unresolved records and required approver;
- next owner and acceptance condition.

If the active agent runtime supports peer IDs, the lead assigns and shares them. Otherwise, return the handoff in the completion report; never guess an address.

## Completion condition

Issue #509 is complete only when pinned inputs reproduce all twelve API projects offline with zero unexplained contract differences; cacheless and incremental bytes match; ownership/bootstrap/crash-safety gates pass; VBA docs and attribution are corpus/package complete; performance budgets pass; Current Channel import evidence is approved; and #506/#508/#507 capability, analyzer, managed, and real-COM gates pass. Retire the legacy generator only then.
