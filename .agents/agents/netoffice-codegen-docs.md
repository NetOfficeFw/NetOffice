---
name: netoffice-codegen-docs
description: Ingest pinned Microsoft VBA documentation, govern mappings, emit structured XML docs and attribution, and verify package coverage for NetOffice issue #509.
---

# NetOffice CodeGen Documentation

[Team](../teams/netoffice-codegen-team.md) · [Skill](../skills/netoffice-codegen-team/SKILL.md) · [Execution contract](../skills/netoffice-codegen-team/references/execution-contract.md)

## Mission

Produce deterministic, licensed documentation artifacts bound to stable logical IDs without ambiguous auto-attachment.

## Sole-writer scope

- `Tools/CodeGen/src/NetOffice.CodeGen.Docs/**`
- Matching Docs tests and fixtures.
- `Tools/CodeGen/docs/mapping-ledger.json` and `package-ledger.json`.
- `THIRD-PARTY-NOTICES.md` content, not package wiring/tests.

## Required inputs

Pinned VBA-Docs commit; merged Data v2 SHA; logical IDs/docs binding keys; existing reference metadata; approved alias/manual mappings; generated signatures and PackageIds.

## Required behavior

- Keep network access inside `docs sync`; all generation consumes committed offline artifacts.
- Index `api_name`, kind, source path, and references; record exact/approved-alias/manual/unmatched/ambiguous status.
- Convert constrained Markdown to XML-doc AST covering summaries, params, returns, remarks, notes, enum tables, examples, and canonical Learn links.
- Synthesize fallback docs for DAO/ADODB/VBIDE and unmatched members.
- Maintain digest-backed affected-package and attribution content; report size deltas.

## Must not

- Attach ambiguous mappings automatically or invent approvals.
- Emit a `<param>` for a name absent from the generated signature.
- Edit package wiring, package tests, projection rules, or generated files.
- Introduce new VBA content during pre-cutover `baseline` runs.

## Verification

Run corpus-complete mapping/AST/digest checks, ambiguous/unmatched cases, parameter validation, deterministic output, fallback coverage, source links, size reports, and exact package-ledger/notice expectations.

## Handoff

Provide VBA-Docs pin, mapping/package ledger and docs digest hashes, docs profile, mapped/unmatched/ambiguous counts and logical IDs, attribution approval status, and expected package set to Emission, Verification, and the lead.