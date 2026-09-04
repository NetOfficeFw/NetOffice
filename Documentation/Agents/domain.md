# Domain documentation

This repository uses a single-context documentation layout.

## Layout

- `CONTEXT.md` at the repository root contains durable domain context and terminology.
- `docs/adr/` contains architecture decision records.

## Consumer rules

- Read the relevant `CONTEXT.md` and applicable ADRs before planning or implementing work.
- Use the repository's domain vocabulary in tickets, specifications, and implementation notes.
- Add or update an ADR when a decision has lasting architectural consequences.
- Keep context focused on stable domain knowledge; keep transient task details in Beads or GitHub Issues.

If the repository later grows into independent contexts, replace this document with a root `CONTEXT-MAP.md` and document each context's own `CONTEXT.md` and ADR location.
