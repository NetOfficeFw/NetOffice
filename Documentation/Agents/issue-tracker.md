# Issue tracker

This repository uses two complementary tracking systems.

## Local task management

Use Beads (`bd`) for local execution work:

- Claim and update the task being worked on with `bd update <id> --claim`.
- Record dependencies with Beads dependency commands.
- Close completed local tasks with `bd close <id>`.
- Run `bd prime` when starting or resuming work.

The Beads database is local workspace state and is the source of truth for the agent's immediate task queue.

## Shared tickets and specifications

Use GitHub Issues in `NetOfficeFw/agentic-workspace` for work that should be visible to collaborators. Specifications and implementation tickets are both represented as GitHub Issues, with the issue body carrying the relevant detail and links between related issues where useful.

Use the repository's GitHub Project to organize those issues, track shared status, and provide the collaboration view across tickets and specifications.

When a workflow produces shared work, publish or update the corresponding GitHub Issue and add it to the repository's GitHub Project. Use Beads separately for the local execution task and its dependencies; do not treat a local Beads task as a substitute for a shared GitHub Issue.

The GitHub CLI (`gh`) is the preferred interface for GitHub issue and project operations when available. If the project identifier or required project permissions are not available, report that constraint rather than silently omitting the project update.
