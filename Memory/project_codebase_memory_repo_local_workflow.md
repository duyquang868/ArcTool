---
name: project-codebase-memory-repo-local-workflow
description: ArcTool uses a repo-local codebase-memory-mcp workflow via .codebase-memory/run-cbm.cmd; closure re-index is the final OPTIONAL user-directed step (not a gate on closure), and project identity must stay stable across different machine-specific folder paths.
type: project
---
ArcTool's codebase-memory workflow is repo-local: Cowork must launch `codebase-memory-mcp` through `.codebase-memory/run-cbm.cmd` so `CBM_CACHE_DIR` points to the repository's `.codebase-memory/` directory rather than a machine-global cache.

**Why:** Direct MCP execution was indexing into a different internal store, which made graph files in `.codebase-memory/` look stale and broke the team's intended re-index workflow.

**How to apply:** For cross-file architecture, dependency, coupling, impact, or unfamiliar-symbol questions, consult `codebase-memory-mcp` first. Keep `CLAUDE.md` limited to the enforceable execution rules, and keep longer workflow rationale/classification in `.Dossier/Detailed Technical Dossier - ArcTool Knowledge Workflow.md`. Do not hardcode project identity from one machine path because the same repo is used on multiple machines with different folder names.

**Re-index timing (revised):** `index_repository` only reads files already on disk, so it never needs to run before durable persistence — running it before or after the final reply produces the same graph. Treat it as the last, optional, user-directed step: once `Memory/`, `.Dossier`, `CLAUDE.md`, and ADR writes for the session are done, offer re-index as a choice rather than auto-running it as a closure gate. If the user is out of budget or the session ends first, re-index can be run later from a brand-new chat, because it depends only on the persisted files, not on chat context. Running out of context/tokens must never leave durable files unwritten for the sake of re-index; file persistence always comes first, re-index always comes last and is skippable.

**`manage_adr(mode="update")` is destructive — always read then resubmit the whole store:** the call REPLACES the entire contents of `.codebase-memory/adr.md`. It does not append and it does not merge. Submitting only the new entry silently destroys every entry and prose section already present, with no error and no failed call.

**Why:** this destroyed ADR content repeatedly between 2026-07 and 2026-08-04. 11 items were lost (4 entries plus all 6 prose sections, then 7 more entries one wave at a time) before detection on 2026-08-04 via a store-vs-`CLAUDE.md` mismatch. Only 4 of 13 entries were ever committed to git; the other 9 were recoverable solely from machine-local Claude session transcripts, which preserve every `manage_adr` payload verbatim. All 11 were restored on 2026-08-05.

**How to apply:** call `manage_adr(mode="get")` first, then resubmit the complete store with the new entry appended — for every update, including brand-new entries, not just revisions. Never submit an entry-only payload. Committing `.codebase-memory/adr.md` to git after meaningful ADR changes is the cheap durable backup. Root cause, full call timeline, and per-entry provenance: `.Dossier/ADR Store Loss - Root Cause and Recovery Inventory.md`; verbatim archive of the transcript-recovered entries: `.Dossier/ADR Store Loss - Verbatim Recovery Archive.md`.
