---
name: project-codebase-memory-repo-local-workflow
description: ArcTool uses a repo-local codebase-memory-mcp workflow via .codebase-memory/run-cbm.cmd; closure re-index is mandatory and project identity must stay stable across different machine-specific folder paths.
type: project
---
ArcTool's codebase-memory workflow is repo-local: Cowork must launch `codebase-memory-mcp` through `.codebase-memory/run-cbm.cmd` so `CBM_CACHE_DIR` points to the repository's `.codebase-memory/` directory rather than a machine-global cache.

**Why:** Direct MCP execution was indexing into a different internal store, which made graph files in `.codebase-memory/` look stale and broke the team's intended re-index workflow.

**How to apply:** For cross-file architecture, dependency, coupling, impact, or unfamiliar-symbol questions, consult `codebase-memory-mcp` first. Keep `CLAUDE.md` limited to the enforceable execution rules, and keep longer workflow rationale/classification in `.Dossier/Detailed Technical Dossier - ArcTool Knowledge Workflow.md`. When the user says a meaningful session, phase, or section is ending, re-run `index_repository` before closure. Do not hardcode project identity from one machine path because the same repo is used on multiple machines with different folder names.
