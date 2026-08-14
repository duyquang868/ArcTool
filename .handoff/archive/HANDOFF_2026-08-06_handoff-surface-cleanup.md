# ArcTool — HANDOFF TO NEXT SESSION
**Updated:** 2026-08-06  
**Status:** ARCHIVED 2026-08-07 — superseded by the startup-context compaction handoff

---

## Mission outcome

This session cleanup converts the handoff from a stale setup-phase snapshot into a clean canonical
normal-session handoff surface.

- `.handoff/` is the established repository location for **normal-session** transfer state.
- Work-package sessions remain package-local and are **not** redirected into `.handoff/`.
- Script retention stays locked: generated or reused scripts are durable assets and must be kept.
- Script usage logging stays locked: counts are cumulative and never reset.
- The earlier setup-phase snapshot was archived to preserve history without leaving obsolete
  "implementation in progress" guidance in the live handoff.

---

## Current durable state

The `.handoff/` protocol is already in force and reflected across the repository.

Established handoff layout:
- `.handoff/README.md`
- `.handoff/HANDOFF_TO_NEXT_SESSION.md`
- `.handoff/SCRIPT_USAGE_LOG.md`
- `.handoff/archive/`
- `.handoff/scripts/`

Established supporting records:
- `Memory/feedback_session_handoff_protocol.md`
- `Memory/MEMORY.md`
- `CLAUDE.md` already contains the `.handoff/` code-map subtree and the rule routing
  normal-session handoff state into `.handoff/`

Archived historical snapshot:
- `.handoff/archive/HANDOFF_2026-08-06_handoff-protocol-setup.md`

---

## Files created or updated for this cleanup

Created:
- `.handoff/archive/HANDOFF_2026-08-06_handoff-protocol-setup.md`

Retained as active handoff surface:
- `.handoff/HANDOFF_TO_NEXT_SESSION.md`
- `.handoff/README.md`
- `.handoff/SCRIPT_USAGE_LOG.md`

Previously established supporting records still referenced by this handoff:
- `Memory/feedback_session_handoff_protocol.md`
- `Memory/MEMORY.md`
- `CLAUDE.md`

---

## Scripts used in this session

Tracked reusable script inventory relevant to the handoff protocol:
- `.codebase-memory/run-cbm.cmd` — repo-local launcher for `codebase-memory-mcp`

No new reusable helper script was generated for this cleanup.

---

## Build / verification status

No build run.

No Revit runtime action run.

Verification completed:
- `.handoff/README.md` defines scope, retention, archiving, and the work-package exception.
- `.handoff/SCRIPT_USAGE_LOG.md` uses cumulative additive-count semantics.
- `Memory/feedback_session_handoff_protocol.md` defines the durable handoff rule and required contents.
- `CLAUDE.md` already reflects the `.handoff/` routing rule and code-map subtree, so no repository-policy follow-up remains pending for this cleanup.
- The previous live handoff snapshot was archived before cleaning the canonical latest handoff.

Verification outcome:
- No stale live guidance remains claiming `.handoff/` protocol implementation is still in progress.

---

## Locked decisions

1. **Dedicated handoff folder = `.handoff/`.**
   - Reason: session handoff changes continuously and does not fit the purpose of `Memory/` or
     `.Dossier/`.

2. **All scripts and logs are retained.**
   - Reason: repeated recreation of near-identical one-off helpers wastes tokens across sessions.

3. **Script usage log is cumulative and never resets.**
   - Reason: script usage is a durable asset, not chat-only context.

4. **Work-package sessions keep using package-local persistence.**
   - Reason: `.claude/workpackages/*` and `.claude/quick-dimension-bugfix/*` already have a
     dedicated context-protection model.

5. **No ADR update for this change.**
   - Reason: this is workflow/persistence policy, not a new product architecture decision.

---

## Next-session starting point

Use this file as the latest normal-session handoff surface.

When a future normal session ends with tracked changes, archive the current version to
`archive/HANDOFF_<YYYY-MM-DD>_<short-slug>.md`, then replace this file with the new latest handoff.

Do not move or duplicate any work-package state file into `.handoff/`.

---

## Invariants to preserve

1. **Runtime stays operator-owned.** No Revit launch, Revit MCP, or smoke action without explicit request.
2. **Normal-session handoff belongs in `.handoff/`, not `Memory/` or `.Dossier/`.**
3. **Script retention has no deletion shortcut.** Keep scripts and log them durably.
4. **Script usage counts are additive.** Never reset prior counts.
5. **`CLAUDE.md` edits must stay in English and in-place only.**
6. **Work-package exception stays intact.** `.claude/quick-dimension-bugfix/` and `.claude/workpackages/*` remain authoritative for package sessions.

---

## Reference files

- Source protocol: `D:/Quang mini/OneDrive - MSFT/BotTrader/.memory/chuyen-giao-du-lieu-protocol.md`
- Handoff-surface definition: `.handoff/README.md`
- Archived setup snapshot: `.handoff/archive/HANDOFF_2026-08-06_handoff-protocol-setup.md`
- Work-package exemplar handoff: `.claude/quick-dimension-bugfix/HANDOFF_TO_NEXT_SESSION.md`
- Work-package state model: `.claude/workpackages/_TEMPLATE/06_EXECUTION_STATE.md`
- Work-package shared contract: `.claude/workpackages/_TEMPLATE/01_SHARED_CONTRACT.md`
- Persistence timing rule: `Memory/feedback_persist_memory_before_final_reply.md`
- Work-package workflow rule: `Memory/feedback_multi_agent_work_package_workflow.md`
- Protocol memory record: `Memory/feedback_session_handoff_protocol.md`
- Memory index: `Memory/MEMORY.md`
- Technical operating context: `CLAUDE.md`
