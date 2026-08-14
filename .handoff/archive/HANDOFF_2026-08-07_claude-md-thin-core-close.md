# ArcTool — HANDOFF TO NEXT SESSION
**Updated:** 2026-08-07  
**Status:** ACTIVE — `CLAUDE.md` startup-context compaction pass is complete through section 9; next session should start on a new scope

---

## Goal and user request

Primary request completed: reduce startup context / token load / overflow risk by compacting `CLAUDE.md` only.

Locked scope the user enforced during the compaction pass:
- ignore other repo files for the optimization work
- load and edit only `CLAUDE.md`
- finish the compaction there before expanding scope

The user then requested normal-session handoff persistence: "chuyển giao dữ liệu ... để ... bắt đầu phiên mới và làm tiếp các mục khác".

This session did not touch product source code.

---

## Current phase / microtask

Current phase: handoff after completing the `CLAUDE.md` thin-core pass.

Completed across this compaction mission:
- compacted the remaining tail of `## 5. Technical decisions already locked`
- compacted `## 6. Active roadmap`
- compacted `## 7. Closed technical dossier — recent closure record`
- compacted `## 8. Coding rules`
- compacted `## 9. API references`
- refreshed normal-session handoff state
- archived the previous handoff snapshot to `.handoff/archive/HANDOFF_2026-08-07_claude-md-compaction-sections-5-7.md`

Current next-session starting point:
- do not continue `CLAUDE.md` compaction by default; that pass is done
- resume from the user's next requested item outside this optimization scope

---

## Files modified in this session

Modified:
- `CLAUDE.md`
- `.handoff/HANDOFF_TO_NEXT_SESSION.md`

Created earlier in the broader compaction effort and still relevant as current root pointer targets:
- `.Dossier/ArcTool Locked Technical Decisions.md`
- `.Dossier/ArcTool Revit API Reference Notes.md`

Archived during this handoff close:
- `.handoff/archive/HANDOFF_2026-08-07_claude-md-compaction-sections-5-7.md`

Referenced but not edited in this closeout step:
- `.handoff/SCRIPT_USAGE_LOG.md`
- `.Dossier/ArcTool Prompt Cache Tiers.md`
- `Memory/feedback_session_handoff_protocol.md`

No product code files were edited for this request.

---

## Exact `CLAUDE.md` progress

Compaction status:
- section 5 tail compacted into grouped locked-decision bullets
- section 6 compacted into concise roadmap bullets
- section 7 compacted into a short recent-closure record with dossier pointers
- section 8 compacted into root-only coding invariants plus one subsystem pointer
- section 9 compacted into the Revit API reference-of-record rule plus one durable API register pointer

Final size snapshot after the tail compaction pass:
- `CLAUDE.md` = 188 lines, 17,832 bytes
- `git diff --stat -- CLAUDE.md` = 92 insertions, 329 deletions

Resulting state:
- root file is materially thinner
- implementation-affecting invariants remain at the root
- deeper detail is routed behind `.Dossier` pointers
- startup/prompt-cache pressure should be lower because `CLAUDE.md` sits in T1 per `.Dossier/ArcTool Prompt Cache Tiers.md`

---

## Scripts generated or used in this session

Used / referenced durable script asset:
- `.codebase-memory/run-cbm.cmd` — repo-local launcher for `codebase-memory-mcp`; known existing script from the cumulative inventory, not executed in this session

Generated this session:
- none

Cumulative script-log update needed:
- none; no helper script was created or executed in this session, so `.handoff/SCRIPT_USAGE_LOG.md` counts stay unchanged

---

## Locked decisions and reasons

1. **Scope stayed `CLAUDE.md`-only for the optimization pass.**
   - Reason: the user explicitly constrained the work to startup-context reduction in that file.

2. **No product-code edits during context compaction.**
   - Reason: the task was documentation/context-density work, not feature or bug work.

3. **Preserve numbering, headings, and technical meaning while shrinking prose.**
   - Reason: `CLAUDE.md` is an operating document and must remain stable and navigable.

4. **Keep root content to implementation-affecting invariants; route depth to dossier pointers.**
   - Reason: repeated detail in `CLAUDE.md` is startup-token tax and harms prompt-cache efficiency.

5. **No Revit runtime action without explicit user instruction.**
   - Reason: runtime remains operator-controlled by project rule.

6. **Handoff state belongs in `.handoff/`, not chat only.**
   - Reason: the user explicitly requested data handoff for a fresh session, and the handoff protocol memory requires repository persistence.

---

## Done / unfinished / blocked

Done:
- `CLAUDE.md` compaction completed through section 9
- handoff snapshot updated to reflect the true finished state
- previous handoff snapshot archived before replacement

Unfinished:
- the user's next substantive repo task ("các mục khác") was not started in this session
- earlier scope-drift artifacts outside `CLAUDE.md` were not reconciled during this closeout

Blocked:
- none technically

Notable caveat:
- `.Dossier/ArcTool Locked Technical Decisions.md` and `.Dossier/ArcTool Revit API Reference Notes.md` are useful durable pointer targets, but they were created during the broader compaction effort outside the user's strict `CLAUDE.md`-only scope

---

## Verification run

Verification completed:
- confirmed handoff protocol requirements from `Memory/feedback_session_handoff_protocol.md`
- confirmed prompt-cache rationale from `.Dossier/ArcTool Prompt Cache Tiers.md`
- confirmed final `CLAUDE.md` size: 188 lines, 17,832 bytes
- confirmed `CLAUDE.md` diff shrink trend: 92 insertions, 329 deletions
- confirmed the previous handoff file was stale and replaced it with the finished-state snapshot
- archived the previous handoff snapshot before overwriting

Not run:
- no build
- no tests
- no Revit runtime action
- no Revit MCP action
- no re-index

Reason not run:
- task scope was documentation compaction plus handoff persistence only; runtime/build/index actions were neither required nor requested

---

## Next-session starting point

Start from the user's next requested item outside the `CLAUDE.md` thin-core mission.

Practical resume point:
- treat `CLAUDE.md` compaction as closed
- use this handoff only as continuity state
- if follow-up concerns arise about the thin-core split, review `CLAUDE.md`, `.Dossier/ArcTool Locked Technical Decisions.md`, and `.Dossier/ArcTool Revit API Reference Notes.md` in that order

---

## Invariants to preserve

1. `CLAUDE.md` stays English-only and edited in place.
2. Preserve numbering, headings, and technical meaning when compacting root operating docs.
3. Keep root files short; push deep history/rationale into `.Dossier`.
4. Do not expand documentation-compaction tasks into code changes unless the user changes scope.
5. Runtime stays operator-controlled: no Revit launch, MCP, or smoke without explicit request.
6. Normal-session handoff stays in `.handoff/`; work-package state stays package-local.
7. Treat `.handoff/HANDOFF_TO_NEXT_SESSION.md` as volatile T4 state, never durable T1/T3 context.

---

## Reference files

- Handoff protocol memory: `Memory/feedback_session_handoff_protocol.md`
- Script ledger: `.handoff/SCRIPT_USAGE_LOG.md`
- Prompt-cache rationale: `.Dossier/ArcTool Prompt Cache Tiers.md`
- Archived previous handoff snapshot: `.handoff/archive/HANDOFF_2026-08-07_claude-md-compaction-sections-5-7.md`
- Current thin-core root file: `CLAUDE.md`
- Durable decision pointer target: `.Dossier/ArcTool Locked Technical Decisions.md`
- Durable API pointer target: `.Dossier/ArcTool Revit API Reference Notes.md`
