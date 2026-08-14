# ArcTool — HANDOFF TO NEXT SESSION
**Updated:** 2026-08-07  
**Status:** ACTIVE — `CLAUDE.md` rules-only root refactor is complete and the phase is closed; the next phase must start in a new chat

---

## Goal and user request

Primary request for this phase:
- audit whether any rule still requires updating `CLAUDE.md` after every session
- enforce the user's storage model that `CLAUDE.md` must contain only project working rules
- preserve all existing meaning while splitting durable technical content into dedicated reference channels
- restructure only from already-existing `CLAUDE.md` content; do not invent new doctrine

Locked user intent for this phase:
- `CLAUDE.md` is only the project's working-rule and operating-rule document
- durable technical knowledge must live in separate channels and be referenced on demand
- preservation-first refactor: extract existing content rather than replace it with newly authored technical content

---

## Current phase / microtask

Current phase: close the `CLAUDE.md` thin-root documentation pass after verifying that the remaining root content is rules-only and that removed durable content has a correct destination.

Completed in this phase:
- audited the live `CLAUDE.md` against the user's clarified storage model
- corrected the root editing policy so section 10 now matches the rules-only intent
- verified the real repo structure before preserving section 2 `Code map`
- confirmed section-5 locked-decision content is preserved in `.Dossier/ArcTool Locked Technical Decisions.md`
- confirmed roadmap / closure / API material is represented by dedicated pointer destinations instead of root payload
- recovered the exact BUG-06 / BUG-07 / BUG-08 rows that had been removed from `CLAUDE.md`
- created a durable `.Dossier` register for those recovered open bugs
- updated `CLAUDE.md` section 4 so the root now points to that new open-bug register
- archived the prior unrelated handoff state and rewrote the active handoff for this closed documentation phase

This handoff is the closure record for the `CLAUDE.md` rules-only refactor phase.

---

## Files modified in this session

Modified:
- `CLAUDE.md`
- `.handoff/HANDOFF_TO_NEXT_SESSION.md`

Created:
- `.Dossier/ArcTool Open Bug Register.md`
- `.handoff/archive/HANDOFF_2026-08-07_result-only-no-worktree-close.md`

Referenced but not modified:
- `.Dossier/ArcTool Locked Technical Decisions.md`
- `.Dossier/ArcTool Revit API Reference Notes.md`
- `.Dossier/Quick Dimension - Implementation Roadmap.md`
- `.Dossier/Detailed Technical Dossier - Coordinate Feature.md`
- `.handoff/README.md`
- `.handoff/SCRIPT_USAGE_LOG.md`
- `Memory/feedback_phase_per_chat_protocol.md`

No product source-code files were edited for this request.

---

## Exact implementation progress

1. `CLAUDE.md` root structure
   - kept the file as the operating document
   - retained working rules, compact platform invariants, and routing pointers only
   - kept section numbering and headings intact
   - aligned section 10 so future edits do not repopulate the root with roadmap, bug-matrix, or historical narrative content

2. Durable extraction validation
   - verified the locked-decision payload already lives in `.Dossier/ArcTool Locked Technical Decisions.md`
   - verified roadmap and closure routing remains delegated to `.Dossier` / `Memory/` destinations
   - verified API guidance stays rooted in Revit API lookup plus `.Dossier/ArcTool Revit API Reference Notes.md`

3. Preservation-gap repair
   - identified one real preservation defect: BUG-06 / BUG-07 / BUG-08 had been removed from `CLAUDE.md` without a durable destination
   - recovered the exact rows from the `git diff` against `CLAUDE.md`
   - created `.Dossier/ArcTool Open Bug Register.md` and persisted the three rows verbatim
   - updated section 4 of `CLAUDE.md` so the root points to the new register

---

## Evidence found during verification

Key evidence gathered while verifying preservation fidelity:
- the live `CLAUDE.md` already held the corrected thin-root section 10 policy at close time
- `.Dossier/ArcTool Locked Technical Decisions.md` contains the extracted decision payload for General/platform, Excel to Revit, Coordinate, and Quick Dimension
- `.handoff/SCRIPT_USAGE_LOG.md` shows no script activity in this phase, so the ledger required no update
- `git diff -- CLAUDE.md` exposed the three removed bug rows exactly as:
  - `| BUG-06 | ArrangeDimension | Missing guard for activeView.Scale == 0 / unsupported view contexts | Medium |`
  - `| BUG-07 | FilterManager | Idling-based refresh architecture does not scale on large models | Low |`
  - `| BUG-08 | CreateVoidFromLink | SetParam("Height", -beamHeight) is still a workaround, not a clean model | Low |`
- the durable fix for those rows is now `.Dossier/ArcTool Open Bug Register.md`

Interpretation:
- the thin-root refactor is now preservation-first instead of lossy
- `CLAUDE.md` no longer carries open-bug payload directly, but the extracted content remains recoverable through an explicit pointer

---

## Scripts generated or used in this session

Generated:
- none

Used:
- none

Cumulative script-log update needed:
- none; `.handoff/SCRIPT_USAGE_LOG.md` remains unchanged because no script was created or executed in this phase

---

## Locked decisions and reasons

1. **Keep `CLAUDE.md` as a rules-only root.**
   - Reason: this is the user's explicit storage model and avoids root-document drift.

2. **Preserve meaning by extraction, not by rewriting doctrine.**
   - Reason: the user explicitly prohibited replacing existing content with newly invented technical material.

3. **Use `.Dossier` for recovered cross-feature bug records.**
   - Reason: section 10 now routes bug matrices and non-root bug detail out of `CLAUDE.md`, and no existing durable record owned BUG-06/07/08.

4. **Do not update `.handoff/SCRIPT_USAGE_LOG.md`.**
   - Reason: no script asset was created or executed in this phase.

5. **Do not run build, tests, Revit, MCP, or re-index.**
   - Reason: this phase was documentation/persistence work only, and runtime actions were neither required nor authorized.

---

## Done / unfinished / blocked

Done:
- `CLAUDE.md` now conforms to the user's rules-only root model
- the root editing policy is aligned with that model
- the removed BUG-06/07/08 rows are durably preserved in `.Dossier/ArcTool Open Bug Register.md`
- section 4 now points to that durable bug register
- the active handoff now matches this closed phase

Unfinished:
- no further documentation move remains within this phase's stated scope

Blocked:
- none technically

---

## Verification run

Verification completed:
- reviewed live `CLAUDE.md` structure and confirmed sections 3/4/6/7 are pointer-only
- confirmed section 10 now enforces the rules-only root policy
- verified the code-map section against the real repo layout
- verified extraction fidelity for locked decisions through `.Dossier/ArcTool Locked Technical Decisions.md`
- checked the script ledger before close and left it unchanged
- recovered the removed bug rows from `git diff` and persisted them durably
- completed handoff archival and rewrite before final reply

Not run:
- no build
- no tests
- no Revit runtime action
- no Revit MCP action
- no re-index

Reason not run:
- this was documentation/persistence work only, and those actions were neither required nor requested

---

## Next-session starting point

Start a NEW chat for the next phase.

At the start of that new chat:
- treat the `CLAUDE.md` rules-only refactor as closed
- use `.handoff/HANDOFF_TO_NEXT_SESSION.md` as the transfer surface, not in-chat summaries
- resume only from the user's next explicitly chosen scope

If the next chat needs the source of record for this phase, read in this order:
1. `CLAUDE.md`
2. `.Dossier/ArcTool Open Bug Register.md`
3. `.Dossier/ArcTool Locked Technical Decisions.md`
4. `.handoff/README.md`
5. `Memory/feedback_phase_per_chat_protocol.md`
6. this handoff file

---

## Invariants to preserve

1. `CLAUDE.md` stays a technical operating document, not a narrative project history.
2. Durable technical knowledge is classified into `.Dossier`, `Memory/`, ADR, or `.handoff/` instead of being accumulated at the root.
3. Revit runtime remains operator-controlled: no Revit launch, MCP, or smoke without explicit request.
4. Revit API answers/fixes still require lookup against https://www.revitapidocs.com/2026/.
5. One chat equals one phase; close the phase via handoff before the final reply.

---

## Reference files

- Root operating document: `CLAUDE.md`
- Recovered open bugs: `.Dossier/ArcTool Open Bug Register.md`
- Locked decisions register: `.Dossier/ArcTool Locked Technical Decisions.md`
- API register: `.Dossier/ArcTool Revit API Reference Notes.md`
- Handoff protocol: `.handoff/README.md`
- Phase-boundary rule: `Memory/feedback_phase_per_chat_protocol.md`
- Script ledger: `.handoff/SCRIPT_USAGE_LOG.md`
- Previous active handoff archived as: `.handoff/archive/HANDOFF_2026-08-07_result-only-no-worktree-close.md`
