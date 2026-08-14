---
name: phase-per-chat-protocol
description: One chat equals one phase; cut the chat proactively at the phase boundary, write .handoff/HANDOFF_TO_NEXT_SESSION.md, then open a new chat instead of relying on auto-compact or summary-of-summary.
metadata:
  type: feedback
---

**One chat = one phase.** Cut proactively; never let the session run until auto-compact fires.

Standing rule for ArcTool normal sessions:

1. A chat carries exactly one phase of work. A phase is one deliverable unit: one bugfix, one
   roadmap step, one documentation pass, one audit, one closure.
2. When the phase reaches a hand-offable state, close the chat deliberately. Do not keep going
   into an unrelated next phase in the same chat just because context still fits.
3. Phase close sequence, executed as the LAST tool actions before the final reply:
   - archive the current `.handoff/HANDOFF_TO_NEXT_SESSION.md` to `.handoff/archive/HANDOFF_<YYYY-MM-DD>_<slug>.md`
   - rewrite `.handoff/HANDOFF_TO_NEXT_SESSION.md` for the phase that just ended
   - update `.handoff/SCRIPT_USAGE_LOG.md` additively if any script was created or used
   - persist any `Memory/` / `.Dossier/` / `CLAUDE.md` durable content the phase produced
   - state plainly in the final reply that the phase is closed and the next phase belongs in a new chat
4. The next phase starts in a NEW chat, seeded from `.handoff/HANDOFF_TO_NEXT_SESSION.md`.
5. Never treat in-chat summarization as the transfer mechanism. If auto-compact has already
   fired and the phase is still open, close the phase at the next safe point rather than
   continuing to stack summaries.

Signals that the phase boundary has arrived (any one is enough):
- the deliverable of the current phase is complete and persisted
- the user shifts to a materially different scope or subsystem
- the work is blocked pending user input, runtime evidence, or an external decision
- the chat is approaching context pressure while the current phase is already hand-offable

**Why:** a summary of a summary is a net loss — it costs about as many tokens as the original
record while being strictly less accurate, and each additional layer compounds the drift.
`.handoff/HANDOFF_TO_NEXT_SESSION.md` is a lossless, reviewable, diffable artifact that exists
precisely to carry state across the boundary, so the boundary should be chosen by the operator,
not by the auto-compact threshold.

**How to apply:** track which phase the current chat owns. When that phase closes, run the phase
close sequence above, then tell the user the phase is closed and the next one should start in a
fresh chat. Do not silently begin a new phase in the same chat. This rule composes with
[[feedback_session_handoff_protocol]] (what the handoff must contain),
[[feedback_persist_memory_before_final_reply]] (durable writes must precede the final reply), and
[[feedback_multi_agent_work_package_workflow]] (work-package sessions keep package-local state).
