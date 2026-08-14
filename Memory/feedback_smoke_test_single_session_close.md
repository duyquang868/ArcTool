---
name: smoke-test-single-session-close
description: Each Quick Dimension smoke-test feedback/audit must be completed within one chat session; immediately persist all findings to repo-local Markdown/XML, then ask whether to end the working session without archiving it.
type: feedback
---
Treat every smoke-test review as a one-session unit of work, because a full smoke-test audit consumes a large amount of tokens and cannot safely span multiple chat sessions.

**Terminology (do not confuse):** "Persist / store / archive the information" ALWAYS means writing durable files — Markdown under `Memory/` and `.Dossier`, XML logs, and `CLAUDE.md` updates. It does NOT mean putting the chat into Archive mode. The `archive_session` action (moving a chat to the Archived list) is a separate, irreversible operation that must NEVER be used as the way to "save" a smoke test. Only run `archive_session` when the user explicitly asks to archive the chat with those words.

**Why:** The user runs each smoke test in a single chat session on purpose to bound token cost. Analysis context (XML logs, annotated images, per-candidate reasoning) is expensive to rebuild, so it must be captured into durable files before the working session ends rather than carried forward into another session.

**How to apply — do this in order at the end of any smoke-test feedback session:**
1. Finish the full analysis first (arithmetic vs stations, geometry/classification, reference ownership/metadata, latent `NewDimension` risks). Do not close mid-analysis.
2. Immediately record the complete outcome of that smoke test into durable repo-local memory under `Memory/` — capture ALL of: confirmed successes, real defects/anomalies (with a `BUG-xx` id when applicable and its severity), warnings, notes/caveats, the concrete numeric/reference evidence, and every unresolved gate. Prefer appending to the relevant `project_qd_*` evidence file; if the smoke round is new, create a dedicated `project_qd_*_smoke_*.md` note. Also add/refresh the pointer line in `Memory/MEMORY.md`.
3. Reflect durable status in `CLAUDE.md` (English) only as a short summary/bug-row/status update, per the existing lean-CLAUDE and code-map-review rules.
4. After durable files in steps 2-3 are written and verified, ASK the user whether to end/close the current working session. If the user says close, report that the smoke-test work is complete and stop the conversation; do NOT call `archive_session`. If the user declines, keep working in the same chat. Calling `archive_session` is permitted only when the user explicitly requests the separate Archive-mode action.
5. Codebase-memory re-index (`.codebase-memory/run-cbm.cmd`) is the final, OPTIONAL, user-directed step, not a precondition for asking to close. Offer it once durable files are safe; run it only if the user opts in, in this chat or a later one. Because re-index only reads already-persisted files, it is never worth risking a token/context cutoff for — file persistence in steps 2-3 is what must never be skipped.

**Enforcement note:** This is a memory/workflow rule, not an automated hook. A hook cannot reliably detect the moment "smoke-test analysis is complete". Session closure therefore stays deliberate: persist files first, confirm whether the user wants to end the working session, then offer re-index as a final optional step. Ending a working session means stop working/responding after the closure summary; it is not the same as archiving the chat.
