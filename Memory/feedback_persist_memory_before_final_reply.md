---
name: persist-memory-before-final-reply
description: For any meaningful ArcTool session, write durable memory/dossier/CLAUDE.md updates BEFORE the final reply of that turn, because no tool can run after the final response is sent. Do not defer persistence to "after" answering.
type: feedback
---
Persist durable knowledge as the LAST tool actions before the final reply, never after it. Memory is created by tool calls, and once the closing response is emitted no further tools run in that turn; any unwritten reasoning is then unrecoverable.

**Why:** The user wants memory to exist the moment a turn ends so returning later (same chat or next day) resumes from durable files instead of lost in-context reasoning. "Right after you finish answering" is not technically possible; the correct equivalent is "right before the final reply."

**When this applies (meaningful sessions only):** a confirmed bug fix or root cause, a locked architecture/reference decision, an important smoke-test result, a new working rule, a session/phase/section closure, or a handoff to a future session. Skip it for trivial Q&A, chit-chat, or work that produces nothing durable — persisting after every small answer only adds noise.

**How to apply — in order, before sending the closing message:**
1. Finish the substantive work first; do not persist mid-analysis.
2. Update the right durable channel(s): repo-local `Memory/` for cross-session preferences/constraints/pointers, `.Dossier` for bounded deep records, `CLAUDE.md` (English) for short high-leverage invariants/status, and `manage_adr` for stable architecture decisions. Check for an existing record and update it in place instead of duplicating.
3. If the turn ends a session/phase/section with meaningful tracked changes, write a self-contained handoff note into the right durable file so a fresh chat can resume without this context.
4. Refresh the pointer line in `Memory/MEMORY.md` when a new memory file is added.
5. Send the final reply, reporting what was persisted.

**Re-index is deliberately excluded from this pre-reply bundle.** `index_repository` only reads files already on disk, so it carries no data-loss risk and does not need to race the token budget. Treat it as the last, optional, user-directed step, per `project_codebase_memory_repo_local_workflow.md`: offer it after persistence, run it only if the user opts in (in this chat or a later one), and never let it block or delay writing the durable files above.

**Boundary:** This does not force a memory write on every turn. It fixes timing — when persistence is warranted, do it before the final reply, not conceptually "afterward." It complements, and does not replace, `feedback_smoke_test_single_session_close.md` (which still governs smoke-test closure order and forbids using `archive_session` as a save mechanism).
