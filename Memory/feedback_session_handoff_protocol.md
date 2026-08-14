---
name: session-handoff-protocol
description: Normal-session handoff data lives in .handoff/ with cumulative script retention and logging; work-package sessions stay package-local.
metadata:
  type: feedback
---

When the user says **"chuyển giao dữ liệu"** (or a close variant), treat that as an instruction to
persist the full normal-session handoff state to repository files instead of leaving it in chat.
For ArcTool, that state belongs in `.handoff/`, not `Memory/` or `.Dossier/`, because it changes
from session to session.

Required contents of the handoff write:
- user's original request and current goal
- current phase / microtask
- files created and modified, with exact paths
- every script generated or used in the session, with path and purpose
- locked decisions and reasons
- what is done, unfinished, or blocked
- verification run, and what could not run with cause
- cumulative script-usage updates
- the concrete next step for the next session

Script retention rule:
- keep all generated or reused scripts as durable assets
- update `.handoff/SCRIPT_USAGE_LOG.md` additively; never reset counts
- read the script log before generating a new helper so existing scripts are reused first

Exception:
- this protocol does not replace the multi-agent work-package workflow
- work-package sessions keep using package-local `04_EVIDENCE_QUEUE.md`,
  `06_EXECUTION_STATE.md`, and the package `HANDOFF_TO_NEXT_SESSION.md`
- reusable scripts from work-package sessions still belong in the cumulative script log

**Why:** normal-session handoff is volatile operational state, not durable preference memory or a
bounded technical dossier. Scripts and their usage history are cumulative assets and should survive
across chats so the same helper is reused instead of regenerated.

**How to apply:** on a meaningful normal-session close, or when the user explicitly says
"chuyển giao dữ liệu", write/update `.handoff/HANDOFF_TO_NEXT_SESSION.md`, archive the previous
snapshot if replacing it, and update `.handoff/SCRIPT_USAGE_LOG.md` with additive counts before the
final reply. Do not move work-package state into `.handoff/`. Link this rule with
[[feedback_persist_memory_before_final_reply]], [[feedback_multi_agent_work_package_workflow]], and
[[project_codebase_memory_repo_local_workflow]].
