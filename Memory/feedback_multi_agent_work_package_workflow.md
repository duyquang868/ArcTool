---
name: feedback_multi_agent_work_package_workflow
description: Use a master-orchestrated multi-agent work package for ArcTool tasks that span 3+ files or require runtime evidence; Agent-tool dispatch is standing-authorized in this workflow.
type: feedback
---

Use the ArcTool work-package workflow as the default execution model when a task spans 3 or more
source files, needs runtime/smoke evidence, or is a roadmap phase / architecture audit / regression
matrix. In that workflow, Claude may spawn subagents without asking for per-dispatch confirmation;
workers stay narrow, read only the shared contract + one task file + the minimum evidence excerpt,
and return the compact schema envelope.

For light dependency-linear micro-chains inside one phase, prefer one workflow script / one master
turn that carries only the immediate upstream result forward. The master must not auto-load the
whole package into startup context; bootstrap only when the session is actually working that
package, start with the minimum file set, and load task/result/evidence files only when the current
dispatch needs them.

**Why:** Multi-file bug fixing and logic tracing can overflow one long chat. The Quick Dimension
package proved that one master plus many small workers, strict write locks, and compact result
reporting scale better and resume cleanly. Lazy bootstrap keeps the master context roughly flat
instead of growing with every task/result/evidence file in the package.

**How to apply:**
1. If the task is small and local, work directly instead of building a package.
2. If the task crosses the activation threshold, create `.claude/workpackages/<slug>/` from the
   template and use the `arctool-work-package` skill.
3. Bootstrap a package only when that session is actively working it. Start with
   `01_SHARED_CONTRACT.md`, `03_TASK_MANIFEST.md`, `05_RESULT_SCHEMA.md`, and
   `06_EXECUTION_STATE.md`; read `04_EVIDENCE_QUEUE.md` only when a ready task needs operator
   evidence or a worker returns `BLOCKED` for evidence.
4. For a light sequential chain, keep the dependency graph in one workflow script and feed each
   worker only the shared contract, one ready task file, the exact upstream result file when it
   matters, and the minimum evidence excerpt.
5. Heavy verification artifacts belong to workers, not the master. The master should route XML,
   journal, screenshot, and other large evidence files by path and let the analysis worker read
   them directly. Master-side reads are reserved for contradiction resolution or final closure
   packaging after the worker envelope is already back.
6. Keep runtime evidence requests master-only and operator-runbook-based.
7. Keep Gemma 4 limited to code-generation tasks inside the package; Claude owns investigation,
   design, review, build verification, and orchestration. See [[feedback_chief_architect_gemma_worker_workflow]].
7. Do not use `isolation: "worktree"` for `result only` tasks. Their only write target is the
   canonical package-local markdown result file, so worktree isolation adds failure risk without
   preventing any real conflict. Reserve worktrees for true parallel source-file write isolation.
8. Never bulk-read all `tasks/`, all `results/`, or the whole evidence queue into startup context.
   The master should consume compact envelopes, not worker content.
