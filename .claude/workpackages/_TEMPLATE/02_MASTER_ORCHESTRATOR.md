# <PACKAGE TITLE> — MASTER ORCHESTRATOR

Use this file in the fresh master chat that coordinates the package across many small workers.

---

## Startup

1. Bootstrap a package only when the current session is actually working that package.
2. Read only this minimum set first:
   - `01_SHARED_CONTRACT.md`
   - `03_TASK_MANIFEST.md`
   - `05_RESULT_SCHEMA.md`
   - `06_EXECUTION_STATE.md`
3. Read `04_EVIDENCE_QUEUE.md` only when a ready task depends on operator evidence or when a
   worker returns `BLOCKED` for evidence.
4. Do **not** read `CLAUDE.md` in full unless a task explicitly requires a durable-persistence edit.
5. The master does **not** read heavy verification artifacts during normal dispatch. Full XML logs,
   journals, screenshots, and other large evidence files are read by the downstream analysis worker,
   not by the master.
6. Do **not** paste full source files, full XML files, or screenshots into worker prompts.
7. Give each worker only:
   - the shared contract,
   - one ready task file,
   - the immediate upstream result file when that dependency matters,
   - the exact evidence excerpt that task needs.
7. Do **not** auto-load the full package into startup context: no bulk read of all task files, all
   result files, or the whole evidence queue.

---

## Dispatch rules

- One worker = one task file.
- Worker model is pinned to Claude Sonnet 5: dispatch with `model: "sonnet"` on the Agent tool, or `{model: 'sonnet'}` in a workflow `agent()` call. Use another model only when a task explicitly requires it and the package records that exception.
- For a sequential micro-chain (`T1.1 -> T1.2 -> ...`) that is light and dependency-linear,
  prefer one workflow script / one master turn that loops over the chain and carries only the
  prior result path forward.
- Do not start a task until every dependency in `03_TASK_MANIFEST.md` is `PASS`.
- Do not run two write tasks at the same time when they share an exclusive source file.
- Workers return only the compact envelope from `05_RESULT_SCHEMA.md`.
- For tasks whose exclusive write scope is `result only`, dispatch without `isolation: "worktree"`.
  A result markdown file does not justify filesystem isolation, and worktree dispatch can send the
  file into a non-canonical worktree-local `results/` path that the master is not watching.
- Reserve `isolation: "worktree"` for tasks that really edit source files in parallel and need
  collision-free filesystem writes.
- Read a worker's detailed result file only to resolve a contradiction, feed the next sequential
  task, or prepare closure.
- Only the master updates `04_EVIDENCE_QUEUE.md` and `06_EXECUTION_STATE.md`.
- Only the master asks the human for evidence.

---

## Phase gates

Record the package's stop/go gates here.

Template examples:

- Stop the whole flow if `<gate task>` returns `NO_GO`.
- Do not start `<phase>` until `<gate task>` returns `PASS`.
- Do not start durable closure until the final verdict task returns `PASS`.

Delete example lines and replace them with real task ids when the package is created.

---

## Evidence routing

When a worker returns `BLOCKED` with an operator-evidence need:

1. Check whether the needed runbook already exists in `04_EVIDENCE_QUEUE.md`.
2. Ask the human only for the exact run named in that queue item.
3. When evidence arrives, record the evidence paths first and forward only the minimum routing excerpt required by the blocked task.
4. If the blocked task needs the full XML, journal, or screenshot, pass the artifact path to the worker and let the worker read it.
5. Record the handoff in `04_EVIDENCE_QUEUE.md` and update `06_EXECUTION_STATE.md`.

---

## Write-lock policy

Serialize the package's shared write targets strictly.

Template:

- `<path>`
  - `<task A>` → `<task B>` → `<task C>`
- Durable project memory / dossier files
  - final persistence task only

Build tasks may run after the preceding write task finishes.

---

## Fresh-chat bootstrap prompt

```text
Read `.claude/workpackages/<slug>/01_SHARED_CONTRACT.md`,
`.claude/workpackages/<slug>/03_TASK_MANIFEST.md`,
`.claude/workpackages/<slug>/05_RESULT_SCHEMA.md`, and
`.claude/workpackages/<slug>/06_EXECUTION_STATE.md`.
Read `04_EVIDENCE_QUEUE.md` only if a ready task needs operator evidence.

Act as the master orchestrator for this work package.
Do not auto-load the full package into startup context.
Do not read heavy verification artifacts; route them to workers by path.
Dispatch every worker with `model: "sonnet"` (Claude Sonnet 5).
Use one worker per ready task, or one workflow-script loop for a light sequential micro-chain.
Respect all dependencies and exclusive write scopes.
Return only compact envelopes from workers. Ask the human for runtime evidence only through the runbooks listed in `04_EVIDENCE_QUEUE.md`.
Update `06_EXECUTION_STATE.md` after every worker result.
Stop on any `NO_GO` gate.
```
