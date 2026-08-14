# QD BUGFIX — MASTER ORCHESTRATOR

Use this file in the fresh master chat that coordinates the bugfix across many small workers.

---

## Startup

1. Read only these package files first:
   - `01_SHARED_CONTRACT.md`
   - `03_TASK_MANIFEST.md`
   - `04_EVIDENCE_QUEUE.md`
   - `05_RESULT_SCHEMA.md`
   - `06_EXECUTION_STATE.md`
2. Do **not** read `CLAUDE.md` in full unless a task explicitly requires a durable-persistence edit.
3. Do **not** paste full source files, full XML files, or screenshots into worker prompts.
4. Give each worker only:
   - the shared contract,
   - one task file,
   - the exact evidence excerpt that task needs.

---

## Dispatch rules

- One worker = one task file.
- Do not start a task until every dependency in `03_TASK_MANIFEST.md` is `PASS`.
- Do not run two write tasks at the same time when they share an exclusive source file.
- Workers return only the compact envelope from `05_RESULT_SCHEMA.md`.
- Read a worker's detailed result file only to resolve a contradiction or prepare closure.
- Only the master updates `04_EVIDENCE_QUEUE.md` and `06_EXECUTION_STATE.md`.
- Only the master asks the human for evidence.

---

## Phase gates

- Stop the whole flow if `T1.7` returns `NO_GO`.
- Do not start Phase 4 until `T3.4` returns `PASS`.
- Do not start Phase 6 until `T5.5` returns `PASS`.
- Do not start Phase 7 until `T6.6` returns `PASS`.

---

## Evidence routing

When a worker returns `BLOCKED` with an operator-evidence need:

1. Check whether the needed runbook already exists in `04_EVIDENCE_QUEUE.md`.
2. Ask the human only for the exact run named in that queue item.
3. When evidence arrives, forward only the minimum excerpt required by the blocked task.
4. Record the handoff in `04_EVIDENCE_QUEUE.md` and update `06_EXECUTION_STATE.md`.

---

## Write-lock policy

Serialize these files strictly:

- `ArcTool.Core/Services/QuickDimensionDoorWindowCandidateCollector.cs`
  - `T2.3` → `T4.2` → `T5.1`
- `ArcTool.Core/Services/QuickDimensionReadOnlyXmlLogService.cs`
  - `T5.2` → `T5.3`
- Durable project memory / dossier files
  - `T7.1` only

Build tasks may run after the preceding write task finishes.

---

## Fresh-chat bootstrap prompt

```text
Read `.claude/quick-dimension-bugfix/01_SHARED_CONTRACT.md`,
`.claude/quick-dimension-bugfix/03_TASK_MANIFEST.md`,
`.claude/quick-dimension-bugfix/04_EVIDENCE_QUEUE.md`,
`.claude/quick-dimension-bugfix/05_RESULT_SCHEMA.md`, and
`.claude/quick-dimension-bugfix/06_EXECUTION_STATE.md`.

Act as the master orchestrator for this Quick Dimension bugfix package.
Use one worker per ready task. Respect all dependencies and exclusive write scopes.
Return only compact envelopes from workers. Ask the human for runtime evidence only through the runbooks listed in `04_EVIDENCE_QUEUE.md`.
Update `06_EXECUTION_STATE.md` after every worker result.
Stop on any `NO_GO` gate.
```