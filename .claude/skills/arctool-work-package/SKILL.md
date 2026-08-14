---
name: arctool-work-package
description: Multi-agent micro-task work package workflow for ArcTool. Use when a task spans 3 or more source files, needs runtime/smoke investigation, or is a roadmap phase or architecture audit — multi-file bug fixes, cross-file logic tracing, impact analysis, regression matrices, context-overflow-prone investigations. Covers building the work package scaffold, dispatching one subagent per task file, exclusive write scopes, evidence routing, and compact result envelopes.
---

# ArcTool Multi-Agent Work Package

## When this is mandatory

Build a work package when ANY of these is true:

- the task touches **3 or more source files**;
- the task needs **runtime / smoke investigation** (Revit evidence, journals, XML logs);
- the task is a **roadmap phase**, **architecture audit**, or **regression matrix**.

Do NOT build a package for a single-file edit, a one-line fix, a rename, or a
question answerable from the knowledge graph. Direct work is correct there;
wrapping trivial work in six scaffold files is waste.

## Standing authorization

The user has granted **standing authorization** to spawn subagents via the Agent
tool for this workflow. Do not ask for per-dispatch confirmation. See `CLAUDE.md`
`Mandatory editing rules`.

## Procedure

### 1. Build the package

1. Pick a slug: `.claude/workpackages/<slug>/` (kebab-case, names the mission, not the date).
2. Copy all six files from `.claude/workpackages/_TEMPLATE/`, create an empty `results/`.
3. Fill `01_SHARED_CONTRACT.md`: mission, mission-specific invariants, domain model,
   **source ownership map with verified line ranges**, fixtures. Resolve the ownership
   map with `codebase-memory-mcp` (`search_graph`, `trace_path`, `get_code_snippet`) —
   not by reading whole files.
4. Fill `03_TASK_MANIFEST.md`: phases, one row per task, dependencies, and one
   **exclusive write scope** per task. Most tasks write `result only`.
5. Write one task file per row under `tasks/`. A task file states: objective, the exact
   inputs it may read, its write scope, its acceptance condition, and what the downstream
   task needs from it.

### 2. Dispatch

- **One worker = one task file.** Never merge two task files into one subagent.
- Worker model is pinned to **Claude Sonnet 5**: pass `model: "sonnet"` on every Agent-tool dispatch, or `{model: 'sonnet'}` in a workflow `agent()` call. Deviate only when the package explicitly documents a justified exception.
- For a strictly sequential, light micro-chain inside one phase, prefer **one workflow script /
  one master turn** over many separate master turns: keep the dependency chain in script code and
  dispatch the next worker only after the previous one returns `PASS`.
- A worker receives exactly: `01_SHARED_CONTRACT.md`, its own task file, the immediate upstream
  result file when the chain depends on it, and the minimum evidence excerpt. Nothing else.
- A worker must **not** read `CLAUDE.md` in full, and must not read whole source files
  unless its task file names them.
- Do not start a task until every dependency in the manifest is `PASS`.
- Do not run two write tasks concurrently when they share a source file.
- Workers return **only** the compact envelope from `05_RESULT_SCHEMA.md`. Detailed
  findings go to `results/<TASK_ID>_result.md`.
- For `result only` tasks, do **not** dispatch with `isolation: "worktree"`. Their write
  scope is only the package-local markdown result file, so there is no source-file conflict to
  isolate and worktree dispatch can orphan the result outside the canonical package `results/`
  directory.
- Use `isolation: "worktree"` only when parallel workers must edit overlapping source files or
  otherwise need real filesystem isolation beyond a result markdown write.
- Read a worker's result file only to resolve a contradiction, feed the next sequential task, or
  prepare closure.

### 2b. Master startup context budget

- The master does **not** auto-load a package into context. Bootstrap only when the current
  session is actually working that package.
- Minimum bootstrap set: `01_SHARED_CONTRACT.md`, `03_TASK_MANIFEST.md`, `05_RESULT_SCHEMA.md`,
  `06_EXECUTION_STATE.md`. Read `04_EVIDENCE_QUEUE.md` only when a ready task needs operator
  evidence or a worker returns `BLOCKED` for evidence.
- Never bulk-read all `tasks/`, all `results/`, or the whole evidence queue at session start.
  Load the ready task file and the exact upstream result only at dispatch time.
- The master's context should stay roughly flat regardless of how many tasks the phase contains,
  because it consumes compact envelopes, not worker content.
- The master must not read heavy verification artifacts itself. Full XML logs, journals, screenshots,
  large tables, and other evidence payloads are worker inputs, not master context.
- When the human supplies evidence paths, the master records the paths, updates `04_EVIDENCE_QUEUE.md`,
  and dispatches the analysis worker with those paths or a tiny routing excerpt. The worker reads the
  heavy artifact and returns a compact envelope.
- If a heavy artifact is too large to summarize safely in one short excerpt, forward the path only.
  Do not pre-read the file in the master just to prepare the worker prompt.
- The master may read a heavy evidence file only for closure packaging or contradiction resolution
  after the worker result is already back and the compact envelope is insufficient.

### 3. Write-lock

Two workers must never hold the same source file in `write_scope` at the same time.
Record the serialization order explicitly in the manifest's lock summary, e.g.
`Foo.cs` is edited only by `T2.3` → `T4.2` → `T5.1`, in that exact order. Build tasks
run after the preceding write task finishes.

### 4. Runtime boundary

Revit runtime is operator-owned. No worker may launch Revit, open an `.rvt`, call a
Revit MCP tool, or run a smoke test. A worker's runtime proof stops at a **written
operator runbook**; the user runs it and returns evidence.

Only the master asks the user for evidence, and only through an `EV-<n>` item in
`04_EVIDENCE_QUEUE.md`. A worker that needs runtime evidence returns `BLOCKED` with a
`blocker:` line — it never asks the user directly.

### 5. Status semantics

- `PASS` — objective met; downstream may start.
- `BLOCKED` — missing evidence, missing upstream decision, or a denied tool call. Name
  exactly what is missing and who supplies it. Never guess to force a `PASS`.
- `NO_GO` — a gate task concluded "do not proceed". This is a valid, useful outcome.

### 6. Bookkeeping

Only the master updates `04_EVIDENCE_QUEUE.md` and `06_EXECUTION_STATE.md`. Update
execution state after **every** worker result — that file is what lets a fresh chat
resume the package without this conversation.

### 7. Where Gemma 4 fits

Gemma 4 (LM Studio MCP) is the worker for **code-generation tasks only**. Investigation,
spec design, review, build verification, evidence analysis, and orchestration stay with
Claude subagents. See `Memory/feedback_chief_architect_gemma_worker_workflow.md`.

### 8. Closing

When the package reaches durable closure, follow the `arctool-session-learn` skill to
classify and persist the outcome. Persist durable files **before** the final reply;
re-index is the last optional user-directed step.

## Anti-patterns

- Pasting a whole source file, whole XML log, or screenshot into a worker prompt.
- Letting a worker ask the user for runtime evidence.
- Building a package for a one-line fix.
- Two workers editing the same source file in the same wave.
- A worker returning its findings in the reply instead of its result file.
- Duplicating the mission's technical detail into `CLAUDE.md` — long-form context belongs in
  `.Dossier/Detailed Technical Dossier - Multi-Agent Work Package Workflow.md`.

## Reference

- Template: `.claude/workpackages/_TEMPLATE/`
- Worked example (live, do not relocate): `.claude/quick-dimension-bugfix/`
- Rationale and lessons: `.Dossier/Detailed Technical Dossier - Multi-Agent Work Package Workflow.md`
