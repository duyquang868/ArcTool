# Detailed Technical Dossier - Multi-Agent Work Package Workflow

## 1. Purpose

This dossier records the permanent ArcTool workflow for multi-file, context-heavy work that is
safer and more reliable as one master orchestrator plus many small workers than as one long chat.

The model exists to prevent context overflow during multi-file bug fixing, cross-file logic tracing,
runtime evidence review, and roadmap-scale implementation work.

---

## 2. Activation threshold

Use a work package when any of the following is true:

- the task touches 3 or more source files;
- the task needs runtime or smoke evidence;
- the task is a roadmap phase, architecture audit, or regression matrix.

Do not package trivial work. Single-file tweaks, short answers, and small direct fixes should stay
in the normal direct workflow.

---

## 3. Package anatomy

Each package lives under `.claude/workpackages/<slug>/` and is built from the generic template.

### 3.1 Shared contract — `01_SHARED_CONTRACT.md`

Defines the mission, hard invariants, domain model, source ownership map, fixtures, build command,
and whole-package acceptance gates.

This file is the package constitution. Every worker reads it first.

### 3.2 Master orchestrator — `02_MASTER_ORCHESTRATOR.md`

Defines the master startup read set, dispatch rules, phase gates, evidence routing, write-lock
policy, and a fresh-chat bootstrap prompt.

This file is for the master chat, not the workers.

### 3.3 Task manifest — `03_TASK_MANIFEST.md`

Defines the dependency graph, execution order, and exclusive write scopes.

One task row corresponds to one worker task file. This is what keeps concurrency safe.

### 3.4 Evidence queue — `04_EVIDENCE_QUEUE.md`

Defines all operator-evidence requests in one place.

Only the master asks the human for runtime evidence. Workers raise `BLOCKED` with a precise
`blocker:` line and stop there.

### 3.5 Result schema — `05_RESULT_SCHEMA.md`

Defines the two-output rule:

- one detailed result file under `results/`;
- one compact envelope returned to the master.

The compact envelope is the anti-overflow mechanism. It must stay short and uniform.

### 3.6 Execution state — `06_EXECUTION_STATE.md`

Defines the package status table used for resume, handoff, and fresh-chat continuation.

The master updates it after every worker result.

---

## 4. Master / worker contract

The master may:

- read package-level files;
- choose ready tasks from the manifest;
- dispatch one worker per task file;
- ask the user for runtime evidence through the evidence queue;
- update execution state and durable closure files.

A worker may:

- read the shared contract;
- read exactly one task file;
- read only the minimum evidence excerpt supplied by the master;
- write only the files named in its `write_scope`.

A worker may not:

- read `CLAUDE.md` in full by default;
- ask the user for evidence directly;
- paste large source files, XML logs, or screenshots into its reply;
- take extra writable files outside its declared scope.

---

## 5. Write-lock discipline

Two workers must never own the same writable source file at once.

The manifest must serialize every shared source file explicitly. A correct package records the exact
edit order per file and places build tasks after the preceding write task.

This is the main reason the model scales safely beyond one worker.

---

## 6. Runtime boundary

Revit runtime remains operator-controlled.

A package may prepare instrumentation, builds, static analysis, and exact operator runbooks, but it
must never launch Revit, open a model, invoke Revit MCP, or run a smoke test unless the user
explicitly asks for that runtime action in the current chat.

Operator-returned journals are valid independent evidence and should be correlated with XML,
screenshots, and source changes.

---

## 7. Gemma 4 boundary

Gemma 4 remains part of the ArcTool ecosystem but is narrowed to code-generation tasks only.

Claude remains responsible for:

- architecture;
- cross-file reasoning;
- package orchestration;
- evidence interpretation;
- review correctness;
- build verification;
- durable persistence decisions.

Gemma is not the orchestrator and is not the runtime investigator.

---

## 8. Lessons proven by the Quick Dimension package

The Quick Dimension package at `.claude/quick-dimension-bugfix/` proved the model in real work.

### 8.1 Sequential write locks prevented source collisions

`ArcTool.Core/Services/QuickDimensionDoorWindowCandidateCollector.cs` was edited only in the exact
sequence `T2.3` → `T4.2` → `T5.1`.

`ArcTool.Core/Services/QuickDimensionReadOnlyXmlLogService.cs` was edited only in the exact
sequence `T5.2` → `T5.3`.

This prevented overlapping patches on the same source owner.

### 8.2 `NO_GO` is a valid result

Phase gates such as `T1.7` and `T3.4` showed that a worker can conclude "do not proceed" without
failing the workflow. The package model must preserve this semantic explicitly.

### 8.3 `BLOCKED` is normal when runtime evidence is missing

`T6.5` proved that runtime-evidence absence is a normal package state, not a failure. The correct
response is a precise blocker plus an operator runbook entry, not speculation.

### 8.4 Compact envelopes prevent context blow-up

The result schema forced each worker to return only a short envelope while keeping full detail in
its result file. This is what made long multi-phase work resumable.

---

## 9. Relationship to other durable workflow layers

- `CLAUDE.md` keeps only the short operating rules: activation threshold, standing Agent-tool
  authorization, and worker discipline.
- `.Dossier` keeps the long rationale and anatomy of the model.
- `Memory/` keeps user/project handling preferences about when and how to use the model.
- ADR records the durable architectural decision that this is the default execution model for
  multi-file or runtime-heavy ArcTool work.

---

## 10. Closure rule

When a package reaches a meaningful boundary, durable persistence must be written before the final
reply of that turn.

Only after durable files are safe may the user be offered the final optional `index_repository`
re-index step.
