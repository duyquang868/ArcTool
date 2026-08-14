# ArcTool Work Packages

Use `.claude/workpackages/_TEMPLATE/` as the starting scaffold for any ArcTool task that
spans multiple files, needs runtime evidence, or would otherwise overflow one chat.

## When to create a package

Create a package when the task:

- touches **3 or more source files**;
- needs **runtime / smoke investigation**;
- is a **roadmap phase**, **architecture audit**, or **regression matrix**.

Do not create one for a single-file tweak or a small direct answer.

## Naming

Create the package at `.claude/workpackages/<slug>/`.

- Use kebab-case.
- Name the mission, not the date.
- Good: `filter-manager-copy-paste`, `coordinate-updater-regression`
- Bad: `session-08-04`, `misc-fixes`

## Required files

Copy every file from `_TEMPLATE/` and create:

- `tasks/`
- `results/`

The six scaffold files are:

1. `01_SHARED_CONTRACT.md`
2. `02_MASTER_ORCHESTRATOR.md`
3. `03_TASK_MANIFEST.md`
4. `04_EVIDENCE_QUEUE.md`
5. `05_RESULT_SCHEMA.md`
6. `06_EXECUTION_STATE.md`

## Live example

`.claude/quick-dimension-bugfix/` is the worked example and remains the active historical
reference. Do not rename or relocate it; many result files and handoff pointers already depend
on that path.

## Expected flow

1. Copy `_TEMPLATE/` to a new slug.
2. Fill the contract with mission-specific invariants and verified source ownership.
3. Define one task row per micro-task in the manifest.
4. Write one worker task file per row under `tasks/`.
5. Let the master orchestrator dispatch one worker per ready task.
6. Keep runtime evidence requests in `04_EVIDENCE_QUEUE.md`.
7. Keep worker status in `06_EXECUTION_STATE.md`.
8. Persist durable lessons before the final reply; re-index only if the user chooses.
