# <PACKAGE TITLE> — TASK MANIFEST

This file defines the execution order, dependency graph, and exclusive write scopes.

---

## Phase 1 — Preflight and scope locks

| Task | Purpose | Depends on | Exclusive write scope |
|---|---|---|---|
| `<T1.1>` | Confirm source owners and no-touch boundaries | — | result only |
| `<T1.2>` | Lock technical invariants with citations | `<T1.1>` | result only |
| `<T1.3>` | Preflight GO / NO-GO gate | `<T1.2>` | result only |

## Phase 2 — Design / patch / verify

| Task | Purpose | Depends on | Exclusive write scope |
|---|---|---|---|
| `<T2.1>` | Design the source change | `<T1.3>` | result only |
| `<T2.2>` | Apply the patch | `<T2.1>` | `<source-file>` |
| `<T2.3>` | Build or statically verify the candidate | `<T2.2>` | result only |
| `<T2.4>` | Prepare operator runbook or downstream handoff | `<T2.3>` | result only |

## Phase 3 — Runtime evidence and verdict

| Task | Purpose | Depends on | Exclusive write scope |
|---|---|---|---|
| `<T3.1>` | Review runtime evidence set A | `<T2.4>` + `<EV-1>` | result only |
| `<T3.2>` | Final verdict / GO / NO-GO | `<T3.1>` | result only |

## Phase 4 — Durable closure

| Task | Purpose | Depends on | Exclusive write scope |
|---|---|---|---|
| `<T4.1>` | Persist durable knowledge in repo-local stores | `<T3.2>` | `Memory/`, `.Dossier/`, `CLAUDE.md`, ADR |
| `<T4.2>` | Draft the final master closure message | `<T4.1>` | result only |

---

## Source-file lock summary

- `<source-file-A>` is edited only by `<task ids in exact order>`.
- `<source-file-B>` is edited only by `<task ids in exact order>`.
- Only the final persistence task may update durable memory / dossier / ADR files.

---

## Result-file convention

Every task writes exactly one detailed result file:

- `.claude/workpackages/<slug>/results/<TASK_ID>_result.md`

For any task whose exclusive write scope is `result only`, that canonical package path is the only
write target and does not justify `isolation: "worktree"`. Worktree isolation is reserved for real
source-file write conflicts, not markdown-only result emission.

The master consumes only the compact envelope unless it must resolve a contradiction.
