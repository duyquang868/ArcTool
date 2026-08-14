# RETIRE QUICK DIMENSION — TASK MANIFEST

This file defines the execution order, dependency graph, and exclusive write scopes.

---

## Phase 1 — Preflight and scope locks

| Task | Purpose | Depends on | Exclusive write scope |
|---|---|---|---|
| `T1.1` | Confirm QD source owners, archive candidates, and no-touch active boundaries | — | result only |
| `T1.2` | Lock retirement invariants, ribbon/API constraints, and preflight GO / NO-GO gate | `T1.1` | result only |

## Phase 2 — Design and apply retirement

| Task | Purpose | Depends on | Exclusive write scope |
|---|---|---|---|
| `T2.1` | Design exact archive layout and active-source cleanup plan | `T1.2` | result only |
| `T2.2` | Remove QD ribbon registrations from active startup | `T2.1` | `ArcTool.Core/App.cs` |
| `T2.3` | Move QD command files into archive area and clean command-surface fallout | `T2.2` | `ArcTool.Core/Commands/QuickDimension*.cs`, `ArcTool.Core/Archive/QuickDimension/Commands/` |
| `T2.4` | Move QD model/service files into archive area and clean project/source fallout | `T2.3` | `ArcTool.Core/Models/QuickDimension*.cs`, `ArcTool.Core/Services/QuickDimension*.cs`, `ArcTool.Core/Archive/QuickDimension/Models/`, `ArcTool.Core/Archive/QuickDimension/Services/` |
| `T2.5` | Review csproj/source layout after archive move and apply minimal cleanup | `T2.4` | `ArcTool.Core/ArcTool.Core.csproj` |
| `T2.6` | Build or statically verify the retired candidate | `T2.5` | result only |

## Phase 3 — Durable closure

| Task | Purpose | Depends on | Exclusive write scope |
|---|---|---|---|
| `T3.1` | Persist durable knowledge in repo-local stores and handoff files | `T2.6` | `Memory/`, `.Dossier/`, `CLAUDE.md`, `.handoff/` |
| `T3.2` | Draft the final master closure message | `T3.1` | result only |

---

## Source-file lock summary

- `ArcTool.Core/App.cs` is edited only by `T2.2`.
- `ArcTool.Core/Commands/QuickDimension*.cs` are edited/moved only by `T2.3`.
- `ArcTool.Core/Models/QuickDimension*.cs` and `ArcTool.Core/Services/QuickDimension*.cs` are edited/moved only by `T2.4`.
- `ArcTool.Core/ArcTool.Core.csproj` is edited only by `T2.5`.
- Durable memory / dossier / handoff files are edited only by `T3.1`.

---

## Result-file convention

Every task writes exactly one detailed result file:

- `.claude/workpackages/retire-quick-dimension/results/<TASK_ID>_result.md`

For any task whose exclusive write scope is `result only`, that canonical package path is the only
write target and does not justify `isolation: "worktree"`. Worktree isolation is reserved for real
source-file write conflicts, not markdown-only result emission.

The master consumes only the compact envelope unless it must resolve a contradiction.
