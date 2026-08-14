# QD BUGFIX — TASK MANIFEST

This file defines the execution order, dependency graph, and exclusive write scopes.

---

## Phase 1 — Preflight and scope locks

| Task | Purpose | Depends on | Exclusive write scope |
|---|---|---|---|
| `T1.1` | Confirm source owners and no-touch boundaries | — | result only |
| `T1.2` | Restate BUG-11 invariant with API citations | `T1.1` | result only |
| `T1.3` | Lock BUG-10 metadata-only scope | `T1.1` | result only |
| `T1.4` | Lock audit/logging fix scope | `T1.1` | result only |
| `T1.5` | Confirm build/runtime constraints | `T1.1` | result only |
| `T1.6` | Check package consistency before edits | `T1.2`,`T1.3`,`T1.4`,`T1.5` | result only |
| `T1.7` | Preflight GO / NO-GO gate | `T1.6` | result only |

## Phase 2 — Instrument the diagnostic fixture

| Task | Purpose | Depends on | Exclusive write scope |
|---|---|---|---|
| `T2.1` | Design minimal instrumentation for reference-owned station proof | `T1.7` | result only |
| `T2.2` | Prepare exact patch plan from T2.1 | `T2.1` | result only |
| `T2.3` | Apply instrumentation patch | `T2.2` | `ArcTool.Core/Services/QuickDimensionDoorWindowCandidateCollector.cs` |
| `T2.4` | Build instrumented debug candidate | `T2.3` | result only |
| `T2.5` | Operator runbook for wall `379469` both shells | `T2.4` | result only |

## Phase 3 — Analyze EV-1 and decide if BUG-11 fix shape is proven

| Task | Purpose | Depends on | Exclusive write scope |
|---|---|---|---|
| `T3.1` | Extract per-reference station evidence from EV-1 | `T2.5` + EV-1 | result only |
| `T3.2` | Compare swapped vs ordered same-type instances | `T3.1` | result only |
| `T3.3` | Finalize the reference-owned station rule | `T3.2` | result only |
| `T3.4` | GO / NO-GO gate for production BUG-11 fix | `T3.3` | result only |

## Phase 4 — Apply BUG-11 production fix

| Task | Purpose | Depends on | Exclusive write scope |
|---|---|---|---|
| `T4.1` | Design the production source change | `T3.4` | result only |
| `T4.2` | Apply BUG-11 fix | `T4.1` | `ArcTool.Core/Services/QuickDimensionDoorWindowCandidateCollector.cs` |
| `T4.3` | Build BUG-11-fixed candidate | `T4.2` | result only |
| `T4.4` | Static regression review before secondary fixes | `T4.3` | result only |

## Phase 5 — Secondary fixes and regression candidate

| Task | Purpose | Depends on | Exclusive write scope |
|---|---|---|---|
| `T5.1` | Apply BUG-10 metadata fix | `T4.4` | `ArcTool.Core/Services/QuickDimensionDoorWindowCandidateCollector.cs` |
| `T5.2` | Fix `actualSegmentCount` logging | `T5.1` | `ArcTool.Core/Services/QuickDimensionReadOnlyXmlLogService.cs` |
| `T5.3` | Add `valueSource` to segment audit | `T5.2` | `ArcTool.Core/Services/QuickDimensionReadOnlyXmlLogService.cs` |
| `T5.4` | Build full regression candidate | `T5.3` | result only |
| `T5.5` | Prepare operator-ready regression handoff | `T5.4` | result only |

## Phase 6 — Operator regression and verdict

| Task | Purpose | Depends on | Exclusive write scope |
|---|---|---|---|
| `T6.1` | Operator runbook for regression matrix | `T5.5` | result only |
| `T6.2` | Review wall `379467` evidence | `T6.1` + EV-2 | result only |
| `T6.3` | Review wall `379469` evidence | `T6.1` + EV-2 | result only |
| `T6.4` | Review wall `379470` evidence | `T6.1` + EV-2 | result only |
| `T6.5` | Operator runbook for reopen validation (rollback deferred 2026-08-04) | `T5.5` | result only |
| `T6.6` | Final regression / reopen verdict | `T6.2`,`T6.3`,`T6.4`,`T6.5` + EV-3 | result only |
| `T6.7` | Record deferred rollback-validation track as a standalone future task | `T6.5` | `.Dossier/Quick Dimension - Deferred Rollback Validation Task.md` |

## Phase 7 — Durable closure

| Task | Purpose | Depends on | Exclusive write scope |
|---|---|---|---|
| `T7.1` | Persist durable knowledge in repo-local stores | `T6.6`,`T6.7` | `Memory/`, `.Dossier/` (except the T6.7 file), `CLAUDE.md`, ADR |
| `T7.2` | Draft the final master closure message | `T7.1` | result only |

### Scope amendment — 2026-08-04

Forced-rollback validation was removed from this mission's closure gates (operator-approved;
rationale in `01_SHARED_CONTRACT.md` section 8). `T6.5` therefore delivers a reopen-only
runbook, `T6.6` renders a regression + reopen verdict, and the deferred rollback work is
recorded once by `T6.7` as an independent future task that this package does not execute.

---

## Source-file lock summary

- `QuickDimensionDoorWindowCandidateCollector.cs` is edited only by `T2.3`, `T4.2`, `T5.1`, in that exact order.
- `QuickDimensionReadOnlyXmlLogService.cs` is edited only by `T5.2`, `T5.3`, in that exact order.
- No task edits `QuickDimensionChainCreationService.cs`. This is why rollback behavior is
  byte-identical to the pre-fix build and is not a regression surface for this mission.
- `.Dossier/` durable writes are serialized `T6.7` → `T7.1`; `T6.7` owns only its own new file
  and `T7.1` owns every other durable record. They must never run concurrently.
- Only `T6.7` and `T7.1` may update durable memory / dossier / ADR files.

---

## Result-file convention

Every task writes exactly one detailed result file:

- `.claude/quick-dimension-bugfix/results/<TASK_ID>_result.md`

The master consumes only the compact envelope unless it must resolve a contradiction.