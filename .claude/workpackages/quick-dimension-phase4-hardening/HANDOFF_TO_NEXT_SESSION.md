# Quick Dimension Phase 4 Hardening — Handoff to Next Session

## 1. Current package state
- Package path: `.claude/workpackages/quick-dimension-phase4-hardening/`
- Status: **IN PROGRESS**
- Scaffold status: **FULLY MATERIALIZED THROUGH `T7.2`**
- Worker dispatch status: **`T1.1`, `T1.2`, `T1.3`, `T1.4`, `T1.5`, and `T1.6` dispatched and closed `PASS`**
- Next dispatch target: `T2.1` — clean-fixture spec (`tasks/T2.1_clean_fixture_spec.md`)
- Results directory status: `results/T1.1_result.md` through `results/T1.6_result.md` exist; no later task result files exist yet

## 2. Closed facts this package must preserve
- BUG-10 is closed and must stay closed.
- BUG-11 is closed and must stay closed.
- EV-2 regression matrix already passed on walls `379467`, `379469`, and `379470`, both shells.
- EV-3 reopen persistence already passed on dimensions `385355`, `385356`, `385632`, `385584`, `385719`, and `385720`.
- Forced rollback validation was intentionally deferred to `.Dossier/Quick Dimension - Deferred Rollback Validation Task.md` because no task in this package may edit `ArcTool.Core/Services/QuickDimensionChainCreationService.cs`.

## 3. Hard invariants
- Revit runtime is operator-owned. Claude must not launch Revit, open an `.rvt`, invoke Revit MCP, or run smoke tests.
- Do not touch `ArcTool.Core/Services/QuickDimensionChainCreationService.cs` in this package.
- Do not weaken the Quick Dimension audit. `Exact`, complete `Reversed`, and `Mismatch` stay the only accepted order relations. Do not whitelist `LocalPairSwap`.
- Do not reopen BUG-10 or BUG-11 fix shapes while doing performance or regression work.
- Grid remains excluded from the wall-axis production flow. Session 4.3 is safe-failure verification only.
- ADR safety rule: routine closure, handoff, scaffolding, and forensic-note persistence must not write the ADR store.

## 4. Package contents now present
- Core package files: `01_SHARED_CONTRACT.md` through `06_EXECUTION_STATE.md`
- Task files: `T1.1` through `T7.2`
- Package-local handoff: this file
- Evidence queue: `EV-1` through `EV-6`, all still pending
- Execution state: `T1.1` `PASS`, `T1.2` `PASS`, `T1.3` `PASS`, `T1.4` `PASS`, `T1.5` `PASS`, `T1.6` `PASS`; `T2.1` … `T7.2` still `PENDING`

## 4b. Locked baseline inherited from `T1.2`
`T1.4` and every later task must treat these as fixed input and must not reinterpret them:
- BUG-10 closed; a future mismatch is a **fresh** regression, never evidence BUG-10 was still open.
- BUG-11 closed; a future mirror/flip anomaly is a fresh regression or a newly scoped unsupported case.
- EV-2 six-run matrix `Exact` and EV-3 reopen `PASS` are mission-entry evidence, not provisional claims.
- Rollback deferral is a **separate task boundary**, never an internal Phase 4 blocker and never a
  reason to edit protected creation logic.
- Grid exclusion in wall-axis mode is a **present-source implementation boundary**, not a future option.
- The package starts from a working production feature; it is not resuming BUG-10/BUG-11 repair work.
- Implementation boundaries (what may change) stay distinct from evidence boundaries (what is closed).

## 4c. Grid scope rule locked by `T1.3`
`T1.4`, `T4.1`, `T4.2`, and `T4.3` must treat this as fixed input:
- **CANONICAL RULE:** Session 4.3 is a Grid safe-failure matrix only: in wall-axis mode every Grid
  variant is outside the candidate flow by unconditional pre-wall-resolution exclusion, so downstream
  tasks may verify only unchanged behavior and honest unsupported outcomes, never Grid support expansion.
- Source basis: `CollectWallAxisCandidates` emits the `Grid` disabled diagnostic as its first statement,
  before selected-wall resolution, ungated by `QuickDimensionOptions.IncludeGrids`.
- Session 4.3 must be designed as an **A/B negative control** (same wall case with and without a Grid
  variant), never as support probing. Straight/cropped/hidden/arc variants cannot influence wall-axis
  behavior, so the runbook must not ask whether Grids get dimensioned or partly work.
- Grid cases are classified **Unsupported-by-design** by default; any Grid-caused behavioral difference
  is a **Defect**, not partial support and not an invitation to widen scope.
- FORBIDDEN downstream: widening source support, enabling Grid candidates, editing
  `QuickDimensionGridCandidateCollector.cs`, treating `IncludeGrids` as relevant to wall-axis mode, or
  using Grid-variant runs to reopen BUG-10/BUG-11.
- Open observation, no source change proposed: the diagnostic text says "disabled by Quick Dimension
  options" while the wall-axis exclusion is actually unconditional. Wording only; do not act on it in
  this package without an explicit task.

## 5. Dispatch order reminder
- Next dispatch is `T1.6`; `T1.1`, `T1.2`, `T1.3`, `T1.4`, and `T1.5` are closed.
- Do not dispatch a task before all dependencies pass.
- One worker owns exactly one task file.
- The master alone updates `04_EVIDENCE_QUEUE.md` and `06_EXECUTION_STATE.md`.
- Respect the manifest lock summary before any write task is dispatched.

## 6. What remains unfinished outside the package
- The global ADR safety-rule lock is already persisted outside package execution (`CLAUDE.md`, `Memory/feedback_adr_store_update_lock.md`, `Memory/MEMORY.md`, and the roadmap pointer added on 2026-08-05). Package task `T7.1` still stays pending because final Phase 4 outcome does not exist yet.
- Package-local closure audit is not written yet; that belongs to `T7.2`.
- No runtime evidence has been requested from the operator yet.
- No Phase 5 or Phase 6 source work has been executed yet.
- No optional codebase-memory re-index has been run.

## 7. Resume rule for the next chat
- Treat this package as the source of truth.
- Read `06_EXECUTION_STATE.md` first, then dispatch `T1.6` with the contract, its task file, and
  `results/T1.5_result.md` as the upstream input.
- One micro-task per chat under the current operator constraint; do not batch `T1.5` with later tasks.
- Do not ask the operator for runtime evidence until a runbook task writes the matching `EV-<n>` request.
- Do not touch ADR during routine package execution.
