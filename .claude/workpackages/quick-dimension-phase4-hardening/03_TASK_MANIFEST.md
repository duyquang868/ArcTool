# QD PHASE 4 HARDENING — TASK MANIFEST

The master owns this file.
Workers never edit it.

---

## Phase 1 — Preflight and scope locks

| Task | Purpose | Depends on | Exclusive write scope |
|---|---|---|---|
| `T1.1` | Lock the verified source ownership map, no-touch set, and line ranges into a durable result | — | `result only` |
| `T1.2` | Lock the evidence-backed baseline: BUG-10/BUG-11 closed, rollback deferred, Grid disabled in wall-axis flow | `<T1.1>` | `result only` |
| `T1.3` | Resolve the Session 4.3 contradiction formally: Grid matrix = safe-failure, not support expansion | `<T1.2>` | `result only` |
| `T1.4` | Define the package acceptance vocabulary: Supported vs Unsupported-by-design vs Defect | `<T1.3>` | `result only` |
| `T1.5` | Verify the approved VS MSBuild path and record known-benign build noise to ignore | `<T1.4>` | `result only` |
| `T1.6` | Perform package consistency review and authorize Session 4.1 work | `<T1.5>` | `result only` |

## Phase 2 — Session 4.1 clean-model acceptance

| Task | Purpose | Depends on | Exclusive write scope |
|---|---|---|---|
| `T2.1` | Author the clean-fixture spec the operator must build | `<T1.6>` | `result only` |
| `T2.2` | Write the pre-committed analytic oracle and expected dimension values for that fixture | `<T2.1>` | `result only` |
| `T2.3` | Write the operator runbook for EV-1 clean-fixture execution | `<T2.2>` | `04_EVIDENCE_QUEUE.md`, `result only` |
| `T2.4` | Judge EV-1 evidence against the oracle and classify the clean fixture outcome | `<T2.3>` | `result only` |
| `T2.5` | Publish Session 4.1 verdict and either authorize Session 4.2 or stop on NO_GO | `<T2.4>` | `result only` |

## Phase 3 — Session 4.2 wall + Door/Window complexity matrix

| Task | Purpose | Depends on | Exclusive write scope |
|---|---|---|---|
| `T3.1` | Author the case matrix for wall/opening complexity, including empty, dense, flush, close-spaced, mirrored, and mid-run cases | `<T2.5>` | `result only` |
| `T3.2` | Write static predictions for each case before runtime evidence exists | `<T3.1>` | `result only` |
| `T3.3` | Design the explicit mirror/flip observation probe and success criteria | `<T3.2>` | `result only` |
| `T3.4` | Write the operator runbook for EV-2 wall/opening complexity execution | `<T3.3>` | `04_EVIDENCE_QUEUE.md`, `result only` |
| `T3.5` | Judge EV-2 evidence for the simple and dense supported cases | `<T3.4>` | `result only` |
| `T3.6` | Judge EV-2 evidence for close-spaced and end-flush edge cases | `<T3.4>` | `result only` |
| `T3.7` | Judge EV-2 evidence for mirrored/flipped and mid-run joint cases | `<T3.4>` | `result only` |
| `T3.8` | Publish the full Session 4.2 verdict and classify every case | `<T3.5>`, `<T3.6>`, `<T3.7>` | `result only` |

## Phase 4 — Session 4.3 grid safe-failure matrix

| Task | Purpose | Depends on | Exclusive write scope |
|---|---|---|---|
| `T4.1` | Author the Grid variant matrix: straight, cropped, hidden, and arc | `<T3.8>` | `result only` |
| `T4.2` | Write the safe-failure predictions and honest diagnostic expectations for each Grid variant | `<T4.1>` | `result only` |
| `T4.3` | Write the operator runbook for EV-3 Grid safe-failure execution | `<T4.2>` | `04_EVIDENCE_QUEUE.md`, `result only` |
| `T4.4` | Judge EV-3 evidence per Grid variant, including exact no-dimension / dialog / XML expectations | `<T4.3>` | `result only` |
| `T4.5` | Publish Session 4.3 verdict and either authorize performance work or stop on NO_GO | `<T4.4>` | `result only` |

## Phase 5 — Session 4.4 performance baseline and conditional optimization

| Task | Purpose | Depends on | Exclusive write scope |
|---|---|---|---|
| `T5.1` | Design behaviour-neutral read-only timing instrumentation and where it will emit | `<T4.5>` | `result only` |
| `T5.2` | Apply instrumentation to the read-only timing path without changing behaviour | `<T5.1>` | `ArcTool.Core/Services/QuickDimensionReadOnlyEngine.cs`, `ArcTool.Core/Services/QuickDimensionReadOnlyXmlLogService.cs`, `result only` |
| `T5.3` | Build the instrumented code with the approved VS MSBuild path | `<T5.2>` | `result only` |
| `T5.4` | Write the operator runbook for EV-4 performance baseline collection | `<T5.3>` | `04_EVIDENCE_QUEUE.md`, `result only` |
| `T5.5` | Analyze EV-4 timing evidence and identify the single hottest collector path, if any | `<T5.4>` | `result only` |
| `T5.6` | Decide GO / NO_GO for optimization based on measured hotspot strength and regression risk | `<T5.5>` | `result only` |
| `T5.7` | If GO: design one single-file optimization patch; if NO_GO: record why optimization is rejected | `<T5.6>` | `result only` |
| `T5.8` | If GO: apply the one allowed optimization to the chosen collector file | `<T5.7>` | `ArcTool.Core/Services/QuickDimensionDoorWindowCandidateCollector.cs` or `ArcTool.Core/Services/QuickDimensionWallAxisAggregatorService.cs` or `ArcTool.Core/Services/QuickDimensionWallCandidateCollector.cs`, `result only` |
| `T5.9` | Build the optimized candidate or record the explicit no-change outcome | `<T5.8>` | `result only` |
| `T5.10` | If GO: write EV-5 rerun runbook; if NO_GO: synthesize the measured baseline as the Session 4.4 verdict input | `<T5.9>` | `04_EVIDENCE_QUEUE.md`, `result only` |
| `T5.11` | Publish Session 4.4 verdict from EV-4 only or EV-4 + EV-5, and authorize regression work | `<T5.10>` | `result only` |

## Phase 6 — Session 4.5 ArcTool regression and instrumentation disposition

| Task | Purpose | Depends on | Exclusive write scope |
|---|---|---|---|
| `T6.1` | Author the ArcTool-wide regression checklist: startup, ribbon load, command discovery, closed stacks untouched | `<T5.11>` | `result only` |
| `T6.2` | Write the operator runbook for EV-6 ArcTool regression execution | `<T6.1>` | `04_EVIDENCE_QUEUE.md`, `result only` |
| `T6.3` | Judge EV-6 startup/load/regression evidence | `<T6.2>` | `result only` |
| `T6.4` | Decide instrumentation disposition: remove before closure or retain behind an explicit debug gate | `<T6.3>` | `result only` |
| `T6.5` | Apply the instrumentation disposition decision | `<T6.4>` | `ArcTool.Core/Services/QuickDimensionReadOnlyEngine.cs`, `ArcTool.Core/Services/QuickDimensionReadOnlyXmlLogService.cs`, `result only` |
| `T6.6` | Build the closure candidate after instrumentation disposition | `<T6.5>` | `result only` |
| `T6.7` | Publish Session 4.5 final technical verdict and authorize durable closure | `<T6.6>` | `result only` |

## Phase 7 — Durable closure and handoff

| Task | Purpose | Depends on | Exclusive write scope |
|---|---|---|---|
| `T7.1` | Persist Phase 4 outcome and lock the ADR overwrite-prevention rule in repo-local durable channels without rewriting the ADR store | `<T6.7>` | `CLAUDE.md`, `.Dossier/Quick Dimension - Implementation Roadmap.md`, `Memory/MEMORY.md`, `Memory/feedback_adr_store_update_lock.md`, `result only` |
| `T7.2` | Publish final package closure state, refresh the package handoff, and record the optional re-index choice boundary | `<T7.1>` | `06_EXECUTION_STATE.md`, `HANDOFF_TO_NEXT_SESSION.md`, `result only` |

---

## Lock summary

- `04_EVIDENCE_QUEUE.md` is edited only by `T2.3` → `T3.4` → `T4.3` → `T5.4` → `T5.10` → `T6.2`, in that exact order.
- `ArcTool.Core/Services/QuickDimensionReadOnlyEngine.cs` is edited only by `T5.2` then `T6.5`, in that exact order.
- `ArcTool.Core/Services/QuickDimensionReadOnlyXmlLogService.cs` is edited only by `T5.2` then `T6.5`, in that exact order.
- Exactly one collector file may be edited in Phase 5, and only by `T5.8`, after `T5.6` says GO.
- `QuickDimensionDoorWindowCandidateCollector.cs`, `QuickDimensionWallAxisAggregatorService.cs`, and `QuickDimensionWallCandidateCollector.cs` are mutually exclusive candidates for `T5.8`; only one may enter write scope.
- No task edits `ArcTool.Core/Services/QuickDimensionChainCreationService.cs`.
- No task edits `ArcTool.Core/Services/QuickDimensionGridCandidateCollector.cs`.
- No task edits any spike/probe command or model.
- No task edits the Excel or Coordinate stacks.
- Durable persistence (`T7.1`) is master-owned if `CLAUDE.md` must be edited, because the file's own rules require full-context review before in-place updates.

---

## Scope amendment — 2026-08-05

- Session 4.3 is explicitly interpreted as a **Grid safe-failure matrix**, not a Grid feature-expansion phase.
- Forced rollback validation remains outside this package and is tracked only by `.Dossier/Quick Dimension - Deferred Rollback Validation Task.md`.
- ADR protection for this package means strengthening the operating-rule layers and handoff files; it does **not** mean writing a new ADR unless a new architecture decision genuinely emerges.
