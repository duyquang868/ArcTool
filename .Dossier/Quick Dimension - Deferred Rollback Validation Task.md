# Quick Dimension — Deferred Rollback Validation Task

Status: DEFERRED (not a blocker)
Deferred on: 2026-08-04 (master decision, operator-approved)
Owning subsystem: Quick Dimension chain creation
Primary source of record: `ArcTool.Core/Services/QuickDimensionChainCreationService.cs`

This dossier is self-contained. A future session can pick it up cold, without any
work-package file, and execute it as an independent task. Its purpose is to stop
forced-rollback validation from being re-litigated inside unrelated Quick Dimension
missions.

---

## 1. The decision

Forced-rollback validation of `QuickDimensionChainCreationService.CreateChainDimension`
is **out of scope** for the BUG-11 / BUG-10 mission (named-reference identity↔station
atomicity, and `HostWallOpeningGeometry` fallback candidate owner metadata). It is
deferred to this standalone future task.

Why deferral is technically sound, not a gap:

1. **No fix touched the rollback surface.** The BUG-11 / BUG-10 mission edited only
   `QuickDimensionDoorWindowCandidateCollector.cs` (collector-side identity/station and
   candidate metadata) and `QuickDimensionReadOnlyXmlLogService.cs` (audit/logging).
   `QuickDimensionChainCreationService.cs` was never in any task's write scope, so its
   transaction and rollback behavior is byte-identical to the pre-fix build. Rollback is
   therefore not a regression surface for either defect.
2. **A rollback run cannot confirm or refute either fix.** Both fixes are observable only
   in *committed* output: the audit's `referenceOrderRelation` / identity / owner /
   segment-value gates, and candidate ownership metadata. A rolled-back run produces no
   committed dimension and no meaningful audit comparison.
3. **Positive evidence is already stronger than a forced-rollback probe.** The regression
   matrix committed cleanly on every required run with `Exact` reference order and all
   audit gates true, and every created dimension survived save / close / reopen.
   Reopen-persistence validation is settled and is *not* part of this deferral.

Consequence: mission closure gates were narrowed to the commit-path half. Absence of
rollback evidence must not be treated as an open mission defect.

---

## 2. Control-flow finding to preserve (do not re-derive)

Established by the T6.5 rollback/reopen analysis (2026-08-03). Reuse these anchors as-is;
re-deriving them from source is wasted effort and risks drift.

**Every operator-reachable invalid input is rejected with `Result.Cancelled` *before* the
creation transaction starts**, in `ArcTool.Core/Commands/QuickDimensionCreateChainSmokeCommand.cs`:

| Rejected condition | Anchor |
|---|---|
| Unsupported active view | `QuickDimensionCreateChainSmokeCommand.cs:37-45` |
| Invalid wall / curtain wall / non-line wall | `QuickDimensionCreateChainSmokeCommand.cs:49-76`, `:249-277` |
| Side pick landing on the wall axis | `QuickDimensionCreateChainSmokeCommand.cs:81-99` |
| Read-only result with fewer than two distinct-station candidates | `QuickDimensionCreateChainSmokeCommand.cs:107-116`; contract at `QuickDimensionContract.cs:576-577` |

**The creation transaction starts only at** `QuickDimensionChainCreationService.cs:88-89`.

**The rollback branches execute only on internal post-start failures**
(`QuickDimensionChainCreationService.cs:93-149`, rollback helper at `:153-159`):

- `doc.Create.NewDimension(...)` returns null;
- created reference count differs from the expected candidate count;
- commit status is not `Committed`;
- an exception is thrown inside the transaction.

Direct consequence: a curved wall, curtain wall, unsupported view, or an axis-landing side
pick proves **cancellation**, never rollback. Those inputs never reach
`Transaction.Start()`, so they cannot be used as rollback fixtures.

API vocabulary consulted for the same analysis:
`Transaction.RollBack()` — https://www.revitapidocs.com/2026/bd1e69e9-961e-1c07-f70a-a29b90c6eb97.htm ;
`Document.Close(bool)` — https://www.revitapidocs.com/2026/5948b03d-5537-33d4-6e38-a8f16d5d6779.htm

---

## 3. Operator constraint (decided, not an open question)

The operator explicitly stated they **cannot manufacture a creation failure by normal
modelling actions**, and approved deferring forced-rollback validation to a separate
future task on 2026-08-04.

Therefore:

- **Do not issue a runtime request for forced-rollback validation as-is.** There is no
  verified operator-directable fixture that reaches `Transaction.Start()` and then fails.
- Naming any specific wall or shell as "the rollback fixture" would be guesswork and is
  forbidden. Revit runtime remains operator-owned: no session may launch Revit, open an
  `.rvt`, call Revit MCP, or run a smoke test on its own initiative.
- Any future rollback validation requires a *new deliberate mechanism* (section 5), not a
  better-worded runbook over the existing build.

---

## 4. Trigger conditions that reopen this task

Reopen only when one of these is actually observed — not speculatively:

1. **Field-observed rollback defect** — a real Quick Dimension creation failure reported
   from normal use, where the model is left in a wrong state afterwards.
2. **Dirty-transaction report** — Revit warns about an unclosed/unfinished transaction, an
   undo entry appears for a failed creation, or a subsequent command fails because a
   transaction was left open by Quick Dimension.
3. **Corrupted or truncated read-only XML after a failed creation** — the combined XML from
   a failed run is malformed, partially overwritten, or lost, indicating the
   temp-file + `File.Replace` audit append is not failure-isolated on the rollback path.
4. **Orphan `Dimension`** — a dimension element exists in the model after the command
   reported failure.
5. **A future change puts `CreateChainDimension` in scope.** If any mission edits the
   transaction, reference-array assembly, commit-status handling, or rollback helper in
   `QuickDimensionChainCreationService.cs`, rollback validation is no longer deferrable and
   must be reinstated as a closure gate for that mission.

---

## 5. Entry approach for the future task (sketch only)

Goal: force a post-`Transaction.Start()` failure **without changing the production fix
surface**. This is an approach sketch, deliberately not an implementation.

- **Fault-injection harness, debug-only.** Introduce a debug/diagnostic-only switch that
  makes one creation attempt fail *after* the transaction has started — for example by
  substituting a deliberately invalid or stale `ReferenceArray`, or by raising a controlled
  exception at the injection point — so each of the four rollback branches can be exercised
  in turn.
- **Isolation requirements.** The switch must be off by default, must be unreachable from
  the normal ribbon/operator path, and must not alter candidate collection, reference
  selection, dimension-line derivation, audit semantics, or any behavior validated by the
  BUG-11 / BUG-10 mission. Preferred shape is a separate debug command or harness entry
  point that calls the existing service, rather than new branches threaded through
  `CreateChainDimension` itself.
- **Coverage plan.** One run per rollback branch: null `NewDimension` result, reference-count
  mismatch, non-`Committed` commit status, and thrown exception.
- **Baseline pairing.** Each injected-failure run should be paired with one clean committed
  run on the same fixture, so "model unchanged after rollback" is measured against a known
  good state rather than an assumption.
- **Removal or gating.** Decide up front whether the harness ships gated or is removed after
  validation; it must never be reachable in a release build's operator flow.

---

## 6. Evidence required to close the future task

The future task closes only with all of the following, per exercised rollback branch:

1. **Clean rollback** — the command reports failure, and the model state is identical to the
   pre-run baseline (no geometry, no parameter, no view changes).
2. **No dirty transaction** — no Revit warning about an unfinished transaction, no undo
   entry for the failed attempt, and subsequent commands run normally in the same session.
3. **Read-only XML preserved** — the pre-existing read-only XML for that run is intact and
   well-formed; a failed creation must not truncate, corrupt, or delete it. Audit-append
   failure must remain failure-isolated.
4. **No orphan `Dimension`** — no dimension element from the failed attempt exists in the
   model, verified by element id and by reopening the document.
5. **Branch attribution** — each observation is tied to the specific rollback branch that
   was injected, so partial coverage is visible rather than implied.
6. **Fix surface untouched** — a diff confirming the production collector, audit, and
   creation logic validated by the BUG-11 / BUG-10 mission are unchanged by the harness
   work; plus one clean committed regression run to prove the commit path still passes all
   audit gates.
