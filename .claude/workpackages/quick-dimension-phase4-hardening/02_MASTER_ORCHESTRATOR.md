# QD PHASE 4 HARDENING — MASTER ORCHESTRATOR

Read `.claude/workpackages/quick-dimension-phase4-hardening/01_SHARED_CONTRACT.md`,
`.claude/workpackages/quick-dimension-phase4-hardening/03_TASK_MANIFEST.md`,
`.claude/workpackages/quick-dimension-phase4-hardening/05_RESULT_SCHEMA.md`, and
`.claude/workpackages/quick-dimension-phase4-hardening/06_EXECUTION_STATE.md`.
Read `.claude/workpackages/quick-dimension-phase4-hardening/04_EVIDENCE_QUEUE.md` only when a ready task depends on operator evidence or when a worker returns `BLOCKED` for evidence.

Act as the master orchestrator for this work package.
Use one worker per ready task. Respect all dependencies and exclusive write scopes.
Return only compact envelopes from workers. Ask the human for runtime evidence only through the
runbooks listed in `04_EVIDENCE_QUEUE.md`. Update `06_EXECUTION_STATE.md` after every worker result.
Stop on any `NO_GO` gate.

---

## Mission in one paragraph

This package hardens the already-closed Quick Dimension wall-axis chain-creation feature for Phase 4
of the roadmap: prove clean-fixture correctness against a pre-written analytic oracle, classify the
supported and unsupported wall/opening matrix, prove Grid variants fail safely rather than widening
support, measure collector cost on a larger model and optimize only if evidence justifies it, then
confirm no ArcTool-wide regression before durable closure.

## Standing mission boundaries

- Revit runtime is operator-owned.
- `QuickDimensionChainCreationService.cs` stays out of write scope for the whole package.
- `GetReferenceOrderRelation` stays strict: `Exact`, complete `Reversed`, or `Mismatch` only.
- BUG-10 and BUG-11 stay closed unless new runtime evidence proves a fresh defect.
- Grid scope in Session 4.3 means **safe-failure only**, not support expansion.
- Durable persistence must lock the ADR overwrite-prevention rule without touching the ADR store
  unless a genuine new ADR is unavoidable.

## Dispatch protocol

### 1. What each worker gets

A worker receives exactly:
1. `01_SHARED_CONTRACT.md`
2. its own task file
3. the minimum evidence excerpt named by the master, or the raw evidence path when the artifact is heavy

Nothing else unless the task file explicitly names another file.
A worker never reads `CLAUDE.md` in full.
A worker never asks the user for runtime evidence directly.
The master does not pre-read full XML logs, journals, screenshots, or other heavy verification artifacts for the worker.

### 2. Result protocol

Every worker in this package is dispatched with `model: "sonnet"` (Claude Sonnet 5).
Every worker returns only the compact envelope defined in `05_RESULT_SCHEMA.md`.
Detailed reasoning belongs in `results/<TASK_ID>_result.md`.
The master reads that result file only when needed to resolve a contradiction, prepare a downstream
brief, or persist closure.

### 3. Dependency protocol

- Start only tasks whose `Depends on` set is fully `PASS`.
- Do not infer readiness from chat text; use `06_EXECUTION_STATE.md` only.
- If a task returns `BLOCKED`, update `04_EVIDENCE_QUEUE.md` and `06_EXECUTION_STATE.md` before
  doing anything else.
- If a task returns `NO_GO`, stop the wave, record the gate, and wait for a human decision.

### 4. Write-lock protocol

- Never dispatch two write tasks concurrently if their `write_scope` overlaps.
- Respect the serialization order in the manifest lock summary.
- Build tasks must run only after the preceding write task on the same file is `PASS`.

## Evidence protocol

The master is the only actor allowed to ask the human for runtime evidence.
Every ask must reference one `EV-<n>` item in `04_EVIDENCE_QUEUE.md` and must request only:
- the exact operator action from the runbook,
- the expected XML / screenshot / dialog / journal outputs,
- the exact ids or visible outcomes that downstream tasks need.

When evidence arrives:
1. mark the corresponding `EV-<n>` as `SUPPLIED`;
2. record the artifact paths first;
3. forward only the minimum relevant routing excerpt to the blocked worker, or just the artifact path when the file is heavy;
4. update `06_EXECUTION_STATE.md` after the worker returns.

## Package phases

1. **Phase 1 — Preflight and scope locks**
   - resolve the Grid contradiction, freeze vocabulary, lock the baseline, and prove the build path.
2. **Phase 2 — Session 4.1 clean-model acceptance**
   - design the clean fixture, pre-commit the analytic oracle, write the operator runbook, and judge
     the returned evidence.
3. **Phase 3 — Session 4.2 wall + Door/Window complexity matrix**
   - author the case list and predictions, settle the mirror/flip uncertainty, write the runbook,
     and classify every case as Supported, Unsupported-by-design, or Defect.
4. **Phase 4 — Session 4.3 grid safe-failure matrix**
   - prove every Grid variant fails honestly without widening source support.
5. **Phase 5 — Session 4.4 performance and conditional optimization**
   - instrument read-only timing, gather baseline evidence, and either stop at "no hotspot worth the
     risk" or ship one evidence-justified single-file optimization.
6. **Phase 6 — Session 4.5 ArcTool regression and instrumentation disposition**
   - verify the feature does not destabilize startup or the rest of ArcTool, then remove or retain
     instrumentation by explicit decision.
7. **Phase 7 — Durable closure**
   - persist the Phase 4 outcome, hand off the package, and offer re-index only as the final,
     optional, user-directed step.

## Master checklist before every reply

- No heavy verification artifact was read into master context during this dispatch.
- `06_EXECUTION_STATE.md` matches the latest worker/evidence state.
- `04_EVIDENCE_QUEUE.md` reflects real `PENDING`, `SUPPLIED`, or `CANCELLED` evidence.
- No reply claims runtime proof unless the operator actually supplied evidence.
- No reply implies Grid support widened.
- No reply suggests touching the ADR store casually; the current mission plan is to lock the rule in
  operating layers instead.

## Closure condition

The package closes only when `T7.2` is `PASS`, all durable files are written, and the final state is
readable by a fresh chat without this conversation. Re-index is not part of closure.
