# QD BUGFIX — OPERATOR EVIDENCE QUEUE

The master owns this file. Agents never write here.

Rules:
- Only the master asks the human for evidence. Agents raise `blocker:` instead.
- When evidence arrives, the master forwards **only the relevant excerpt** to the specific
  analysis task. Never paste a whole XML into a prompt if a station/reference table suffices.
- Each request must name: the runbook file, what to run, and exactly what to return.

---

## Request template

```markdown
### EV-<n> — <phase> — <status: PENDING | SUPPLIED | CANCELLED>
- runbook: <task file that defines the exact operator steps>
- needed for: <task ids that are blocked>
- asked on: YYYY-MM-DD
- what the operator must run: <command, wall id, shell(s)>
- what to return:
  - [ ] combined XML path(s) (one per run — audit is appended to the read-only XML)
  - [ ] created dimension id(s)
  - [ ] annotated screenshot(s) of displayed segment values
  - [ ] transaction/commit or rollback observation
  - [ ] optional journal excerpt
- supplied on: —
- forwarded to: —
```

---

## Live queue

### EV-1 — Phase 2 instrumentation smoke — SUPPLIED
- runbook: `tasks/T2.5_operator_runbook_379469.md` + `results/T2.5_result.md`
- needed for: T3.1, T3.2, T3.3, T3.4
- asked on: 2026-08-03
- what the operator must run: `QuickDimensionCreateChainSmokeCommand` on wall **379469**,
  both shells (Left/Exterior first, then Right/Interior), using the instrumented build from T2.4.
- what to return:
  - [x] 2 combined XML files (one per shell run, audit appended in place)
  - [x] 2 created dimension ids (one per shell)
  - [x] 2 annotated screenshots showing displayed segment values (one per shell)
  - [x] commit/rollback observation per run
  - [x] optional journal excerpt
- supplied on: 2026-08-03
- forwarded to: T3.1, T3.2, T3.3, T3.4

### EV-2 — Phase 6 regression matrix — SUPPLIED
- runbook: `tasks/T6.1_operator_runbook_regression_matrix.md`
- needed for: T6.2, T6.3, T6.4, T6.6
- asked on: 2026-08-03
- what the operator must run: `QuickDimensionCreateChainSmokeCommand` on walls **379467**, **379469**, and **379470**, both shells each (Left/Exterior and Right/Interior), for **6 runs total** on the fixed build from T5.5.
- what to return:
  - [x] 6 combined XML paths (one per run — audit appended in place)
  - [x] 6 created dimension ids (one per run)
  - [x] 6 annotated screenshots showing displayed segment values (one per run)
  - [x] commit/rollback observation per run
  - [x] optional journal excerpt
- supplied on: 2026-08-04
- forwarded to: T6.2, T6.3, T6.4

### EV-3 — Reopen validation — SUPPLIED
- runbook: `tasks/T6.5_rollback_reopen_runbook.md`
- needed for: T6.6, T7.1
- what the operator must run: save/close/reopen validation for the six committed EV-2
  dimensions. Forced-rollback validation was removed from this mission's closure scope on
  2026-08-04 and deferred to a separate future task.
- what to return:
  - [x] reopen persistence observation for committed dimensions `385355`, `385356`,
    `385632`, `385584`, `385719`, `385720`
  - [x] post-reopen screenshot
  - [x] operator confirmation that displayed values stayed unchanged and dimensions still
    appear immediately at the picked side/position
  - [ ] optional journal excerpt
- supplied on: 2026-08-04
- forwarded to: T6.6, T7.1
