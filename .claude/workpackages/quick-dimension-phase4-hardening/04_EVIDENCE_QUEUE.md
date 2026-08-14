# QD PHASE 4 HARDENING — EVIDENCE QUEUE

The master owns this file.
Workers never edit it.

Update rules:
- Set evidence status to `PENDING`, `SUPPLIED`, or `CANCELLED`.
- Keep `needed for` aligned to the manifest task ids.
- Keep `what the operator must run` concrete and executable.
- Do not ask the operator for evidence until the corresponding runbook task is `PASS`.
- Do not delete history; update in place.

---

## Evidence items

### EV-1 — Session 4.1 clean-model acceptance — PENDING
- runbook: `tasks/T2.3_operator_runbook_clean_fixture.md` (exact steps in `results/T2.3_result.md` §1)
- needed for: `T2.4`, `T2.5`
- asked on: —
- scope: single clean-fixture acceptance, not a matrix. One fixture, one wall, two runs.
- what the operator must run:
  - Build the clean fixture exactly as specified by `T2.1`: one straight non-curtain host wall of
    8000 mm, no joins, no mid-run T-junctions, no visible grids, with hosted `O1` Door 900 mm at
    left jamb 1000 mm, `O2` Window 1200 mm at left jamb 3500 mm, `O3` Door 800 mm at left jamb
    6000 mm.
  - Complete the fixture pre-checks in `results/T2.3_result.md` §1.3 before any run.
  - Run `QuickDimensionCreateChainSmokeCommand` twice on that same wall, in this exact order:
    Run A = `Left/Exterior`, Run B = `Right/Interior`.
  - Return one evidence bundle per run.
- what to return (per run, Run A then Run B):
  - [ ] combined XML path (one combined XML per run)
  - [ ] created dimension id, or the explicit `no dimension created` outcome
  - [ ] annotated screenshot showing all 7 segment values readable, witnesses visible, and labels
        `O1` / `O2` / `O3`
  - [ ] visible dialog / cancel outcome text, or `no dialog`
  - [ ] optional journal excerpt
- evidence quality gate before handing to `T2.4`:
  - [ ] both runs present, in order
  - [ ] all 7 segment values readable in each run's screenshot
  - [ ] fixture pre-checks confirmed

### EV-2 — Session 4.2 wall + Door/Window complexity matrix — SUPPLIED (partial)
- runbook: `tasks/T3.4_operator_runbook_wall_opening.md` (exact steps in `results/T3.4_result.md` §2–4)
- needed for: `T3.5`, `T3.6`, `T3.7`, `T3.8`
- asked on: 2026-08-07 (runbook task `T3.4` closed `PASS`; request is now open to the operator)
- supplied on: 2026-08-07 (partial only: `R6_C06A` baseline XML path supplied and analyzed in `results/T3.7_result.md`)
- scope: 8 mandatory runs covering the 7-case matrix from `T3.1`; `C06` (mirror/flip) is split into
  2 mandatory sub-runs (baseline + one orientation variant) plus up to 2 optional stronger sub-runs.
- what the operator must run:
  - Execute runs `R1`–`R8` in this exact order and with the exact shell per run defined by
    `results/T3.4_result.md`: `R1_C01` empty wall, `R2_C02` single opening, `R3_C03` several
    openings, `R4_C04` close-spaced openings, `R5_C05` opening flush with end anchor, `R6_C06A`
    mirror/flip baseline, `R7_C06B` mirror/flip orientation variant, `R8_C07` mid-run T-junction.
  - Use `QuickDimensionCreateChainSmokeCommand` for every run. One run = one picked wall + one side
    pick + one combined XML.
  - For `R6`/`R7`, keep wall, family/type, width, insertion station, and shell pick fixed and change
    only the orientation flag(s); capture orientation-state evidence for both runs.
  - For `R4`/`R5`, record the actual measured near-collision gap reached in the model, not just the
    setup target.
  - Return one evidence bundle per run, labeled by run id (`R1_C01` .. `R8_C07`).
- what to return (per run):
  - [ ] case-to-run mapping table using labels `R1_C01`..`R8_C07`
  - [ ] combined XML path for every run
  - [ ] created dimension id, or explicit no-dimension/reduced-dimension outcome
  - [ ] case-specific required counts/stations/diagnostics per `results/T3.4_result.md` §2
        (includes: orientation-state evidence for `R6`/`R7`; measured near-collision gap for
        `R4`/`R5`)
  - [ ] annotated screenshot per run
  - [ ] visible dialog / cancel / diagnostic observation, or `no dialog`
  - [ ] optional journal excerpt(s)
- per-run intake log (append one line per supplied run):
  - `R5_C05` — supplied 2026-08-07 — XML `...387179_Left_20260807_183426.xml` — classified FAIL (opening-driven, masked by anchor-only dimension); recorded in `.handoff/archive/HANDOFF_2026-08-07_qd-r5-c05-fail-evidence-close.md`
  - `R6_C06A` — supplied 2026-08-07 — XML `...387179_Left_20260807_185428.xml` — dimension `387510`, 4 refs / 3 segments (1450/900/3150), no dialog — accepted as mirror/flip control sample; screenshot file not present in `PA4` listing (operator pasted it in chat instead); orientation fields absent from XML schema
  - `R7_C06B` (flip sub-case) — supplied 2026-08-07 — XML `C:\Users\ADMIN\Desktop\PA4\ArcTool_QD_ReadOnlySummary_387179_Left_20260807_190513.xml` — controlled comparison vs baseline on same wall / same door / same shell; no change in station ownership, output order, stable-reference tokens, or segment values; dimension `387543`
  - `R7_C06B` (mirror sub-case) — supplied 2026-08-07 — XML `C:\Users\ADMIN\Desktop\PA4\ArcTool_QD_ReadOnlySummary_387562_Left_20260807_190738.xml` — non-regression observed, but fixture is confounded vs baseline (different wall `387562` vs `387179`, different door instance `387563` vs `387482`, different shell `Interior` vs `Exterior`); dimension `387564`; candidate-3 stable-reference token differs (`52`→`51`) but cannot be attributed to mirror alone
  - `R8_C07` — supplied 2026-08-07 — XML `C:\Users\ADMIN\Desktop\PA4\ArcTool_QD_ReadOnlySummary_387562_Left_20260807_202458.xml` — mid-run aggregator executed (`candidateCount=24`) but every scanned candidate ended with `acceptedMidRunStationCount=0`; final chain remained anchor + door only (`0/1450/2350/5500`), dimension `387683`; concrete negative verdict for mid-run T-junction support on this supplied case
- open instrumentation gap found during `R6_C06A` intake: the read-only summary XML records **no** mirrored / hand-flipped / facing-flipped fields for hosted instances, so orientation state is not machine-observable from current evidence and must be supplied by operator statement for `R7_C06B`.

### EV-3 — Session 4.3 Grid safe-failure matrix — SUPPLIED
- runbook: `tasks/T4.3_operator_runbook_grid_safe_failure.md` (exact steps in `results/T4.3_result.md` §1–6)
- needed for: `T4.4`, `T4.5`
- asked on: 2026-08-08 (runbook task `T4.3` closed `PASS`)
- supplied on: 2026-08-08 — 4 XML runs, operator-declared mapping `V1` → `V4` in the order listed below
- scope: 4 mandatory safe-failure runs only. Grid is bystander content in the view/model; it is never the command selection target. Use the same straight non-curtain host wall across all runs where practical so the wall/opening/mid-run chain is comparable against the control.
- what the operator must run:
  - Execute runs `G1_V1` → `G2_V2` → `G3_V3` → `G4_V4` in this exact order using `QuickDimensionCreateChainSmokeCommand` on the same host wall and same side pick where practical.
  - Run mapping: `G1_V1` = `GRID-V1` straight normal-view control; `G2_V2` = `GRID-V2` cropped straight grid; `G3_V3` = `GRID-V3` hidden straight grid; `G4_V4` = `GRID-V4` arc grid.
  - Isolate one variable per run only. Do not combine cropped + arc, hidden + arc, or multiple grid variants in one run.
  - For every run, verify the grid curve type via Properties before running. For `G2_V2`, deliberately record crop-region state and that the grid crosses the crop boundary. For `G3_V3`, record which hide mechanism was used (element `Hide in View` vs. category V/G override). For `G4_V4`, independently confirm the arc grid is valid Revit content before running.
  - Confirm the expected safe-failure shape for every run: no grid-attributable dimension segment, no crash, and the combined XML contains the exact diagnostic string `Grid collection is disabled by Quick Dimension options.`
- what to return (per run, labeled `G1_V1`..`G4_V4`):
  - [ ] variant-to-run mapping confirmation
  - [ ] combined XML path (one per run)
  - [ ] created dimension id, or explicit no-dimension/cancel outcome
  - [ ] exact visible dialog text, or `no dialog`
  - [ ] screenshot showing the attempted context and the outcome
  - [ ] optional journal excerpt
  - [ ] run-specific setup evidence: curve type confirmation for all runs; crop-region state for `G2_V2`; hide mechanism for `G3_V3`; independent arc-grid validity confirmation for `G4_V4`
  - [ ] for `G2_V2` and `G3_V3`, explicit statement whether diagnostic shape or chain differed at all from the `G1_V1` control, even if nothing crashed
- evidence quality gate before handing to `T4.4`:
  - [ ] all four runs present and labeled
  - [ ] one combined XML per run
  - [ ] dimension-or-no-dimension outcome explicit per run
  - [ ] `G3_V3` hide mechanism recorded
  - [ ] `G4_V4` arc-grid validity confirmed

### EV-4 — Session 4.4 performance baseline — PENDING
- runbook: `tasks/T5.4_operator_runbook_performance_baseline.md` (exact steps in `results/T5.4_result.md` §1–4)
- needed for: `T5.5`, `T5.6`, `T5.7`
- asked on: 2026-08-08
- scope: one larger model/view context, one designated host wall + side pick, one warm-up run, then three measured runs on the same context.
- what the operator must run:
  - Load the instrumented build from `T5.3`.
  - Choose one larger project-like context that is meaningfully denser than the clean fixture and report its scale using: wall count, door+window count, and view element count.
  - In that one context, use the same host wall and same side pick for all runs where practical.
  - Run `QuickDimensionCreateChainSmokeCommand` four times total in this exact order:
    - `EV4_WARMUP` — warm-up only; do not use its numbers for judgment.
    - `EV4_M1`
    - `EV4_M2`
    - `EV4_M3`
  - Keep the model, active view, picked wall, shell/side, and visible scene unchanged across `EV4_M1..M3`.
  - Return one combined XML per run. The measured runs must contain `ReadOnlyResult/PerformanceTimings` with these attributes:
    - `totalWallAxisCollectionMs`
    - `wallEndAnchorCollectionMs`
    - `midRunAggregationMs`
    - `openingCollectionMs`
    - `duplicateStationReductionMs`
- what to return:
  - [ ] the scale descriptors for the chosen context: wall count / door-window count / view element count
  - [ ] run-to-file mapping for `EV4_WARMUP`, `EV4_M1`, `EV4_M2`, `EV4_M3`
  - [ ] combined XML path for each run
  - [ ] created dimension id or explicit no-dimension outcome for each run
  - [ ] explicit confirmation that `EV4_M1..M3` used the same wall, same side pick, and same view context
  - [ ] optional screenshot(s)
  - [ ] optional journal excerpt(s)
- evidence quality gate before handing to `T5.5`:
  - [ ] all four runs present and labeled
  - [ ] `EV4_M1..M3` all include the timing block and the five timing attributes above
  - [ ] scale descriptors supplied
  - [ ] repeated measured runs are comparable by construction

### EV-5 — Session 4.4 post-optimization rerun — PENDING
- runbook: `tasks/T5.10_operator_runbook_post_optimization.md`
- needed for: `T5.11`
- asked on: —
- what the operator must run:
  - Only if `T5.6` says optimization GO and `T5.8`/`T5.9` both pass.
  - Re-run the same performance scenario as EV-4 on the optimized candidate.
- what to return:
  - [ ] combined XML / log path(s) with timing output
  - [ ] same model-scale descriptors as EV-4
  - [ ] created dimension id(s) or no-dimension outcome(s)
  - [ ] optional screenshot(s)
  - [ ] optional journal excerpt(s)

### EV-6 — Session 4.5 ArcTool regression — PENDING
- runbook: `tasks/T6.2_operator_runbook_regression.md`
- needed for: `T6.3`, `T6.4`, `T6.7`
- asked on: —
- what the operator must run:
  - Launch ArcTool in the target Revit environment using the closure candidate from `T6.6`.
  - Verify ribbon load, command availability, and that closed Excel/Coordinate stacks remain
    unaffected at startup.
  - Run only the regression checks named in the runbook.
- what to return:
  - [ ] startup / load observation
  - [ ] ribbon / command presence observation
  - [ ] any failure dialog or journal excerpt
  - [ ] optional screenshot(s)
  - [ ] explicit statement whether any non-QD ArcTool behavior regressed
