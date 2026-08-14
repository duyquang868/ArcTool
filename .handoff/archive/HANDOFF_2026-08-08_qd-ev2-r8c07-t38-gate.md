# ArcTool — HANDOFF TO NEXT SESSION
**Updated:** 2026-08-07  
**Status:** ACTIVE — Quick Dimension fail-evidence aggregation continues in a new chat; this R5_C05 evidence-only phase is closed

---

## Goal and user request

Primary request for this phase:
- continue analyzing the remaining EV-2 cases after `R1_C01` / `R2_C02` / `R3_C03` / `R4_C04` already passed
- focus on `R5_C05 — Opening flush with end anchor`
- determine whether the observed output proves opening/mid-run acceptance or is only a misleading PASS-like result
- record the failure clearly so it can be aggregated before code adjustment

Locked user clarification during this phase:
- the current runtime output for `R5_C05` is a **fail**
- the user’s manually created screenshot is the **correct expected pass result**
- if an opening anchor coincides with a wall end anchor, the engine should merge the coincident point, avoid a zero-value segment, preserve the opposite opening jamb, and still create the remaining chain

---

## Current phase

Phase unit for this chat: **R5_C05 evidence capture only**.

Completed in this phase:
- read and analyzed XML `C:\Users\ADMIN\Desktop\PA4\ArcTool_QD_ReadOnlySummary_387179_Left_20260807_183426.xml`
- classified the observed runtime result as an opening-driven fail masked by a successful anchor-only dimension
- locked the expected pass behavior from the user’s manual reconstruction
- archived this phase and rewrote the handoff for the next chat

No source edit, build, runtime launch, Revit MCP call, or re-index happened in this phase.

---

## Files modified in this session

Modified:
- `.handoff/HANDOFF_TO_NEXT_SESSION.md`

Created:
- `.handoff/archive/HANDOFF_2026-08-07_qd-r5-c05-fail-evidence-close.md`

Referenced but not modified:
- `Memory/project_qd_chain_creation_audit_handoff.md`
- `C:\Users\ADMIN\Desktop\PA4\ArcTool_QD_ReadOnlySummary_387179_Left_20260807_183426.xml`
- user screenshots showing actual output (`5500`) and manual expected output (`1200 + 4300`)

---

## Exact implementation progress

1. Evidence intake
   - consumed the provided XML through direct read
   - compared it against the screenshot and the user’s description of the expected pass shape

2. Runtime classification
   - verified that `FinalCandidates count="2"`
   - both final candidates are wall anchors on wall `387179`
   - verified committed chain creation audit with dimension id `387437`
   - verified the created dimension is a single overall `5500 mm` result

3. Opening-failure interpretation
   - opening/window `387213` emitted geometry diagnostics but was rejected for `MissingReference`
   - no opening candidate survived into the final chain
   - therefore the output is **anchor-only**, not opening-preserving

4. Expected-pass capture
   - locked the user’s intended resolution rule for anchor/opening coincidence
   - expected result shape for this fixture is a 2-segment chain: `1200` + `4300`

---

## Evidence found during verification

### Fixture facts
- Selected wall id: `387179`
- Side: `Left`
- Axis length: `5500 mm`
- One opening/window near-flush to a resolved wall end anchor

### Observed XML outcome
- `FinalCandidates count="2"`
- Candidate 1 = wall start anchor
- Candidate 2 = wall finish anchor
- `ChainCreationAudit attempted="true" succeeded="true" ... dimensionId="387437"`
- single segment value = `5500 mm`

### Opening diagnostics
- opening geometry warning reports:
  - expected span `1200 mm`
  - instance bbox `1651.55 mm`
  - `raw padded-bbox extrema 0 mm`
  - `selected jamb pair n/a`
- hard skip reason:
  - **no valid opening-edge reference was available**

### Evidence conclusion
This case is **FAIL**.

It does **not** prove mid-run wall acceptance.
It does **not** prove collision handling between opening station and end anchor.
It proves the engine can fall back to a superficially successful **anchor-only** output while losing the opening contribution entirely.

---

## Locked expected behavior for later fix work

When an opening anchor and a wall end anchor coincide, the engine should:
1. collapse the coincident wall/opening anchor into **one** station
2. remove the duplicate that would create a zero-length segment
3. preserve the opposite opening jamb
4. preserve the far wall anchor
5. create the chain from the remaining distinct stations

Expected visible output for this fixture:
- `1200`
- `4300`

Not acceptable:
- dropping the opening completely and producing only `5500`

Working defect statement:
> R5_C05 fails because a flush opening near the wall end is dropped entirely instead of collapsing the coincident wall/opening anchor into one station and preserving the opposite opening jamb. Expected output is a 2-segment chain (`1200` + `4300`), not an anchor-only overall dimension (`5500`).

---

## Locked decisions and reasons

1. **Treat R5_C05 as a true fail, not a partial pass.**
   - Reason: the opening contribution is lost; anchor-only success is insufficient.

2. **Use the manual `1200 + 4300` screenshot as the expected pass shape.**
   - Reason: the user explicitly confirmed it reflects the correct intended behavior.

3. **Classify the bug as collision normalization / opposite-jamb preservation.**
   - Reason: the useful opening geometry should survive even when one opening station coincides with the wall anchor.

4. **Do not treat this case as evidence of mid-run acceptance.**
   - Reason: no opening or mid-run candidate survives into `FinalCandidates`.

5. **Close this phase at evidence capture only.**
   - Reason: the user wants a clean transfer and will continue in a new chat.

---

## Done / unfinished / blocked

Done:
- `R5_C05` evidence captured
- actual runtime result classified
- expected pass shape recorded
- handoff archived and rewritten

Unfinished:
- remaining EV-2 cases still need full evidence aggregation as applicable
- no source-level diagnosis yet for the exact collector/dedupe/fallback path
- no code fix yet

Blocked:
- none technically; continuation should start in a fresh chat

---

## Verification run

Verification completed:
- XML read directly
- final candidate set confirmed as anchor-only
- committed dimension id confirmed as `387437`
- failure interpretation confirmed against the user’s manual expected result

Not run:
- no build
- no tests
- no Revit runtime action
- no Revit MCP action
- no re-index

Reason not run:
- this phase was evidence capture and handoff only

---

## Next-session starting point

Start a **NEW chat**.

Immediate carry-forward context:
- `R1_C01` / `R2_C02` / `R3_C03` / `R4_C04` are already considered PASS by the user
- `R5_C05` is now locked as a **fail** with this exact shape:
  - actual runtime output: anchor-only `5500`, dimension id `387437`
  - expected output: `1200 + 4300`
  - root defect category: coincident wall/opening station collapse must preserve the opposite opening jamb

Minimum restatement to trust without re-reading this conversation:
- do **not** count the current R5_C05 output as opening success
- the near opening jamb and wall end anchor should collapse to one station
- the far opening jamb must survive and remain in the chain

---

## Invariants to preserve

1. One chat = one phase; this evidence-only phase is closed and the next one starts in a new chat.
2. Revit runtime is operator-controlled: no Revit launch, `.rvt` open, MCP call, or smoke test without explicit request.
3. Anchor-only committed dimensions are not proof that opening logic worked.
4. For flush opening cases, dedupe must remove only the coincident station, not the whole opening contribution.
5. Any later fix must preserve strict chain-audit semantics; do not relax evidence rules to mask the defect.
6. Revit API docs must be checked before any code-change phase.

---

## Reference files

- Archived handoff for this closed phase: `.handoff/archive/HANDOFF_2026-08-07_qd-r5-c05-fail-evidence-close.md`
- Durable QD closure context: `Memory/project_qd_chain_creation_audit_handoff.md`
- XML evidence: `C:\Users\ADMIN\Desktop\PA4\ArcTool_QD_ReadOnlySummary_387179_Left_20260807_183426.xml`
- Root operating document: `CLAUDE.md`
