# ArcTool — HANDOFF ARCHIVE
**Updated:** 2026-08-07  
**Status:** CLOSED — Quick Dimension R5_C05 evidence capture phase closed; next analysis/fix phase must start in a new chat

---

## Goal and user request

Primary request for this phase:
- analyze Quick Dimension EV-2 remaining case `R5_C05 — Opening flush with end anchor`
- use the attached screenshot plus XML `C:\Users\ADMIN\Desktop\PA4\ArcTool_QD_ReadOnlySummary_387179_Left_20260807_183426.xml`
- determine whether the result is real mid-run/opening acceptance or only an opening-driven false PASS
- record the failure shape so it can be aggregated later for code adjustment

Locked user clarification during this phase:
- the current runtime result is **a fail case**
- the manually created reference image shows the **expected pass shape**
- when an opening anchor coincides with a wall end anchor, the engine should keep only one coincident station, drop the zero-length duplicate, preserve the opposite opening jamb, and still build the remaining chain

---

## Current phase

Phase unit for this chat: **R5_C05 evidence capture only**.

Completed in this phase:
- read the provided XML evidence
- compared the XML outcome against the screenshot and the user-provided expected manual result
- classified the runtime result precisely
- rewrote the handoff so the next chat can continue from the failure summary without re-reading this conversation

No source edit, build, runtime launch, Revit MCP call, or work-package dispatch happened in this phase.

---

## Files modified in this session

Modified:
- `.handoff/HANDOFF_TO_NEXT_SESSION.md`

Created:
- `.handoff/archive/HANDOFF_2026-08-07_qd-r5-c05-fail-evidence-close.md`

Referenced but not modified:
- `Memory/project_qd_chain_creation_audit_handoff.md`
- `C:\Users\ADMIN\Desktop\PA4\ArcTool_QD_ReadOnlySummary_387179_Left_20260807_183426.xml`
- the user-supplied screenshots for actual-vs-expected output

---

## Exact evidence captured

### Fixture
- Selected wall id `387179`
- One opening/window placed near-flush to a resolved wall end anchor
- Side tested: `Left`
- XML file: `C:\Users\ADMIN\Desktop\PA4\ArcTool_QD_ReadOnlySummary_387179_Left_20260807_183426.xml`

### Observed runtime outcome
- Read-only result: `succeeded="true"`
- Final candidates: **2 only**
  - Start Anchor on wall `387179`
  - Finish Anchor on wall `387179`
- Chain creation audit: committed dimension `387437`
- Final displayed dimension: overall `5500 mm`
- No opening candidate survived into `FinalCandidates`

### Decisive diagnostics
- Opening/window `387213` produced a geometry warning and then a hard skip:
  - expected span `1200 mm`
  - instance bbox `1651.55 mm`
  - `raw padded-bbox extrema 0 mm`
  - `selected jamb pair n/a`
  - skipped because **no valid opening-edge reference was available**
- Mid-run aggregation ran but accepted no useful candidate for this case
- Therefore the created dimension is **anchor-only**, not opening-preserving

### Classification
This case is **FAIL**.

It is **not** evidence of mid-run wall acceptance.
It is **not** evidence that opening-vs-anchor collision handling works.
It is an **opening-driven fail masked by a successful anchor-only dimension**.

---

## Expected pass shape locked for later fix work

The user manually recreated the correct intended result.

When an opening anchor coincides with a wall end anchor, the engine should:
1. detect that the opening-side station and wall-end anchor station are coincident
2. keep **one** merged station only
3. remove the duplicate that would otherwise create a `0` segment
4. preserve the **opposite opening jamb**
5. keep the far wall anchor
6. create the chain from those distinct stations

Expected visible chain for this fixture:
- `1200`
- `4300`

Not acceptable:
- dropping the whole opening and falling back to anchor-only `5500`

---

## Technical interpretation to carry forward

The bug is not merely “missing opening reference”.

The failure to preserve the usable opposite jamb means the real defect shape is:
- **anchor/opening station collision not normalized correctly** and/or
- **usable opening station lost when the near jamb collides with the wall anchor** and/or
- **dedupe logic collapses the entire opening contribution instead of only the coincident station**

Working statement for the next phase:
> R5_C05 fails because a flush opening near the wall end is dropped entirely instead of collapsing the coincident wall/opening anchor into one station and preserving the opposite opening jamb. Expected output is a 2-segment chain (`1200` + `4300`), not an anchor-only overall dimension (`5500`).

---

## Done / unfinished / blocked

Done:
- R5_C05 evidence captured
- actual runtime outcome classified
- expected pass behavior recorded from the user’s manual reconstruction
- handoff rewritten for a fresh-chat continuation

Unfinished:
- no source analysis yet for the exact collector/dedupe/fallback code path causing this case
- no code fix yet
- remaining EV-2 cases still need their own evidence/triage state as applicable

Blocked:
- none technically; next chat may continue directly from the failure summary

---

## Verification run

Verification completed:
- XML was read directly
- final candidates verified as wall-anchor-only
- chain creation audit verified as committed with dimension id `387437`
- user clarification locked the expected pass shape and the failure interpretation

Not run:
- no build
- no tests
- no Revit runtime action by Claude
- no Revit MCP action
- no re-index

Reason not run:
- this phase was evidence capture and transfer only

---

## Next-session starting point

Start a **NEW chat**.

Immediate carry-forward objective:
- aggregate this fail with the other remaining EV-2 cases and then isolate the code path that should preserve the far opening jamb under end-anchor coincidence

Minimum context to trust without re-reading this chat:
- `R5_C05` current runtime output is a **false PASS shape**: committed anchor-only dimension `5500`
- expected output is **`1200 + 4300`** after merging the coincident wall/opening anchor and keeping the opposite opening jamb
- the bug category is **collision normalization / opening-station preservation**, not generic chain creation failure

---

## Invariants to preserve

1. Revit runtime is operator-controlled; do not launch Revit or use Revit MCP without explicit user request.
2. Do not treat anchor-only committed dimensions as proof that opening handling succeeded.
3. For flush opening cases, zero-length duplicate stations must be removed without deleting the usable opposite opening jamb.
4. Any later code fix must preserve strict chain-audit semantics; do not relax the audit just to make this case appear green.
5. Revit API lookups remain mandatory before any source fix.

---

## Reference files

- Global handoff: `.handoff/HANDOFF_TO_NEXT_SESSION.md`
- Archived handoff for this phase: `.handoff/archive/HANDOFF_2026-08-07_qd-r5-c05-fail-evidence-close.md`
- Durable QD closure context: `Memory/project_qd_chain_creation_audit_handoff.md`
- XML evidence: `C:\Users\ADMIN\Desktop\PA4\ArcTool_QD_ReadOnlySummary_387179_Left_20260807_183426.xml`
- Root operating document: `CLAUDE.md`
