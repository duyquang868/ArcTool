# ADR store loss — root cause and recovery inventory

Date of investigation and recovery: 2026-08-05.
Subject: `.codebase-memory/adr.md`, the ADR store read and written by `codebase-memory-mcp`.
Scope: why ADR entries disappeared, exactly which ones, what was recoverable and from where, and the
rule that prevents recurrence.

This is a forensic record, not an architecture decision. The prevention rule itself lives in
`CLAUDE.md` `Mandatory editing rules` and in the `## STORE INTEGRITY` section of the store.

---

## 1. Root cause

`manage_adr(project, mode="update", content=...)` **replaces the entire contents** of
`.codebase-memory/adr.md`. It does not append, and it does not merge the submitted text into the
existing store.

Every session that added a new ADR entry submitted only that new entry as `content`. Each such call
therefore wrote a store containing just that entry and silently discarded everything already
present — earlier entries and the `PURPOSE` / `STACK` / `ARCHITECTURE` / `PATTERNS` / `TRADEOFFS` /
`PHILOSOPHY` prose sections alike.

No session malfunctioned, no tool errored, and no call failed. Every call returned success and did
exactly what it was asked to do. The defect is the combination of a destructive tool contract with a
project rule that did not require reading the store first.

### The rule that permitted it

`CLAUDE.md` line 22 read, before this investigation:

> persist it with `manage_adr(project, mode="update", ...)`; **read the current ADR state first when
> revising an existing decision**, …

The read-first requirement was conditioned on *revising an existing decision*. Adding a brand-new
entry read as exempt. Sessions that only ever added new entries followed the rule as written and
still destroyed the store. That conditional clause is the proximate cause and has been amended.

---

## 2. Detection method

The loss was not detected by the tool, by a diff, or by a test. It surfaced from a
`manage_adr(mode="get")` performed in session `1c6bc429` on 2026-08-04 for an unrelated reason: the
returned store held only two entries, while `CLAUDE.md` section 5 cited several older ADR IDs as
locked decisions. That mismatch between the store and `CLAUDE.md` was the only visible symptom.

Reconstruction then used three independent evidence sources:

1. **git** — `git show 3a935f3:.codebase-memory/adr.md` yields the only committed state of the file
   (commit `3a935f3`, 2026-07-12 15:02:56 +0700): 45 lines, 13,002 bytes, 4 entries, all 6 prose
   sections. `git log --all -- .codebase-memory/adr.md` returns that single commit, so git holds
   nothing later.
2. **`git log --all -S`** — searching history for `ADR-2026-07-17B` and `ADR-2026-07-30A` returns
   empty, proving those entries were never committed and that git alone could not recover them.
3. **Claude session transcripts** — `C:\Users\ADMIN\.claude\projects\<project>\*.jsonl` preserve every
   `tool_use` block verbatim, including the full `content` payload of each `manage_adr` call. That
   makes the transcripts a byte-faithful audit log of what each session submitted. Mining them
   recovered the seven never-committed entries.

A recursive scan (`**/*.jsonl`, including `subagents/` and `workflows/`) found 41 `manage_adr` calls
total across 40 session files: 18 `update`, 21 `get`, 2 `sections`. All 18 updates came from
main-session transcripts; no subagent ever wrote the store. A separate scan for direct `Write` /
`Edit` tool calls targeting `adr.md` found **zero** — the file was only ever mutated through
`manage_adr`.

---

## 3. Timeline of `update` calls

All 18 recorded `mode=update` calls, in order. "prose" is how many of the six prose sections the
submitted payload contained.

| Timestamp | Session | Bytes | Entries submitted | prose |
|---|---|---|---|---|
| 2026-07-19T10:57:26 | d3c3f66e | 6460 | 07-17B, 07-17C, 07-18A | 0/6 |
| 2026-07-20T00:41:57 | d3c3f66e | 2073 | 07-19A | 0/6 |
| 2026-07-20T02:14:08 | d1cc7389 | 2334 | 07-19A | 0/6 |
| 2026-07-20T10:13:03 | 07a9f8f2 | 4861 | 07-19A, 07-20A | 0/6 |
| 2026-07-21T07:53:20 | 07a9f8f2 | 5195 | 07-19A, 07-20A | 0/6 |
| 2026-07-21T17:39:40 | 124a31ae | 2341 | 07-22A | 0/6 |
| 2026-07-30T01:33:12 | 05f81f75 | 1849 | 07-30A | 0/6 |
| 2026-07-31T01:13:26 | c0ee9f5b | 2267 | 07-30A | 0/6 |
| 2026-07-31T02:52:37 | 9bd65665 | 3122 | 07-30A | 0/6 |
| 2026-07-31T02:56:28 | 9bd65665 | 3016 | 07-30A | 0/6 |
| 2026-07-31T03:01:41 | 9bd65665 | 3228 | 07-30A | 0/6 |
| 2026-08-01T01:33:02 | 9bd65665 | 3313 | 07-30A | 0/6 |
| 2026-08-01T08:44:20 | 379e6db5 | 3436 | 07-30A | 0/6 |
| 2026-08-02T03:18:00 | b418a38e | 3778 | 07-30A | 0/6 |
| 2026-08-03T12:46:12 | 8c1f09a8 | 3999 | 07-30A | 0/6 |
| 2026-08-04T11:42:03 | 39e654cd | 1760 | 08-04A | 0/6 |
| 2026-08-04T17:20:55 | 1c6bc429 | 2316 | 08-04B | 0/6 |
| 2026-08-04T17:23:00 | 1c6bc429 | 4987 | 06-11, 07-17B, 07-17C, 07-18A, 07-22A, 07-30A, 08-04A, 08-04B | 0/6 |

The seven destructive transitions — each dropping an entry that no later payload carried until the
2026-08-05 rebuild:

| When | Dropped |
|---|---|
| 2026-07-20T00:41:57 | 07-17B, 07-17C, 07-18A |
| 2026-07-21T17:39:40 | 07-19A, 07-20A |
| 2026-07-30T01:33:12 | 07-22A |
| 2026-08-04T11:42:03 | 07-30A |
| 2026-08-04T17:20:55 | 08-04A |

The 2026-08-04T17:23:00 call was a partial recovery: it restored eight IDs from `CLAUDE.md` section 5
and the then-current store, but it did not restore `07-19A` or `07-20A` (never referenced in
`CLAUDE.md`), did not restore the prose sections, and reproduced the older entries as short summaries
rather than their original text. That partial state is what the 2026-08-05 rebuild replaced.

### What the `get` results prove about the earliest loss

21 `mode=get` results are preserved. Every one before the 2026-08-05 rebuild returned **zero prose
sections and zero `ADR-2026-07-12` entries** — including the earliest retained call
(`d3c3f66e`, 2026-07-19T10:47:37), which returned 6016 bytes beginning directly at
`### ADR-2026-07-17B`.

So the four 2026-07-12-era entries and all six prose sections were already gone *before* the first
`update` that transcripts retain. Transcript retention begins 2026-07-19T01:54:52, while the git
commit is dated 2026-07-12 15:02:56 — a seven-day gap with no retained transcripts. The session that
first removed the committed content falls inside that gap and **cannot be identified** from available
evidence. The mechanism is certain; that specific attribution is not.

This corrects two statements written into the store's `## STORE INTEGRITY` section earlier on
2026-08-05, which claimed the loss "happened seven times between 2026-07-19 and 2026-08-04, in seven
different sessions." The seven-event count applies only to the transcript-only entries. The oldest
loss is an eighth, earlier, unattributable event.

---

## 4. Per-entry loss and restore inventory

11 items were lost in total. All 11 are restored.

| Entry | Committed to git? | Lost at | Restored from |
|---|---|---|---|
| `ADR-2026-06-11` Quick Dimension pivots to wall-axis projection | yes (`3a935f3`) | pre-2026-07-19, unattributable | git, verbatim |
| `ADR-2026-07-12` Wall-end anchor = physical solid caps | yes (`3a935f3`) | pre-2026-07-19, unattributable | git, verbatim |
| `ADR-2026-07-12` Chain readiness requires distinct stations | yes (`3a935f3`) | pre-2026-07-19, unattributable | git, verbatim |
| `ADR-2026-07-12` Read-only summary uses millimeters | yes (`3a935f3`) | pre-2026-07-19, unattributable | git, verbatim |
| 6 prose sections (`PURPOSE`…`PHILOSOPHY`) | yes (`3a935f3`) | pre-2026-07-19, unattributable | git, verbatim |
| `ADR-2026-07-17B` Wall Spike directional full-height resolver | **no** | 2026-07-20T00:41:57 | transcript payload 2026-07-19T10:57:26, verbatim |
| `ADR-2026-07-17C` One selected-wall chain at a time | **no** | 2026-07-20T00:41:57 | transcript payload 2026-07-19T10:57:26, verbatim |
| `ADR-2026-07-18A` Mid-run joint via side-line reference evidence | **no** | 2026-07-20T00:41:57 | transcript payload 2026-07-19T10:57:26, verbatim |
| `ADR-2026-07-19A` Accepted mid-run stations exclude endpoint artifacts | **no** | 2026-07-21T17:39:40 | transcript payload 2026-07-21T07:53:20, verbatim |
| `ADR-2026-07-20A` Production aggregator + read-only XML audit log | **no** | 2026-07-21T17:39:40 | transcript payload 2026-07-21T07:53:20, verbatim |
| `ADR-2026-07-22A` NewDimension line and opening semantics gates | **no** | 2026-07-30T01:33:12 | transcript payload 2026-07-21T17:39:40, verbatim |
| `ADR-2026-07-30A` Failure-isolated post-commit chain audit | **no** | 2026-08-04T11:42:03 | transcript payload 2026-08-03T12:46:12, verbatim |
| `ADR-2026-08-04A` Multi-agent work package as default workflow | **no** | 2026-08-04T17:20:55, recovered same day | on-disk store, verbatim |
| `ADR-2026-08-04B` Rollback validation is a separate track | **no** | — | on-disk store, verbatim |

For entries submitted more than once, the **last** payload containing the entry as a complete
`### `-delimited block was used, so each restored entry carries its most-developed wording.

Note the exposure this reveals: only 4 of 13 entries were ever protected by git. The other 9 existed
solely inside a file that a single tool call could overwrite, with session transcripts as the only
backup — and transcripts are machine-local and not guaranteed to be retained indefinitely.

---

## 5. Restoration performed on 2026-08-05

The store was rebuilt from three provenance sources, assembled by script with hard assertions rather
than by hand:

- **git portion** — `git show 3a935f3:.codebase-memory/adr.md`, CRLF normalized to LF, sliced by line
  index into head prose (`## PURPOSE` … `## ARCHITECTURE`), the four entries, and tail prose
  (`## PATTERNS` … `## PHILOSOPHY`). The script asserted each slice starts with the expected heading
  and that the entry slice contains exactly four `### ADR-` headings.
- **transcript portion** — the seven never-committed entries, extracted verbatim from the submitted
  payloads.
- **current portion** — `ADR-2026-08-04A` and `ADR-2026-08-04B`, verbatim from the on-disk file.

Entries are ordered by original submission date, not by recovery source. The obsolete
`### ADR registry note (2026-08-04)` written during the partial recovery was dropped and replaced by
two new sections: `## STORE INTEGRITY` and `## SUPERSESSION INDEX`.

Result: 139 lines, 13 entries, all 6 prose sections. Verified by re-reading through
`manage_adr(mode="get")`, which returned the full rebuilt store — confirming the MCP server reads the
file rather than a stale cache.

A standalone verbatim archive of the seven transcript-recovered entries is kept at
`.Dossier/ADR Store Loss - Verbatim Recovery Archive.md`, so their text no longer depends on
transcript retention. The four git-committed entries are deliberately not copied there; they remain
recoverable with `git show 3a935f3:.codebase-memory/adr.md`.

### Why a supersession index was needed

Restoring older entries alongside newer ones creates a hazard the pre-loss store never had: a reader
could treat a superseded 2026-07-12 rule as current. `## SUPERSESSION INDEX` resolves this, sourced
from existing durable files rather than fresh judgement — notably that the wall-end-anchor entry is
superseded by `ADR-2026-07-17B` for Wall Spike and production porting, and that `ADR-2026-07-18A` is
refined by `ADR-2026-07-19A`. `ADR-2026-07-20A` is recorded as referenced by no other durable file
and kept deliberately.

### Verification caveat

Python reported writing 37,395 bytes (LF newlines) while `wc -c` reported 37,415 on the same file.
The discrepancy is unexplained and was not chased, because content-level verification (entry count,
heading order, section presence, tail integrity, live `manage_adr(mode="get")` read-back) all passed.
Do not cite an exact byte size for this file without re-measuring.

---

## 6. Prevention

1. `CLAUDE.md` line 22 now requires reading the current ADR state before **every**
   `manage_adr(mode="update")` and resubmitting the complete store, not only when revising an
   existing decision, and states the full-replace semantics explicitly.
2. `## STORE INTEGRITY` inside `.codebase-memory/adr.md` states the same rule where a session editing
   the store will see it, without needing to consult `CLAUDE.md` first.
3. `Memory/project_codebase_memory_repo_local_workflow.md` carries the rule as a durable
   cross-session constraint, with a pointer from `Memory/MEMORY.md`.
4. Committing `.codebase-memory/adr.md` to git after meaningful ADR changes would have made all 11
   losses trivially recoverable. This is recommended, not mandated — the file is currently tracked
   and modified in the working tree, and commit timing is the operator's decision.

---

## 7. Related records

- Live store, including `## STORE INTEGRITY` and `## SUPERSESSION INDEX`: `.codebase-memory/adr.md`
- Verbatim archive of transcript-recovered entries:
  `.Dossier/ADR Store Loss - Verbatim Recovery Archive.md`
- Durable rule: `Memory/project_codebase_memory_repo_local_workflow.md`
- Operating rule: `CLAUDE.md` `Mandatory editing rules`
- The `T7.1` result file inside `.claude/quick-dimension-bugfix/results/` recorded a smaller,
  provisional version of this finding on 2026-08-04 (one entry dropped, six IDs absent, backfill
  left as an open question). It is annotated to point here; its original text is preserved as the
  historical record of what was known at the time.
