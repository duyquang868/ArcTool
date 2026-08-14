# QD BUGFIX — RESULT SCHEMA (v1)

Two outputs per task. Do not mix them.

---

## A. Detailed result file (written by the agent)

Path: `.claude/quick-dimension-bugfix/results/<TASK_ID>_result.md`

```markdown
# <TASK_ID> — <short title>

- status: PASS | BLOCKED | NO_GO
- date: YYYY-MM-DD
- inputs_read: <files / symbols / evidence excerpts actually used>
- write_scope_touched: none | <files>

## Findings
<full detail — quotes, line refs, tables, API citations with URLs>

## Decision
<one paragraph: what is now settled>

## Open questions
<max 3; empty if none>

## Handoff for downstream tasks
<the minimum facts the next task needs, stated plainly>
```

Length target: under 250 lines. No full-file source dumps; quote at most ~30 lines total
and cite `file:line` instead.

---

## B. Compact envelope (the ONLY thing returned to the master)

```text
<QD_MICRO_RESULT>
task: <id>
status: PASS | BLOCKED | NO_GO
confirmed:
- <max 4 bullets, one line each>
unknown:
- <max 2 bullets>
decision: <one sentence>
files_read: <short list>
files_changed: none | <list>
result_file: .claude/quick-dimension-bugfix/results/<TASK_ID>_result.md
next_input: <what the downstream task needs from this task, one or two lines>
blocker: none | <what is missing and who must supply it>
</QD_MICRO_RESULT>
```

Hard limit: 25 lines. No source code, no XML, no screenshots, no transcript in the envelope.

---

## Status meanings

- **PASS** — objective met, downstream task may start.
- **BLOCKED** — missing operator evidence, missing upstream decision, or a denied tool call.
  Name exactly what is needed. Never guess to force a PASS.
- **NO_GO** — the task's own gate concluded "do not proceed" (used by T1.7 and T3.4).
  This is a valid, useful outcome, not a failure.
