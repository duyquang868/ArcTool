# QD PHASE 4 HARDENING — RESULT SCHEMA

Every worker reply to the master MUST be the compact envelope below.
Do not add prose above or below it.
Do not paste code, XML, screenshots, or long reasoning into the reply.
Detailed findings belong in `results/<TASK_ID>_result.md` only.

```text
<MICRO_RESULT>
task: <id>
status: PASS | BLOCKED | NO_GO
confirmed:
- <max 4 bullets, one line each>
unknown:
- <max 2 bullets>
decision: <one sentence>
files_read: <short list>
files_changed: none | <list>
result_file: .claude/workpackages/quick-dimension-phase4-hardening/results/<TASK_ID>_result.md
next_input: <what the downstream task needs from this task, one or two lines>
blocker: none | <what is missing and who must supply it>
</MICRO_RESULT>
```

## Rules

- `PASS` means the task objective is met and every acceptance condition in the task file is true.
- `BLOCKED` means the task cannot proceed without missing evidence, a missing upstream result, or a
  denied tool call. Name the missing input precisely.
- `NO_GO` means the task proved the package should not proceed on the current path. This is a valid
  outcome and must not be softened into `BLOCKED`.
- `confirmed` must contain only facts the worker verified.
- `unknown` must contain only real unresolved points; omit invented caution.
- `files_changed` lists only files inside the task's declared `write_scope`.
- `next_input` should be actionable by the directly downstream task.
- `blocker` is `none` for `PASS` and `NO_GO`; for `BLOCKED`, it names the human or upstream task
  that must supply the missing input.

## Result-file expectations

Each worker also writes a result file:
`results/<TASK_ID>_result.md`

That file should include:
- task id and date,
- exact inputs read,
- findings grouped by the task objective,
- any citations or evidence excerpts needed for auditability,
- the final decision and handoff note for the next task.

The result file is for durable package memory; the envelope is for orchestration only.
