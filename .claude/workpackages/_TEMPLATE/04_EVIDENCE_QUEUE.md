# <PACKAGE TITLE> — OPERATOR EVIDENCE QUEUE

The master owns this file. Agents never write here.

Rules:
- Only the master asks the human for evidence. Agents raise `blocker:` instead.
- When evidence arrives, the master forwards **only the relevant routing excerpt** to the specific
  analysis task. Never paste a whole XML into a prompt if a path, table, or short log excerpt suffices.
- Full XML logs, journals, screenshots, and other heavy verification artifacts are worker-read by default.
  The master should record paths and route them, not load them into master context.
- Each request must name: the runbook file, what to run, and exactly what to return.

---

## Request template

```markdown
### EV-<n> — <phase> — <status: PENDING | SUPPLIED | CANCELLED>
- runbook: <task file that defines the exact operator steps>
- needed for: <task ids that are blocked>
- asked on: YYYY-MM-DD
- what the operator must run: <command, model ids, shell(s), build, scenario>
- what to return:
  - [ ] log / XML / screenshot path(s)
  - [ ] created id(s) or visible outcome(s)
  - [ ] commit / rollback / reopen observation
  - [ ] optional journal / console excerpt
- supplied on: —
- forwarded to: —
```

---

## Live queue

No requests yet.
