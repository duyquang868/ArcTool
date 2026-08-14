# ArcTool `.handoff/` — Session Transfer Surface

This folder is the dedicated destination for **session transfer data** ("chuyển giao dữ liệu").
It exists because handoff data changes on every work session, which does not fit the purpose of
`Memory/` (durable cross-session preferences and constraints) or `.Dossier/` (bounded deep
technical records for closed or clearly scoped subsystems).

---

## 1. Layout

```text
.handoff/
├── README.md                      # this file — scope, retention, exception
├── HANDOFF_TO_NEXT_SESSION.md     # canonical handoff for the LATEST normal session
├── SCRIPT_USAGE_LOG.md            # cumulative script ledger; counts never reset
├── archive/                       # per-session historical handoff snapshots
└── scripts/                       # permanently retained session tooling scripts
```

---

## 2. Scope

Use `.handoff/` for **normal (non-work-package) sessions**.

A handoff write is required whenever the user says **"chuyển giao dữ liệu"** (or a near variant such
as "chuyển giao"), and whenever a meaningful session, phase, or section ends with tracked changes.
The handoff must be written as one of the LAST tool actions before the final reply, because no tool
can run after the closing message.

**One chat = one phase.** A chat carries exactly one deliverable phase. When that phase reaches a
hand-offable state, cut the session deliberately: write the handoff, then tell the user the next
phase belongs in a new chat. Do not roll straight into a different phase in the same chat, and do
not let auto-compact become the transition mechanism. Summary-of-summary is a net loss — it costs
about as much as the original record while being less accurate, and each additional layer compounds
the drift. This file is the lossless transfer surface; use it instead.

Phase-boundary signals (any one is enough): the phase deliverable is complete and persisted; the
user shifts to a materially different scope or subsystem; the work is blocked pending user input,
runtime evidence, or an external decision; or context pressure is rising while the current phase is
already hand-offable. Rule record: `Memory/feedback_phase_per_chat_protocol.md`.

Do NOT put in `.handoff/`:
- durable cross-session preferences or project constraints → `Memory/`
- bounded deep technical records, closure dossiers, root cause analyses → `.Dossier/`
- short high-leverage technical invariants and operating rules → `CLAUDE.md`
- architecture decisions → ADR, through the locked read-resubmit-verify protocol

---

## 3. Required handoff contents

Every handoff must cover all of the following. A thin summary is not acceptable.

1. Goal being pursued and the user's original request.
2. Current phase / microtask against the active roadmap.
3. Files created and modified, with exact paths.
4. **Every script generated or used for that session**, with path and purpose.
5. Locked decisions and the reason for each.
6. What is done, what is unfinished, what is blocked.
7. Verification actually run, and what could not run, with the cause.
8. **Cumulative script usage log update** — see section 4.
9. A concrete next step for the following session.

---

## 4. Script retention and cumulative usage

**All scripts are retained permanently. Nothing generated is thrown away.**

- A script written for a session lives on disk under `.handoff/scripts/` unless it legitimately
  belongs next to the subsystem it serves (for example `.codebase-memory/run-cbm.cmd`). Either way
  it must be listed in `SCRIPT_USAGE_LOG.md` with its real path.
- `SCRIPT_USAGE_LOG.md` counts are **additive and never reset**. If a script already shows 5 uses
  and the current session used it twice, the new total is 7.
- **Before writing a new helper script, read `SCRIPT_USAGE_LOG.md` first and reuse an existing
  entry.** The reason this ledger exists is that regenerating near-identical one-off tools every
  session wastes tokens.

---

## 5. Archiving

`HANDOFF_TO_NEXT_SESSION.md` always holds the LATEST normal-session handoff.

Before overwriting it with a new session's content, copy the current version to
`archive/HANDOFF_<YYYY-MM-DD>_<short-slug>.md`. History is never deleted — only moved.

---

## 6. Work-package exception

This protocol does **not** apply to multi-subagent work-package sessions.

Those sessions already persist their state through package-local files and must keep using them:
- `04_EVIDENCE_QUEUE.md` — evidence requests and `PENDING` / `SUPPLIED` status
- `06_EXECUTION_STATE.md` — per-task `PENDING` / `RUNNING` / `PASS` / `BLOCKED` / `NO_GO`
- the package's own `HANDOFF_TO_NEXT_SESSION.md`

The master keeps updating those files as normal. Do not relocate, mirror, or duplicate any
work-package file into `.handoff/`. The live Quick Dimension package stays at
`.claude/quick-dimension-bugfix/` and must not be moved.

If a work-package session also produced reusable scripts, log those scripts in
`SCRIPT_USAGE_LOG.md` anyway — script retention has no exception.

---

## 7. Reference files

- Handoff structure exemplar: `.claude/quick-dimension-bugfix/HANDOFF_TO_NEXT_SESSION.md`
- Work-package contract and state vocabulary: `.claude/workpackages/_TEMPLATE/01_SHARED_CONTRACT.md`,
  `.claude/workpackages/_TEMPLATE/06_EXECUTION_STATE.md`
- Persistence timing rule: `Memory/feedback_persist_memory_before_final_reply.md`
- Work-package activation rule: `Memory/feedback_multi_agent_work_package_workflow.md`
- Protocol memory record: `Memory/feedback_session_handoff_protocol.md`
- Technical operating context: `CLAUDE.md`
