---
name: arctool-session-learn
description: Classify and persist durable ArcTool knowledge at session close. Use after a meaningful bug fix, roadmap phase close, or architecture decision, or when the user signals a work session/phase/section is ending — decides whether a lesson belongs in CLAUDE.md, .Dossier/, memory/, ADR, or stays session-local, and enforces persisting before the final reply.
---

# ArcTool Session Learn

## When to use

Use this workflow after a meaningful bug fix, roadmap phase close, architecture decision, or when the
user explicitly signals that a work session, phase, or section is ending.

Use it when the session produced knowledge that should survive across machines and future
conversations. Skip it for trivial turns — noisy memory is a cost, not a safety net.

## Goal

Convert transient session discoveries into the correct durable ArcTool knowledge channel without
duplicating content or bloating `CLAUDE.md`.

## ArcTool durable knowledge channels

### `CLAUDE.md`

Use for short, high-leverage technical invariants and operating rules that materially affect future
implementation behavior.

### `.Dossier/`

Use for feature-bounded deep technical records, closure dossiers, root cause analyses, and long-form
implementation context.

### `Memory/`

Use for repository-local cross-session preferences, project constraints, and reference pointers that
are not cleanly derivable from the repository. Add a one-line pointer to `Memory/MEMORY.md`.

### ADR (`manage_adr`)

Use when an architecture rule, feature boundary, reference strategy, or durable trade-off becomes
stable enough to affect future sessions. Read the current ADR state before revising an entry. Never
use ADR for transient debugging notes or routine session progress.

### Tasks or chat

Use for temporary execution state only.

## Classification rubric

Persist the session outcome into exactly one primary durable location.

Choose `CLAUDE.md` when the lesson is short, load-bearing, and should always be present as an
operating rule.

Choose `.Dossier` when the lesson is long-form, bounded to one feature/subsystem, or best preserved
as a narrative technical record.

Choose `Memory/` when the lesson is a durable preference, constraint, or reference pointer that is
not derivable from code.

Keep it only in tasks/chat when it is still temporary, local to the current session, or not yet
reusable.

## Procedure

1. Identify the single highest-value discovery from the session.
2. Decide whether it is durable and non-derivable.
3. Check whether the fact already exists in `CLAUDE.md`, `.Dossier`, or `Memory/`.
4. Update the existing record in place if it already exists.
5. Otherwise, persist it in exactly one primary durable location.
6. Write every durable file **before** the final reply of the turn — no tool can run after it.
7. Re-index (`index_repository`) is the FINAL, OPTIONAL, user-directed step. Offer it as a choice
   after durable files are safe; never gate closure on it.

## Work package closure

When the session was run as a work package (see the `arctool-work-package` skill), closure also
means: `06_EXECUTION_STATE.md` reflects every worker result, and `04_EVIDENCE_QUEUE.md` shows the
real state of each `EV-<n>`. A package left mid-flight must record what is `PENDING` or `BLOCKED`
and who supplies the missing input, so a fresh chat can resume without this conversation.

## Anti-patterns

Do not append long narratives to `CLAUDE.md`.

Do not store temporary progress in `Memory/`.

Do not duplicate the same lesson across all durable layers.

Do not treat machine-local system memory as the ArcTool project source of truth.

## ArcTool examples

### Example 1

A discovered invariant such as "updater registration must use `UIControlledApplication.ActiveAddInId`
captured in `OnStartup()`" belongs in `CLAUDE.md`.

### Example 2

A complete feature closure record or a multi-step root cause analysis belongs in `.Dossier`.

### Example 3

A durable user/project handling preference such as "keep detailed Quick Dimension roadmap outside
`CLAUDE.md`" belongs in repository-local `Memory/`.

### Example 4

A temporary note such as "today we are validating Session 1.1 in Revit" should stay in tasks/chat
unless it becomes a durable lesson.
