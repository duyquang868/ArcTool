# Detailed Technical Dossier - ArcTool Knowledge Workflow

## 1. Purpose

This dossier defines how ArcTool captures durable technical knowledge across work sessions without bloating the root `CLAUDE.md`.

The goal is to preserve implementation-critical lessons, user-validated working rules, and roadmap-relevant decisions in portable repository-local files so the same operating context survives across machines.

---

## 2. Sources of truth

### 2.1 `CLAUDE.md`

Use `CLAUDE.md` for short, load-bearing technical invariants and operating rules that materially affect future implementation behavior.

Keep it lean. It is an operating document, not a historical archive.

In particular, keep the enforceable `codebase-memory-mcp` execution rules there, while this dossier carries the longer rationale, classification guidance, and session-close examples.

### 2.2 `.Dossier/`

Use `.Dossier` for bounded deep technical records, feature closure dossiers, roadmap files, root cause analyses, and long-form implementation context.

A dossier is the correct destination when a finding is too long, too specific, or too narrative for `CLAUDE.md`.

### 2.3 `memory/`

Use the repository-local `memory/` directory for durable cross-session preferences, project constraints, and reference pointers that are not cleanly derivable from code or repository structure.

For ArcTool work, repository-local memory is the primary project memory. Machine-local system memory is not the source of truth.

### 2.4 `.codebase-memory/`

Use `.codebase-memory` as graph-derived structural knowledge for architecture tracing, dependency tracing, impact analysis, and symbol discovery.

Do not treat graph output as a replacement for repository-local memory or dossier narratives. It is structural, not editorial.

### 2.5 Task list and chat context

Use tasks and chat for session-local execution state only.

Do not persist temporary execution progress, partial analysis, or transient notes into durable channels unless they become reusable project knowledge.

---

## 3. Classification rules

### 3.1 Put the knowledge in `CLAUDE.md` when

- it changes how future ArcTool code must be written;
- it is a short invariant that future implementation must preserve;
- it is an operating rule that should always load with the repo.

Examples include locked Revit API boundaries, invariant updater rules, file placement rules, and repository operating policy.

### 3.2 Put the knowledge in `.Dossier` when

- it is a feature-bounded technical record;
- it explains a multi-step root cause;
- it captures a closure dossier or implementation narrative;
- it would make `CLAUDE.md` longer or noisier without improving its operating value.

Examples include a feature closure record, a detailed debugging postmortem, or a roadmap that changes over many sessions.

### 3.3 Put the knowledge in `memory/` when

- it is durable across sessions;
- it is not directly derivable from the codebase;
- it reflects user preference, project handling preference, or external reference location.

Examples include language preference, dossier handling preference, and "report unsupported backend scope immediately" style rules.

### 3.4 Keep the knowledge only in tasks or chat when

- it is only relevant to the current session;
- it is a temporary execution note;
- it can be re-derived quickly from current code or git state;
- it has not yet stabilized into a reusable project rule or lesson.

---

## 4. Session-close checklist

When a meaningful work session, phase, or section is ending:

1. Identify the highest-value technical lesson from the session.
2. Decide whether it is durable and non-derivable.
3. Check whether the fact already exists in `CLAUDE.md`, `.Dossier`, or `memory/`.
4. Update the existing record in place if one already exists.
5. Otherwise, persist it in exactly one primary durable location.
6. If repository structure or architecture changed materially, refresh the project knowledge graph before closing.

---

## 5. Anti-patterns

Do not append long historical narratives to `CLAUDE.md`.

Do not store temporary session progress in `memory/`.

Do not duplicate the same lesson across `CLAUDE.md`, dossier files, and memory unless each location carries clearly different value.

Do not let machine-local system memory diverge from the repository-local `memory/` source of truth for ArcTool work.

---

## 6. ArcTool examples

### Example A — `CLAUDE.md`

`App.cs` must capture `UIControlledApplication.ActiveAddInId` during `OnStartup()` because deriving it later from document event sender proved unreliable.

This is short, critical, and changes how future updater code must be written, so it belongs in `CLAUDE.md`.

### Example B — `.Dossier`

The full closure history for Coordinate and the multi-phase Quick Dimension roadmap belong in `.Dossier` because they are long-form, bounded, and evolve over time.

### Example C — `memory/`

"Keep `CLAUDE.md` in English and keep chat responses in Vietnamese" belongs in `memory/` because it is a durable user preference, not a code invariant.

### Example D — tasks/chat only

"Today Session 1.2 is debugging one wall-opening edge case" is session-local and should stay in tasks or chat unless it produces a durable technical lesson later.
