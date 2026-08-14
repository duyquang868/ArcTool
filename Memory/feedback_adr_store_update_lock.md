---
name: feedback-adr-store-update-lock
description: Standing user instruction (2026-08-05) — ADR writes are exceptional, never part of routine closure persistence, and every write must follow read-resubmit-verify. Do not let the overwrite incident repeat.
metadata:
  type: feedback
---
The user explicitly locked the ADR update rule after the 2026-07/08 overwrite incident: **stop writing ADR by default, and never overwrite the store again.** ADR is no longer one of the normal durable channels for a closing session.

**Why:** the user's words were "đừng có ghi đè adr nửa nhé, khóa luôn quy tắc cập nhật adr luôn cho tôi, đừng để chuyện này tái diển nửa." The mechanical cause is recorded in [[project-codebase-memory-repo-local-workflow]] (`manage_adr(mode="update")` is a full-store replace); this memory records the *policy* the user wants enforced on top of that mechanism. Treating ADR as a routine closure step is what made the destructive call frequent enough to lose 11 items before anyone noticed.

**How to apply:**
1. Default to **not** calling `manage_adr` at all. The standard pre-reply persistence bundle is `Memory/` + `.Dossier` + `CLAUDE.md` + handoff files only.
2. Write ADR **only** when a genuinely new architecture rule, feature boundary, reference strategy, or durable trade-off emerged *and* it cannot live cleanly in `CLAUDE.md`, `.Dossier`, or `Memory/`.
3. Never write ADR for routine bug closure, work-package scaffolding, forensic notes, next-session handoff, transient hypotheses, or session progress.
4. When a write is genuinely warranted, the protocol has no exception: `manage_adr(mode="get")` → resubmit the **complete** store with the new entry appended → read back and verify every previously present entry and prose section survived. An entry-only payload destroys the store silently, with no error.
5. Commit `.codebase-memory/adr.md` to git after any meaningful ADR change — that git copy is the only cheap recovery path.

Forensic record: `.Dossier/ADR Store Loss - Root Cause and Recovery Inventory.md`. Verbatim archive of transcript-recovered entries: `.Dossier/ADR Store Loss - Verbatim Recovery Archive.md`. The enforceable short form of this rule now lives in `CLAUDE.md` `Mandatory editing rules`.
