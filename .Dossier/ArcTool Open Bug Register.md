# ArcTool — Open Bug Register

Last updated: 2026-08-07
Status: Durable open-bug register for non-closed issues that were previously tracked in `CLAUDE.md` and moved out during the rules-only root split.

**Read this file when** root-level bug reminders are needed for operator decisions, or when a subsystem-specific dossier does not already own the open-bug detail.

---

| ID | Area | Issue | Severity |
|---|---|---|---|
| BUG-06 | ArrangeDimension | Missing guard for `activeView.Scale == 0` / unsupported view contexts | Medium |
| BUG-07 | FilterManager | `Idling`-based refresh architecture does not scale on large models | Low |
| BUG-08 | CreateVoidFromLink | `SetParam("Height", -beamHeight)` is still a workaround, not a clean model | Low |
