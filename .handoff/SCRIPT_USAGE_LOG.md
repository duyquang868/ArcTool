# ArcTool — Script Usage Log

This ledger is cumulative.

- Counts are **additive and never reset**.
- Update the existing row for a known script instead of creating a duplicate row.
- Log scripts even when they live outside `.handoff/scripts/`, using their real repository path.
- Read this file before generating a new helper script so existing tools are reused first.

| Script name | Path | Purpose | Total uses | Last used | Notes |
|---|---|---|---:|---|---|
| `run-cbm.cmd` | `.codebase-memory/run-cbm.cmd` | Launch repo-local `codebase-memory-mcp` with the ArcTool-local cache/store. | 1 | 2026-08-06 | Seed entry created during handoff-protocol setup. Update additively on future use. |
