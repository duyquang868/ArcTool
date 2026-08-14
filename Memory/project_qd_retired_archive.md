---
name: project_qd_retired_archive
description: Quick Dimension retired/archive status; retired 2026-08-10 after EV-4, source preserved under ArcTool.Core/Archive/QuickDimension and excluded from compilation; do not treat QD as active roadmap work.
metadata:
  type: project
---

# Quick Dimension retired/archive status — updated 2026-08-10

## Retirement state

- **Quick Dimension is RETIRED.** On 2026-08-10, after operator EV-4, the feature was judged no longer feasible/appropriate to continue developing.
- Retirement means Quick Dimension is no longer active roadmap work in ArcTool and must not be scheduled or framed as the current feature priority.
- This is not a deletion of history: the feature's source, roadmap, and prior runtime evidence remain preserved for reference.

## Archive + compilation state

- Archived source root: `ArcTool.Core/Archive/QuickDimension/`
- Archived source folders:
  - `ArcTool.Core/Archive/QuickDimension/Commands/`
  - `ArcTool.Core/Archive/QuickDimension/Models/`
  - `ArcTool.Core/Archive/QuickDimension/Services/`
- Archived files are excluded from compilation by:
  - `ArcTool.Core/ArcTool.Core.csproj` → `<Compile Remove="Archive\QuickDimension\**\*.cs" />`
- Live ribbon entry points were removed from `ArcTool.Core/App.cs`, so the active product surface no longer exposes Quick Dimension.

## How to use this memory

- Treat Quick Dimension as a closed/retired subsystem unless a future operator explicitly decides to revive it.
- For prior technical/runtime history, read `Memory/project_qd_chain_creation_audit_handoff.md` and `.Dossier/Quick Dimension - Implementation Roadmap.md`.
- For the bounded retirement closure record, read `.Dossier/Quick Dimension - Retirement Record.md`.
