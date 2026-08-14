---
name: feedback_master_context_discipline
description: In ArcTool work packages, master context must stay flat: heavy verification artifacts are worker-read by path, and multi-task phases should run as one chained dispatch rather than a long master chat.
metadata:
  type: feedback
---

In ArcTool work packages, the master must preserve context for orchestration and decision-making, not spend it reading heavy verification artifacts. When a user requests a multi-task package phase such as `T2.1..T2.5`, treat it as one chained dispatch/workflow phase: the master carries only compact envelopes forward, and workers read the heavy evidence directly.

Heavy verification artifacts include full XML read-only summaries, Revit journals, screenshots, and large evidence tables. The master should record their paths and route them to the analysis worker. Master-side reads of those artifacts are allowed only after the worker envelope is back and still insufficient for contradiction resolution or closure packaging.

2026-08-07 violation: during EV-1 of the Quick Dimension Phase 4 package, the master pre-read both combined XML evidence files and grep'd the journal before dispatching T2.4, which unnecessarily inflated master context and contributed to compact pressure.

**Why:** The work-package refactor exists specifically to keep master context roughly flat across long phases. Pre-reading heavy evidence at the master duplicates worker reads, wastes context, and undermines the point of the package workflow.

**How to apply:** For future package phases, keep the master to minimum bootstrap files plus compact envelopes. When evidence arrives, store the path in the evidence queue / execution state, dispatch the blocked analysis worker with the path or a tiny routing excerpt, and let the worker read the artifact. In ArcTool work packages, dispatch those workers with `model: "sonnet"` (Claude Sonnet 5) unless the package explicitly records a justified exception.