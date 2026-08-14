# ArcTool Prompt Cache Tiers

Last updated: 2026-08-06
Status: Durable prefix convention for Anthropic API integrations and context-assembly scripts that package ArcTool repository context themselves. This is not a Claude Code toggle.

## Purpose

This dossier defines a fixed prompt-assembly order for ArcTool so repeated API requests can reuse prompt-cache prefixes safely.

The design follows four tiers:
- T1 — system invariants
- T2 — stable workflow scaffold
- T3 — durable project context
- T4 — dynamic working set

Always assemble tiers in this order: `T1 -> T2 -> T3 -> T4`.

## Hard rules

1. Preserve the exact tier order `T1 -> T2 -> T3 -> T4`.
2. Keep file order stable within each tier.
3. Put cache breakpoints only after T1, T2, and T3.
4. Never insert dynamic content between cached tiers.
5. Keep T4 last and uncached.
6. Never place handoff state, execution state, evidence queues, logs, journals, XML excerpts, or source excerpts in a cached tier.
7. Scope T3 by active subsystem; do not load unrelated subsystem dossiers into the same cached prefix.
8. `Memory/MEMORY.md` is not a cache-tier input because it changes whenever durable memory changes.
9. When a tier changes, every later tier cache is invalidated for the next request.

## Cache-breakpoint layout

Use at most three cacheable breakpoints:
- after T1
- after T2
- after T3

T4 has no breakpoint.

## Tier map

| Tier | Stability | Cache breakpoint | ArcTool files |
|---|---|---|---|
| T1 | Very high | Yes | `CLAUDE.md`; `Memory/feedback_adr_store_update_lock.md`; `Memory/feedback_claude_md_and_chat_language.md`; `Memory/feedback_multi_agent_work_package_workflow.md`; `Memory/feedback_persist_memory_before_final_reply.md`; `Memory/feedback_revit_runtime_operator_control_and_journal_analysis.md`; `Memory/feedback_session_handoff_protocol.md`; `Memory/project_codebase_memory_repo_local_workflow.md` |
| T2 | High | Yes | `.claude/workpackages/README.md`; `.claude/skills/arctool-work-package/SKILL.md`; `.claude/workpackages/_TEMPLATE/02_MASTER_ORCHESTRATOR.md`; `.claude/workpackages/_TEMPLATE/05_RESULT_SCHEMA.md`; `.claude/workpackages/_TEMPLATE/03_TASK_MANIFEST.md` |
| T3 | Medium | Yes | `.Dossier/Detailed Technical Dossier - ArcTool Knowledge Workflow.md`; `.Dossier/Detailed Technical Dossier - Multi-Agent Work Package Workflow.md`; plus exactly one subsystem cluster from section 4 |
| T4 | Low | No | `.handoff/HANDOFF_TO_NEXT_SESSION.md`; `.handoff/SCRIPT_USAGE_LOG.md`; `.claude/workpackages/<slug>/04_EVIDENCE_QUEUE.md`; `.claude/workpackages/<slug>/06_EXECUTION_STATE.md`; package `HANDOFF_TO_NEXT_SESSION.md`; current task file; Revit XML/journal/log excerpts; source excerpts; current user request |

## T1 — System invariants

T1 is the most stable cached prefix. Place `CLAUDE.md` first, then append policy memories in the exact order listed in the tier map.

Include only durable rule-shaped files here:
- repository operating rules
- persistence timing rules
- runtime boundary rules
- handoff routing rules
- ADR safety rules
- language/presentation rules that materially affect future sessions

Do not place subsystem history here.

### Excluded from T1

Keep these out of T1 even if they are durable:
- every `Memory/project_qd_*` file
- `Memory/gemma4_worker_constitution.md`
- `Memory/gemma4_task_delegation_template.md`
- `Memory/gemma4_error_learning_log.xml`
- `Memory/MEMORY.md`

These files are either subsystem-specific, integration-specific, or too volatile for the invariant prefix.

## T2 — Stable workflow scaffold

T2 contains mission-agnostic workflow files that describe how ArcTool multi-file work is packaged and coordinated.

Default T2 order:
1. `.claude/workpackages/README.md`
2. `.claude/skills/arctool-work-package/SKILL.md`
3. `.claude/workpackages/_TEMPLATE/02_MASTER_ORCHESTRATOR.md`
4. `.claude/workpackages/_TEMPLATE/05_RESULT_SCHEMA.md`
5. `.claude/workpackages/_TEMPLATE/03_TASK_MANIFEST.md`

### Optional T2 members

Only for integrations that always run in package mode, optionally append these in a fixed slot and keep that choice stable across all requests of that integration:
- `.claude/workpackages/_TEMPLATE/01_SHARED_CONTRACT.md`
- `.claude/skills/arctool-session-learn/SKILL.md`

### Excluded from T2

Never cache these as part of T2:
- `.claude/workpackages/_TEMPLATE/04_EVIDENCE_QUEUE.md`
- `.claude/workpackages/_TEMPLATE/06_EXECUTION_STATE.md`

Even as templates, they mirror files that must be loaded live in T4.

## T3 — Durable project context

T3 starts with two always-on workflow dossiers, then appends exactly one subsystem cluster.

Base T3 order:
1. `.Dossier/Detailed Technical Dossier - ArcTool Knowledge Workflow.md`
2. `.Dossier/Detailed Technical Dossier - Multi-Agent Work Package Workflow.md`
3. one subsystem cluster from the list below

### Subsystem clusters

#### Quick Dimension
1. `.Dossier/Quick Dimension - Implementation Roadmap.md`
2. `Memory/project_qd_projection_pivot.md`
3. `Memory/project_qd_chain_creation_audit_handoff.md`

#### Coordinate
1. `.Dossier/Detailed Technical Dossier - Coordinate Feature.md`

#### Excel to Revit
1. `.Dossier/Detailed Technical Dossier - Excel to Revit.md`

#### Filter Manager
No extra dossier cluster currently exists. Use only the base T3 files and rely on `CLAUDE.md` sections 3 and 6.

#### ADR / persistence recovery
1. `.Dossier/ADR Store Loss - Root Cause and Recovery Inventory.md`

### T3 selection rule

Pick the subsystem cluster by request family and keep that cluster membership stable across repeated requests of that family. Rotating T3 membership between otherwise-identical requests destroys cache reuse.

## T4 — Dynamic working set

T4 is always last and never cached.

Typical T4 members:
- `.handoff/HANDOFF_TO_NEXT_SESSION.md`
- `.handoff/SCRIPT_USAGE_LOG.md`
- `.claude/workpackages/<slug>/04_EVIDENCE_QUEUE.md`
- `.claude/workpackages/<slug>/06_EXECUTION_STATE.md`
- `.claude/workpackages/<slug>/HANDOFF_TO_NEXT_SESSION.md`
- current package task file or microtask file
- Revit XML excerpts
- Revit journal excerpts
- build or runtime log excerpts
- source excerpts under discussion or edit
- current user request

`.handoff/HANDOFF_TO_NEXT_SESSION.md` is intentionally uncached because it is rewritten across normal sessions and may contain stale session status if reused as a prefix.

## Starter preset

```text
T1  CLAUDE.md
    Memory/feedback_adr_store_update_lock.md
    Memory/feedback_claude_md_and_chat_language.md
    Memory/feedback_multi_agent_work_package_workflow.md
    Memory/feedback_persist_memory_before_final_reply.md
    Memory/feedback_revit_runtime_operator_control_and_journal_analysis.md
    Memory/feedback_session_handoff_protocol.md
    Memory/project_codebase_memory_repo_local_workflow.md
    -> cache breakpoint

T2  .claude/workpackages/README.md
    .claude/skills/arctool-work-package/SKILL.md
    .claude/workpackages/_TEMPLATE/02_MASTER_ORCHESTRATOR.md
    .claude/workpackages/_TEMPLATE/05_RESULT_SCHEMA.md
    .claude/workpackages/_TEMPLATE/03_TASK_MANIFEST.md
    -> cache breakpoint

T3  .Dossier/Detailed Technical Dossier - ArcTool Knowledge Workflow.md
    .Dossier/Detailed Technical Dossier - Multi-Agent Work Package Workflow.md
    <one subsystem cluster>
    -> cache breakpoint

T4  .handoff/HANDOFF_TO_NEXT_SESSION.md
    .handoff/SCRIPT_USAGE_LOG.md
    .claude/workpackages/<slug>/04_EVIDENCE_QUEUE.md
    .claude/workpackages/<slug>/06_EXECUTION_STATE.md
    current task file / microtask
    Revit XML / journal / log excerpts
    source excerpts
    current user request
    (no breakpoint)
```

## Verification rule for integrations

When an integration adopts this tier map, verify prompt caching empirically:
- first request should report cache-creation tokens for the cacheable prefix
- repeated requests with the same prefix should report cache-read tokens and lower effective repeated input cost

If those counters do not move, the prefix changed or the assembly order drifted.

## Boundary

This dossier defines an ArcTool convention for direct Anthropic API integrations and any custom scripts that assemble ArcTool context themselves.

It does not describe a built-in Claude Code repository setting, and it does not change ArcTool source behavior by itself.
