---
name: project_gemma4_worker_instruction_system
description: Gemma 4 worker/student instruction system with LM Studio isolated memory, plugin-backed read access, and Claude-curated knowledge files.
type: project
---

Gemma 4 worker/student framework lives in an isolated LM Studio folder: `C:\Users\ADMIN\.lmstudio\gemma4-arctool-memory\`.

A custom LM Studio plugin provides Gemma read-only tool access to this folder. Plugin path: `C:\Users\ADMIN\.lmstudio\extensions\plugins\lmstudio\gemma-arctool-memory\`. Exposed tools are `list_memory_files`, `read_memory_file`, and `search_memory`. The plugin is whitelist-only, read-only, has no write tool, accepts no arbitrary source path, and must not expose the ArcTool source folder.

Because LM Studio treats this local plugin as a dev plugin rather than an installed Hub plugin, it is launched with `lms dev`. The chosen no-Hub operating mode is Windows autostart: `gemma-arctool-memory-plugin.vbs` in Windows Startup runs `start-gemma-arctool-memory.cmd`, which executes `C:\Users\ADMIN\.lmstudio\bin\lms.exe dev` from the plugin folder.

**Files in the memory folder (all maintained by Claude, read-only for Gemma):**
- `gemma4_worker_constitution.md` — locked identity, coding laws, forbidden actions, pre-submission checklist.
- `gemma4_task_delegation_template.md` — standard task/correction turn format Claude uses to delegate.
- `gemma4_error_learning_log.xml` — structured XML log of verified compile-fix/review-fix entries.
- `gemma4_compiled_lessons.md` — condensed actionable rules distilled from XML entries every ~10-20 entries.
- `gemma4_session_knowledge.md` — curated project knowledge (like Claude's own memory system but for Gemma).

**Why:** User wants Gemma to have durable learning and context without accessing the ArcTool source folder. Claude curates and writes; Gemma only reads. This prevents Gemma from corrupting project files while still letting it self-improve.

**How to apply:**
1. Before delegating coding work to Gemma: read constitution + session_knowledge + compiled_lessons + relevant XML entries from the LM Studio folder.
2. Inject relevant content into Gemma's system/prompt via MCP (`mcp__gemma4-lmstudio__ask`). Plugin tools (`read_memory_file`, `search_memory`) are available to Gemma in LM Studio Chat UI but not through MCP API; Claude must inject context when calling through MCP.
3. After correction turns: review Gemma's drafted `<learning_entry>`, append accepted entries to XML log.
4. Every ~10-20 entries or when patterns stabilize: update `gemma4_compiled_lessons.md` and `gemma4_session_knowledge.md`.
5. After each meaningful work session with Gemma: update `gemma4_session_knowledge.md` with new durable project knowledge (same discipline as Claude's own memory writes).
6. ArcTool `Memory/` retains a copy of the XML log and constitution as archive/backup; the LM Studio folder is the live operational copy.
7. If plugin is not running (LM Studio not started, or VBS not in Startup), Claude must fully inject context via MCP as fallback.