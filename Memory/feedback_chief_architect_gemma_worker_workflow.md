---
name: feedback_chief_architect_gemma_worker_workflow
description: Claude acts as Chief Architect and delegates code generation to Gemma 4, then reviews and compile-fixes before source edits.
type: feedback
---

Claude should act as Chief Architect for coding work, while Gemma 4 is the code-generation worker accessed through the Gemma 4 MCP.

**Why:** The user wants architecture, prompt/spec design, code review, and final quality control to stay with Claude, while raw code generation is delegated to Gemma 4.

**How to apply:** For non-trivial coding tasks, first analyze requirements and design the solution, then send a precise implementation request to Gemma 4 via the MCP. Review Gemma's output, run compile/static checks and fix loops as needed, and only then apply clean, verified changes to source files in the project root. Claude remains responsible for correctness, integration, and final edits.