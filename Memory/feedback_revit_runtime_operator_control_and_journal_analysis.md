---
name: revit-runtime-operator-control-and-journal-analysis
description: Never open Revit, invoke Revit MCP, or run smoke tests unless explicitly requested; prefer journal analysis as an independent diagnostic evidence source.
metadata:
  type: feedback
---

Do not autonomously open Autodesk Revit, invoke any Revit MCP tool, or execute a Revit smoke-test matrix. The user runs Revit smoke tests manually and will explicitly request Revit MCP when they want it used.

**Why:** Revit runtime tests are operator-controlled, modal, geometry-sensitive, and the user trusts manually observed results more than autonomous execution. Unrequested launch/MCP attempts waste time and tokens and can interfere with the user's Revit environment.

**How to apply:**
- Source edits, builds, static analysis, XML/image/journal review, and documentation updates may proceed when requested.
- Never launch `Revit.exe`, open an `.rvt`, invoke `revit-mcp`, click ribbon commands, select elements, or attempt runtime smoke without a direct user request for that specific action.
- Do not infer permission from a request to implement or prepare smoke instrumentation. “Prepare the audit” is not permission to run it.
- When runtime proof is required, stop at a precise operator runbook and wait for the user's XML/images/results.
- Revit journal files are a preferred independent evidence source for bug diagnosis and cross-checking. Search recent journals to discover fixture/model paths, command execution, selected element ids, transaction results, dimension ids, warnings, and historical behavior. Correlate journal evidence with source, XML logs, annotated images, and user observations; never treat journal evidence alone as a substitute for a newly required smoke result.
- A failed or absent Revit MCP connection is not a reason to launch Revit. Report the boundary and continue with non-runtime evidence.

Related: [[project_qd_chain_creation_audit_handoff]], [[smoke-test-single-session-close]], [[feedback_tool_approval_before_editing]].
