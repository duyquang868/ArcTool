---
name: feedback_tool_approval_before_editing
description: Ask for approval before editing files when tool use is blocked or when the user may only want a report first rather than immediate file changes.
metadata: 
  node_type: memory
  type: feedback
  originSessionId: c192776f-8d9b-43de-a0c0-a49ea750a110
---

Ask for approval before editing files when tool use is blocked, not permitted in the current chat, or when the user may only want a report/plan first instead of immediate file changes.

**Why:** The user wants more flexibility around tool constraints and prefers being asked whether files should actually be updated when approval is needed.

**How to apply:** If the environment or user instruction blocks tools, or if the user has not clearly approved file edits yet, pause and ask whether to update files rather than assuming direct edits are desired.
