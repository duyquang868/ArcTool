---
name: feedback_verify_file_tail_after_edits
description: Verify final file tail and syntax structure after edits; avoid append-based repairs that can duplicate class endings.
metadata:
  node_type: memory
  type: feedback
  originSessionId: 2a0f18ea-a36f-48bd-bd4a-f2e1f45f62bd
---

After editing source files, especially large C# files, verify the final file tail and closing structure before reporting success. Do not repair truncated or incomplete edits by blindly appending tail blocks; re-read the affected region and replace the exact broken range instead.

**Why:** In Quick Dimension Phase 2.4, a previous edit duplicated the `WallSideFaceCandidate` class tail after a valid file ending in `QuickDimensionWallCandidateCollector.cs`, which caused Visual Studio to report about 41 cascading syntax errors. The user had to remove the duplicate block manually.

**How to apply:** For future ArcTool source edits, run scoped syntax/structure checks and inspect the final 30-80 lines of modified files after any tail repair. If a file was modified by a linter or another tool, re-read the affected region before editing again and avoid appending content unless creating a new file from scratch.
