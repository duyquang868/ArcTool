---
name: feedback_claude_md_code_map_review
description: "Before updating CLAUDE.md, review every main directory and file represented in section 2 Code map."
metadata: 
  node_type: memory
  type: feedback
  originSessionId: 11143983-8918-4b8a-b2ac-5ca30a8a0fa6
---

Before updating `CLAUDE.md`, first review every main directory and every file represented in section 2 `Code map`; check whether the actual project structure or tracked file state has changed, then update `CLAUDE.md` only after that review.

**Why:** The user wants `CLAUDE.md` to remain an accurate technical operating document and noticed that section 2 can drift from the real project structure, especially UI files.

**How to apply:** For any future `CLAUDE.md` edit, inspect the Code map directories/files against the real project tree and relevant change status before editing. Pay special attention to main folders such as `Commands`, `Services`, `UI`, `Models`, `Utilities`, `Resources`, and `Properties`; do not update documentation from memory alone.
