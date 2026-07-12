---
name: backend-scope-preference
description: Report unsupported backend scope immediately and do not add placeholder UI fields as a workaround
type: feedback
originSessionId: c192776f-8d9b-43de-a0c0-a49ea750a110
---
When the requested UI depends on backend support that does not exist yet, report that gap immediately instead of adding extra informational fields or other UI workarounds.

**Why:** The user prefers not to clutter the dialog with redundant fields when the real issue is missing backend support.

**How to apply:** Use this rule whenever a requested dialog option, trigger scope, or settings field would imply support that the backend cannot actually deliver yet.
