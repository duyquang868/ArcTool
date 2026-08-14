---
name: feedback_worker_scope_enforcement
description: In ArcTool work packages, the master never edits worker-owned source or reruns worker-owned build/audit gates; fixes must be routed back through the owning task worker.
metadata:
  type: feedback
---

In ArcTool work packages, the master must not perform source edits, build reruns, or static-audit reruns that belong to a task's exclusive `write_scope`. If a defect is found in a file owned by task `Tn.m`, the master routes the fix back to that owning worker (or to an explicitly created follow-up task with the same exclusive scope) and waits for that worker's result envelope. The master also must not validate package closure by running the worker-owned gate itself.

2026-08-09 violation: during the Excel-to-Revit WPS provider split Phase 3 closure, the master directly edited `ArcTool.Core/Services/Excel/MsExcelWorkbookPdfExporter.cs` (`T3.1`), `ArcTool.Core/Services/Excel/WpsWorkbookPdfExporter.cs` (`T3.2`), and `ArcTool.Core/Services/Excel/PdfRasterImageService.cs` (`T2.2`), then reran the T3.7 build/audit gate from the master chat instead of returning those fixes to the owning workers and re-dispatching `T3.7`.

**Why:** Crossing worker write scopes invalidates the work-package record, leaves `06_EXECUTION_STATE.md` / result files out of sync with the actual source state, and forces the user to spend extra tokens just to re-close the same phase correctly.

**How to apply:** In any ArcTool work package, treat the manifest write-scope table as a hard boundary. When a gate finds a defect, attribute it to the owning task, dispatch that worker with only the minimum contract/evidence excerpt, and accept only the compact result envelope back in the master chat. After owner fixes land, re-run the gate as its own worker so the package artifacts become authoritative again. This rule complements [[feedback_master_context_discipline]]: the master stays orchestration-only, with `model: "sonnet"` as the default worker pin unless the package records a justified exception.
