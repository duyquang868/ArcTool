---
name: feedback_nullable_annotations_revit_api
description: Under #nullable enable, propagate nullable Revit API returns with nullable annotations instead of assigning to non-nullable locals.
metadata:
  node_type: memory
  type: feedback
  originSessionId: 2a0f18ea-a36f-48bd-bd4a-f2e1f45f62bd
---

When editing ArcTool C# files with `#nullable enable`, treat null-conditional Revit API calls and possibly-null API returns as nullable locals, then narrow with pattern checks. For example, `element?.GetGeometryObjectFromReference(faceReference)` should be assigned to `GeometryObject?`, not `GeometryObject`, before checking `is not PlanarFace`.

**Why:** In Quick Dimension Phase 2.4, `GeometryObject geometryObject = element?.GetGeometryObjectFromReference(faceReference);` in `QuickDimensionWallCandidateCollector.cs` produced CS8600 because the null-conditional call can return null. The user fixed it by changing the local to `GeometryObject?`, preserving behavior because the next line already pattern-checks the value.

**How to apply:** For future ArcTool edits, scan changed `#nullable enable` files for CS8600-style assignments from `?.` and nullable-returning Revit API calls. Prefer nullable locals plus immediate guards/pattern matching over suppressions or broad nullable-disable workarounds.
