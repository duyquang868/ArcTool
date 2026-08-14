# ArcTool — Revit API Reference Notes

Last updated: 2026-08-07
Status: Durable lookup register. Extracted from `CLAUDE.md` section 9 during the change-5 thin-core split; `CLAUDE.md` now keeps only the reference-of-record URL, the no-guesswork rule, and a pointer here.

Reference of record: https://www.revitapidocs.com/2026/

**Read this file when** you are about to touch collectors, coordinate extraction/conversion, updater lifecycle, shared-parameter registration or repair, Filter Manager, or the Excel import / legend path. Skip it for edits that do not call new Revit APIs.

Rule that stays at the root: always look up Revit API and other technical docs before answering or fixing a bug; no guesswork. If no reliable source exists, say so and request human help.

---

## Collection and element access

- `FilteredElementCollector` — performance-sensitive collection pipeline; apply quick filters before slow/LINQ filters.
- `LocationPoint` / `LocationCurve` — the only supported placement/extraction paths in this project. Vertical instances use `LocationPoint.Point`; slanted column/foundation instances use the `LocationCurve` start point.

---

## Coordinate pipeline

- `ProjectLocation.GetProjectPosition(XYZ)` — primary coordinate conversion path.
- `BasePoint.GetSurveyPoint(...)` / `BasePoint.GetProjectBasePoint(...)` — coordinate context resolution.

---

## Updater lifecycle

- `IUpdater` — implemented by `CoordinateUpdater`; `Execute()` never opens a transaction.
- `UpdaterRegistry.*` — registration/unregistration; lifecycle is document-scoped.
- `Element.GetChangeTypeGeometry()` — one trigger per registered category.
- `UIControlledApplication.ActiveAddInId` — must be captured in `OnStartup()`; casting a document-event `sender` to `ControlledApplication` silently breaks registration.

---

## Shared parameters

- `Application.SharedParametersFilename`
- `Application.OpenSharedParameterFile()`
- `ExternalDefinitionCreationOptions`
- `DefinitionBindingMapIterator`
- `BindingMap.Insert` / `BindingMap.ReInsert`

Together these form the shared-parameter registration and repair path. `InstanceBinding` is the locked choice when binding many categories.

---

## Filter Manager

- `ParameterFilterElement` — target API for the still-unimplemented copy/paste logic.

---

## Excel to Revit / legend workaround

- `ImageType.Create(...)`, `ImageInstance.Create(...)` — image import path.
- `ViewDrafting.Create(...)` — drafting view creation.
- `View.Duplicate(...)` — legend creation workaround; Revit exposes no direct legend create path.
- `Marshal.ReleaseComObject(...)` — Excel interop hygiene; release wrappers child → parent, never after a COM `Delete()`.

---

## Quick Dimension

- `HostObjectUtils.GetSideFaces()` — wall side-face reference resolution (closest valid side face).
- `FamilyInstance.GetReferences(FamilyInstanceReferenceType.Left/Right)` — door/window jamb references, with `HostWallOpeningGeometry` as fallback.
- `new Reference(grid)` — grid reference strategy (legacy/optional path; no arc-grid support).
- `Document.Create.NewDimension(...)` — must span the resolved final candidate range, not raw selected-wall `0..axisLength`.

Deeper rationale: `.Dossier/ArcTool Locked Technical Decisions.md` and `.Dossier/Quick Dimension - Implementation Roadmap.md`.
