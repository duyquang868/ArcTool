# Detailed Technical Dossier — Create Void
**Last updated:** 2026-08-11
**Status:** CLOSED implementation snapshot for current Create Void UX/selection behavior

---

## 1. Scope

This dossier records the current stable implementation shape of the ArcTool Create Void command after the linked-beam selection enhancement and compact toolbar refinement completed in the 2026-08-11 session.

It exists so the command can stay dormant for a long period without losing the reasoning, UX direction, or build/runtime constraints that now define its expected behavior.

---

## 2. Entry point and source surface

Primary command:
- `ArcTool.Core/Commands/CreateVoidFromLinkCommand.cs`

Supporting UI:
- `ArcTool.Core/UI/CreateVoidModeToolbar.xaml`
- `ArcTool.Core/UI/CreateVoidModeToolbar.xaml.cs`

Related existing cut command reference:
- `ArcTool.Core/Commands/MultiCutCommand.cs`

---

## 3. Current operator flow

### 3.1 Pre-pick toolbar

Before any Revit pick starts, the command shows one compact WPF toolbar.

The toolbar combines:
- mode choice
- `OST_GenericModel` void family selection

Accepted minimal UI direction:
- two short radio options only:
  - `From Link`
  - `From Selected`
- one centered family selection list
- one bottom confirmation button: `Start`
- rounded Start button corners
- reduced footprint sized to the control group rather than a larger dialog shell

Current compact sizing in XAML:
- `Width="220"`
- `SizeToContent="Height"`
- smaller ComboBox and Start button heights than the earlier prototype

### 3.2 Mode behavior

#### Mode A — `From Link`

This preserves the original bulk behavior:
- operator picks one `RevitLinkInstance`
- command collects all `OST_StructuralFraming` instances from that linked document
- command attempts to create one void instance per linked beam

#### Mode B — `From Selected`

This adds selective linked-beam processing:
- operator uses linked-element selection on beams inside a Revit link
- operator can select multiple linked beams
- operator confirms selection with Revit Finish
- command processes only those chosen linked beams
- selections are constrained to one link instance per run

---

## 4. Technical implementation shape

### 4.1 Command split

`CreateVoidFromLinkCommand` is now structured around three concerns:
1. collect `FamilySymbol` candidates from `OST_GenericModel`
2. show the WPF toolbar and resolve mode + selected family
3. execute one of two selection paths before entering the shared void-generation pipeline

### 4.2 Selection paths

Bulk path:
- method: `PromptForLinkBulkMode(...)`
- Revit selection type: `ObjectType.Element`
- filter: `LinkSelectionFilter`

Selected-beam path:
- method: `PromptForLinkedBeamsMode(...)`
- Revit selection type: `ObjectType.LinkedElement`
- filter: `LinkedBeamSelectionFilter`
- linked beam identity comes from `Reference.LinkedElementId`

### 4.3 Shared void-generation pipeline

After selection is resolved, both modes reuse the same downstream logic:
- activate the chosen `FamilySymbol`
- resolve beam `LocationCurve`
- transform linked endpoints into host coordinates
- compute midpoint and beam direction
- read width/height parameters from instance first, then symbol fallback
- extract a top/bottom `PlanarFace` from linked beam geometry
- convert face reference through `CreateLinkReference(linkInstance)`
- place the void family with `doc.Create.NewFamilyInstance(...)`
- assign `Width`, `Height`, and `Length`

No architectural rewrite of the geometry or placement logic was introduced in this session.

---

## 5. UX decisions now considered stable

1. Do not return to a separate Yes/No mode dialog.
2. Do not split family selection into a second pre-pick form.
3. Keep the pre-pick surface minimal and compact.
4. Keep radio text short.
5. Keep the family display centered in the visible picker field.
6. Keep the button styling slightly polished but restrained.

These decisions came directly from operator feedback during iterative UI refinement.

---

## 6. Verification state

Static verification achieved in this session:
- source updated for dual-mode behavior
- compact WPF toolbar present and wired into the command
- final XAML reduced to the requested smaller footprint
- repository build PASS using locked Visual Studio MSBuild command

Build command of record:
```bash
"/c/Program Files/Microsoft Visual Studio/18/Community/MSBuild/Current/Bin/MSBuild.exe" ArcTool.Core/ArcTool.Core.csproj -property:Configuration=Debug -property:Platform=x64 -verbosity:minimal -nologo
```

Not performed in this session:
- Revit launch
- runtime smoke test
- `.rvt` execution validation
- stale-addin deployment investigation inside Revit

Important runtime note:
- if Revit still displays the older radio label text or the earlier larger toolbar, that discrepancy should be treated first as a stale build / stale loaded artifact problem, not as evidence that the current XAML source is still wrong.

---

## 7. Maintenance notes for future return

- This command may remain untouched for a long period.
- When it is revisited, start from this dossier and the command/UI files rather than from `.handoff/`.
- Preserve the current compact-toolbar concept unless there is a new explicit UX request.
- Use Visual Studio MSBuild for compile verification; `dotnet build` is not the trusted path for this repo because of COM reference handling.

---

## 8. Related durable records

- `Memory/project_create_void_dual_mode_toolbar.md`
- `Memory/project_visual_studio_msbuild_build.md`
- `.handoff/HANDOFF_TO_NEXT_SESSION.md` (session transfer only; not the long-term source of truth for this feature)
