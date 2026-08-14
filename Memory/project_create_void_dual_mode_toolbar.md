---
name: project-create-void-dual-mode-toolbar
description: Create Void now uses one compact WPF pre-pick toolbar that combines mode selection and Generic Model family selection.
metadata:
  type: project
---
Create Void (`ArcTool.Core.Commands.CreateVoidFromLinkCommand`) now supports two pre-pick modes through one compact WPF toolbar. This command is considered dormant/stable after the 2026-08-11 implementation snapshot and may sit unchanged for a long period; the long-form record lives in `.Dossier/Detailed Technical Dossier - Create Void.md`.
- `From Link` = keep the old bulk behavior and create voids for all linked beams from the chosen Revit link.
- `From Selected` = let the operator pick only specific linked beams, then confirm with Revit Finish.

The pre-pick UI is `ArcTool.Core/UI/CreateVoidModeToolbar.xaml` plus code-behind `CreateVoidModeToolbar.xaml.cs`. It replaces the clunky mode prompt path by combining mode choice and `OST_GenericModel` `FamilySymbol` selection into one compact dialog.

Accepted UI direction from this session:
- minimal layout
- short radio text only: `From Link`, `From Selected`
- centered family display text in the ComboBox
- rounded Start button corners
- reduced window footprint to fit the control group (`Width="220"`, `SizeToContent="Height"`, smaller ComboBox and button sizing)

The selection/runtime split in `CreateVoidFromLinkCommand` is:
- bulk mode uses link-instance pick flow
- selected mode uses linked-element beam pick flow
- shared void creation pipeline stays unchanged after the toolbar closes

Build verification in this repo still uses the locked Visual Studio MSBuild command from [[project-visual-studio-msbuild-build]].

**Why:** The user wanted a more professional interaction than a Yes/No mode dialog and rejected the first rough toolbar pass until the layout became compact and minimal.

**How to apply:** For future Create Void work, treat the WPF toolbar as the single pre-pick decision surface, preserve the two-mode behavior, and keep the radio labels/layout minimal unless the operator explicitly changes the UX direction. [[project-visual-studio-msbuild-build]]
