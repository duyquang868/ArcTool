---
name: project-visual-studio-msbuild-build
description: ArcTool must be built with Visual Studio MSBuild for COM reference resolution; dotnet build fails on ResolveComReference.
metadata:
  type: project
---
ArcTool build verification should use Visual Studio MSBuild, not `dotnet build`, because the project contains COM references and `dotnet build` fails with `MSB4803: ResolveComReference is not supported on the .NET Core version of MSBuild`.

Use Git Bash-safe dash switches to avoid `/p` and `/nologo` path conversion:

```bash
"/c/Program Files/Microsoft Visual Studio/18/Community/MSBuild/Current/Bin/MSBuild.exe" ArcTool.Core/ArcTool.Core.csproj -property:Configuration=Debug -property:Platform=x64 -verbosity:minimal -nologo
```

If the exact Visual Studio path is unknown, locate it with:

```bash
"/c/Program Files (x86)/Microsoft Visual Studio/Installer/vswhere.exe" -latest -requires Microsoft.Component.MSBuild -find MSBuild/**/Bin/MSBuild.exe
```

**Why:** ArcTool.Core builds successfully with Visual Studio MSBuild, while `dotnet build ArcTool.Core/ArcTool.Core.csproj -c Debug --no-restore` fails on `ResolveComReference` in this environment.

**How to apply:** For future ArcTool compile verification, use `vswhere` only if the MSBuild path is unknown; otherwise run the Git Bash-safe MSBuild command above. [[project-codebase-memory-repo-local-workflow]]
