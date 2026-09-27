---
title: "2htm Developer Notes"
author: "Jamal Mazrui"
---

# 2htm Developer Notes

How 2htm is built, released and laid out. Since 1.19.0 it is built on the Homer Development Kit (HomerDev), kit 1.43.4 or later, in `C:\HomerDev`.

## The project folder

`C:\2htm` mirrors the installed tree:

- At the top: `2htm.cs` (the program), `build2htm.cmd`, `2htm_setup.iss`, `2htm.cmd`, `2htm.ico`, `accept.inix`, `RepoFiles.txt`, `LocalFiles.txt`, `ReadMe` and `License`.
- `exec` — the built `2htm.exe`. Never in git.
- `help` — this document and the others: `2htm` (the guide), `Announce`, `Developer`, `History`, `Hotkeys`, each as `.md` and `.htm`.
- `logs` — one log per run of the build or any tool.
- `scripts` — the kit's tools, refreshed from `C:\HomerDev\scripts` by every build.
- `work\nuget` — Markdig, fetched once. Never in git.

`version.txt` holds the version and lives only on this machine. The build writes `Version.cs` from it, and the installer reads it directly.

## The four steps

1. `build2htm` — steps the version (`build2htm nobump` keeps it), fetches Markdig if it is missing, compiles `exec\2htm.exe`, writes each `.htm` from its `.md`, puts the project's files in the Homer encoding, checks that the installer ships every file in `help`, and builds `2htm_setup.exe`. Its log is `logs\2htm-build-yyyyMMdd-HHmmss.log`.
2. `scripts\push "message"` — rewrites the whitelist `.gitignore` from `RepoFiles.txt`, commits and pushes.
3. `scripts\tidy` and `scripts\tidy --do-it` — the periodic clean.
4. `scripts\release` — runs `scripts\check`, then tags the pushed commit with the version stamped in `2htm_setup.exe` and publishes the installer.

Try the fresh build with `2htm -g` (the `2htm.cmd` at the top runs `exec\2htm.exe`).

## The compiler, the kit's classes and Markdig

The build finds the Roslyn C# compiler with `vswhere`, or installs the free Visual Studio Build Tools with winget; 2htm has always needed Roslyn.

2htm compiles eight of the kit's classes straight from `C:\HomerDev\CSharp`: Elevate, Inix, Lbc, Log, Paths, Say, Util and Web. Its dialog, `guiDialog.show`, is an LbcDialog, built the way urlFido's is: bands for a field and its button, checkboxes below a separator, and `runWithButtons` with OK, Guide, Default settings and Cancel, looping back to the dialog after Guide, Default settings, or an output folder the user declined to create. Lbc needs Elevate, Inix, Log, Say and Util.

- `Log` keeps the session log in `%LOCALAPPDATA%\2htm\logs`. 2htm's own `logger` writes every line to it as well as to the optional `2htm.log`, and `Main` records anything `run` did not catch. Log needs Paths and Say.
- `Paths.configs()` is where `2htm.inix` lives, and `Paths.installedFolder` is how Help finds `help\2htm.htm` from `exec`.
- `InixCodec.writeValue` saves each setting in place.
- `Elevate` answers F11 (and needs Web).

**Markdig 0.18.3** is fetched once into `work\nuget` (lib\net40, else lib\net35) and embedded as a resource, which `resolveEmbeddedAssembly` loads, so 2htm stays one file. `Main` registers that handler before anything else and must not itself reference a Markdig type, because the runtime resolves a method's types as it compiles it; Log and Elevate do not touch Markdig. Markdig's types are declared in `netstandard.dll`, so the build references the netstandard facade: without it csc reports CS0012 against the Markdig call sites, which looks like a fault in the source and is not.

## Source layout

The whole program is one C# file: `2htm.cs`. It uses standard `System.Windows.Forms` for the parameter dialog, the COM `dynamic` keyword to drive Office, and the [Markdig](https://github.com/xoofx/markdig) library (downloaded automatically by the build script) for Markdown rendering. PDF is converted through Word's PDF Reflow, and PowerPoint through Office COM. The classes inside `2htm.cs`, in the `twoHtm` namespace, are arranged as a shared infrastructure layer (`comHelper`, `configManager`, `fileIntegrity`, `guiDialog`, `guiProgress`, `htmlWriter`, `logger`, `tempManager`) plus per-format converter classes, with a top-level `program` class that parses arguments, optionally shows the dialog, and dispatches.

## Threading and bitness

`Main` is decorated with `[STAThread]`. This is required for two reasons:

- Office COM automation requires a single-threaded apartment. Without it, Word/Excel/PowerPoint COM servers can disconnect mid-operation with HRESULT 0x80010108 (RPC_E_DISCONNECTED) or 0x80010114 (OLE_E_OBJNOTCONNECTED).
- WinForms common dialogs (`OpenFileDialog`, `FolderBrowserDialog`) require an STA thread.

The build is `/platform:x64`. Office COM automation requires the controller process and the installed Office to share the same bitness. Modern Office is 64-bit by default; if a user has 32-bit Office, `comHelper.createApp` surfaces a clear error message pointing at the mismatch and recommending a 32-bit rebuild.

## Conventions

- Camel Type for C#, as `C:\HomerDev\help\CamelType_CSharp.md` describes. Constants added since the move to the kit carry the `c_` prefix. The `o` prefix is reserved for COM objects.
- 2htm shares identifier names with urlCheck, urlFido and extCheck: `sProgramName`, `sProgramVersion`, `sConfigFileName`, `sLogFileName`, `sOutputDir`, the `iLayout*` constants, and the `logger` surface `open`, `close`, `info`, `warn`, `error`, `debug`.
