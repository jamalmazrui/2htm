---
title: "2htm — Convert Documents to Accessible HTML"
author: "Jamal Mazrui"
description: "Convert Documents to Accessible HTML"
---

# 2htm

**Author:** Jamal Mazrui
**License:** MIT

`2htm` is one of three companion accessibility tools by Jamal Mazrui:

- **2htm** — convert documents (Word, Excel, PowerPoint, PDF, Markdown) to accessible HTML
- **extCheck** — check Office and Markdown files for accessibility problems
- **urlCheck** — check web pages for accessibility problems

The three tools share a common command-line and GUI layout, so learning one makes the others easy to pick up.

`2htm` is a Windows tool that converts documents in several formats (Microsoft Word, Excel, PowerPoint, PDF, and Pandoc Markdown) into accessible HTML files. For each input file you give it, 2htm writes a `.htm` companion file alongside the source. The output preserves headings, lists, tables, and image alternative text in a structure that screen readers and other assistive technologies can navigate.

Like its companion tools, `2htm` runs in two modes: a **GUI mode** (a small parameter dialog launched by double-clicking the program, pressing its desktop hotkey, or running with `-g`) and a **command-line mode** (any other invocation, suitable for batch files and pipelines). Both modes accept the same options.

---

## What you need

- Windows 10 or later (64-bit)
- Microsoft Word, Excel, or PowerPoint installed to convert the corresponding `.docx`, `.xlsx`, or `.pptx` files
- No Office installation needed for `.md` or `.pdf` files

You do **not** need to install .NET separately. The .NET Framework 4.8.1 used by `2htm` ships in-box with Windows 10 (since version 22H2) and Windows 11.

**Bitness note.** `2htm` is built as a 64-bit program, and Microsoft Office automation requires the controller process and the installed Office to share the same bitness. Modern Office (Microsoft 365, Office 2019+, Office 2024) is 64-bit by default, so this matches the common case. If you have 32-bit Office on your machine, `2htm` will surface a clear error pointing at the bitness mismatch; you can either install 64-bit Office or rebuild `2htm` with `/platform:x86` (see Development below).

---

## Installing

Download `2htm_setup.exe` from the [GitHub repository](https://github.com/JamalMazrui/2htm) and run it. The setup wizard:

- Prompts you for the installation directory (default: `C:\Program Files\2htm`).
- Includes a brief MIT license summary on the welcome page; the full license text is installed alongside the program as `License.htm`.
- Adds a Start-menu shortcut and a desktop shortcut whose hotkey is **Alt+Ctrl+2**. Pressing **Alt+Ctrl+2** from anywhere in Windows opens the `2htm` dialog.
- The installer also adds a **Convert via 2htm** entry to the File Explorer right-click menu for all file types. Right-clicking any supported file (or pressing **Shift+F10**) and choosing this entry is the fastest way to convert a single file. The accelerator letter is **2** (matching the **Alt+Ctrl+2** desktop hotkey), so the keyboard shortcut from a focused file in File Explorer is **Shift+F10**, then **2**.

The final wizard page offers two checkboxes (both checked by default): launch `2htm` (with a hotkey reminder) and read the HTML documentation.

---

## Running 2htm

### From the dialog (easiest)

Launch `2htm` from any of these:

- The desktop shortcut (or its **Alt+Ctrl+2** hotkey)
- The Start-menu shortcut
- Double-clicking `2htm.exe` in File Explorer
- A Run dialog (`Win+R`) typing `2htm`

The parameter dialog has these controls. Each label has an underlined letter that you can press with **Alt** to jump straight to that control:

- **Source files** [S] — a single file path, a wildcard pattern (e.g., `*.docx`), or several of either separated by spaces. A single path containing spaces does not need quotes — 2htm recognizes the entire trimmed field as one path when it points to an existing file or directory. Quotes are only needed when supplying multiple specs and at least one contains a space.
- **Browse source...** [B] — pick a single source from a file picker
- **Output directory** [O] — where the output is written. Blank means the current working directory.
- **Choose output...** [C] — pick the output directory from a folder picker
- **Strip images** [I] — drop image references from the output
- **Plain text** [P] — produce plain-text `.txt` output instead of HTML
- **Force replacements** [F] — overwrite an existing `<basename>.htm` instead of skipping the input. Without this, 2htm skips an input whose .htm already exists in the output directory.
- **View output** [V] — open the output directory in File Explorer when the run is done
- **Log session** [L] — write a fresh `2htm.log` in the output directory (or current directory if no output directory is set)
- **Use configuration** [U] — load these field values from the saved configuration at startup, and save them back when you press OK
- **Help** [H] — show this help summary and offer to open the full README. F1 also shows Help.
- **Default settings** [D] — clear all fields, uncheck all boxes, and delete the saved configuration if any
- **OK** / **Cancel** — start the run, or cancel without running. Enter is OK; Esc is Cancel.

The Browse source and Choose output pickers open at the directory derived from the corresponding text field's current value when that value points to an existing path; otherwise they open at your Documents folder. With **Use configuration** checked, those text fields are pre-populated from your last session, so the pickers naturally pick up where you left off.

If you press OK with an output directory that does not yet exist, 2htm prompts to create it (default Yes). Choosing No keeps the dialog open with focus on the output field so you can correct it.

When all files have been processed, a final results dialog summarizes what was done.


### From the command line

Open a Command Prompt and run `2htm` with the source as an argument:

```cmd
# Convert one file:
2htm report.docx

# Several files at once:
2htm *.docx *.md

# Files in different folders:
2htm docs\*.docx data\*.xlsx

# Plain text instead of HTML:
2htm -p article.md

# Open the GUI:
2htm -g

```

When invoked without arguments from a GUI shell (Explorer double-click, Start-menu shortcut, desktop hotkey), `2htm` shows the dialog automatically. When invoked without arguments from a console shell, it prints help and exits. The `-g` flag forces GUI mode regardless.

---

## Command-line options

| Option | Long form | Description |
|---|---|---|
| `-h` | `--help` | Show usage and exit |
| `-v` | `--version` | Show version and exit |
| `-g` | `--gui-mode` | Show the parameter dialog |
| `-o <d>` | `--output-dir <d>` | Write output to `<d>` (created if missing); defaults to current directory |
| `-f` | `--force` | overwrite an existing `<basename> |
|   | `--view-output` | After the run, open the output directory in File Explorer |
| `-l` | `--log` | Write `2htm.log` (UTF-8 with BOM) in the output directory; replaced each session |
| `-u` | `--use-configuration` | Read saved defaults from `%LOCALAPPDATA%\2htm\2htm.ini` |
| `-s` | `--strip-images` | Drop image references from the output |
| `-p` | `--plain-text` | Produce plain-text `.txt` output instead of HTML |

Every option in the GUI corresponds one-to-one with a command-line flag, so a workflow prototyped in the dialog can be translated to a batch file without surprises.

---

## Supported input formats

| Extension | Format |
|---|---|
| .docx | Microsoft Word document |
| .xlsx | Microsoft Excel workbook |
| .pptx | Microsoft PowerPoint presentation |
| .pdf | PDF document |
| .md | Pandoc Markdown file |

---

## Output

For each file converted, an HTML file named `<basename>.htm` is written next to the input (or to the output directory if `-o` is given). The HTML preserves heading structure, lists, tables, and image alternative text. CSS is embedded inline so the file can be opened in any browser without external dependencies.

In **plain text** mode (`-p` or the dialog's Plain text checkbox), the output is a `.txt` file instead of `.htm`. Image lines are stripped and the text is normalized for use as input to a synthesizer or for paste-into-email scenarios.

---

## Configuration file

When **Use configuration** is checked in the dialog (or `-u` is on the command line), `2htm` reads and writes a small INI file at:

```
%LOCALAPPDATA%\2htm\2htm.ini
```

It stores the source field, the output directory, and the option checkboxes. Without **Use configuration**, `2htm` leaves nothing on disk between runs. **Default settings** in the dialog deletes this file.

---

## Log file

When **Log session** is checked (or `-l` is on the command line), `2htm` writes a fresh `2htm.log` to the output directory (or current directory if no output directory is set). Any prior log is replaced at the start of the run, so the file always reflects only the current session.

The log captures: program version, command-line arguments, GUI auto-detection, the resolved output directory, per-file events, and any errors (including tracebacks for unexpected failures).

Without **Log session**, `2htm` does not create any log or error file on disk. Errors are reported only to the console (and the GUI results dialog, in GUI mode).

The log is UTF-8 with a byte-order mark, so Notepad opens it correctly.

---

## Notes

- 2htm output preserves heading structure, lists, tables, and image alternative text. The output is a single self-contained `.htm` file with CSS embedded inline; it can be opened in any browser without external dependencies.
- For `.docx`, `.xlsx`, and `.pptx` files, 2htm uses Microsoft Office's COM automation. Office must be installed and runnable; modern (64-bit) Office is the common case.
- For `.md` files, 2htm uses the Markdig library (bundled inside the executable). No Office is needed.
- Plain text mode (`-p` or the dialog's Plain text checkbox) produces a `.txt` file instead. Image lines are stripped and the text is normalized for use as input to a synthesizer or for paste-into-email scenarios.

---

## Development

This section is for developers who want to build the executable from source. End users can skip it.

### Distribution layout

The runtime distribution shipped by `2htm_setup.exe` is just a few files: `2htm.exe` plus the HTML documentation (`ReadMe.htm`, `Announce.htm`, `License.htm`). The Markdown sources, the build script, the installer script, the icon, the program source, and the coding-style guide live in the GitHub repository (and in this `2htm.zip` archive).

### Source layout

The whole program is one C# file: `2htm.cs`. It uses standard `System.Windows.Forms` for the parameter dialog, the COM `dynamic` keyword to drive Office, and the [Markdig](https://github.com/xoofx/markdig) library (downloaded automatically by the build script) for Markdown rendering. PDF support is via the bundled `PdfPig` library; PowerPoint via Office COM. The classes inside `2htm.cs` are arranged as a shared infrastructure layer (`issue`, `results`, `shared`, `comHelper`, `logger`, `configManager`, `guiDialog`) plus per-format converter classes, with a top-level `program` class that parses arguments, optionally shows the dialog, and dispatches.

### Coding style

The source uses what the author calls "Camel Type" (C# variant): Hungarian prefix notation for variables (`b` for boolean, `i` for integer, `s` for string, `ls` for `List<T>`, `d` for dictionary, etc.), lower camelCase for everything other than where the language requires PascalCase. The `o` prefix is reserved for COM objects only; managed C# objects use the lowercase class name as their prefix (e.g., `Form form`, `OpenFileDialog dialog`). Constants follow the same naming as variables — only the `const` or `static readonly` keyword conveys constant-ness. See `Camel_Type_C#.md` in this archive for the full guidelines.

### Threading and bitness

`Main` is decorated with `[STAThread]`. This is required for two reasons:

- Office COM automation requires a single-threaded apartment. Without it, Word/Excel/PowerPoint COM servers can disconnect mid-operation with HRESULT 0x80010108 (RPC_E_DISCONNECTED) or 0x80010114 (OLE_E_OBJNOTCONNECTED).
- WinForms common dialogs (`OpenFileDialog`, `FolderBrowserDialog`) require an STA thread.

The build is `/platform:x64`. Office COM automation requires the controller process and the installed Office to share the same bitness. Modern Office is 64-bit by default; if a user has 32-bit Office, `com.createApp` surfaces a clear error message pointing at the mismatch and recommending a 32-bit rebuild.

### Prerequisites

- The .NET Framework 4.8.1 Developer Pack (provides `csc.exe` and the 4.8.1 reference assemblies). Install from <https://dotnet.microsoft.com/download/dotnet-framework/net481>.
- Inno Setup 6.x to compile the installer.

### Building the executable

Run the included script:

```cmd
build2htm.cmd
```

It auto-detects the compiler, verifies the build environment, embeds the icon into `2htm.exe`, and produces the runtime distribution in `dist\`.

### Building the installer

Open `2htm_setup.iss` in Inno Setup and click Compile. The result is `dist\2htm_setup.exe`.

The installer ships only the runtime files: `2htm.exe` plus the HTML documentation (`ReadMe.htm`, `Announce.htm`, `License.htm`). Markdown sources, the build script, this `.iss` script, the icon, the source file, and any coding-style guideline files live in the GitHub repository.

### Uninstalling

Use Apps & Features in Windows Settings, or run the uninstaller from the `2htm` Start-menu group. The uninstaller removes the program files. It does not touch `%LOCALAPPDATA%\2htm\2htm.ini` or any `2htm.log` files in working directories — delete those manually if you want a fully clean removal.


## License

MIT License. See `License.htm` (installed alongside the program) or `License.md` (in the GitHub repository).