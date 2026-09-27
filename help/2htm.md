---
title: "2htm Guide"
author: "Jamal Mazrui"
description: "Convert Documents to Accessible HTML"
---

# 2htm Guide

This is the full guide. For a quick start, see ReadMe.

**Author:** Jamal Mazrui
**Copyright:** © 2026 Jamal Mazrui
**License:** [MIT](https://opensource.org/license/mit/)
**Project home:** <https://github.com/JamalMazrui/2htm>

`2htm` is a Windows tool that converts documents in several formats (Microsoft Word, Excel, PowerPoint, PDF, and Pandoc Markdown) into accessible HTML files. For each input file you give it, 2htm writes a `.htm` companion file alongside the source. The output preserves headings, lists, tables, and image alternative text in a structure that screen readers and other assistive technologies can navigate.

Like its companion tools `urlCheck` and `extCheck` (see the Announce file for a description of the family), `2htm` runs in two modes: a **GUI mode** (a small parameter dialog launched by double-clicking the program, pressing its desktop hotkey, or running with `-g`) and a **command-line mode** (any other invocation, suitable for batch files and pipelines). Both modes accept the same options.

---

## What you need

- Windows 10 or later (64-bit)
- Microsoft Word, Excel, or PowerPoint installed to convert the corresponding `.docx`, `.xlsx`, or `.pptx` files
- Word 2013 or later for `.pdf` files, which 2htm converts through Word's PDF Reflow
- No Office installation needed for `.md` files

You do **not** need to install .NET separately. The .NET Framework 4.8.1 used by `2htm` ships in-box with Windows 10 (since version 22H2) and Windows 11.

**Bitness note.** `2htm` is built as a 64-bit program, and Microsoft Office automation requires the controller process and the installed Office to share the same bitness. Modern Office (Microsoft 365, Office 2019+, Office 2024) is 64-bit by default, so this matches the common case. If you have 32-bit Office on your machine, `2htm` will surface a clear error pointing at the bitness mismatch; you can either install 64-bit Office or rebuild `2htm` with `/platform:x86` (see Development below).

---

## Installing

Download `2htm_setup.exe` from the [2htm releases page](https://github.com/jamalmazrui/2htm/releases) and run it. It needs administrator rights, because it installs for everyone on the computer. The setup wizard:

- Asks for the installation folder the first time (default: `C:\Program Files\2htm`). An update goes to the same folder without asking.
- Shows a short MIT license summary on the welcome page; the full license is installed as `License.htm`.
- Adds a Start-menu group and a desktop shortcut whose hotkey is **Alt+Control+2**. Pressing **Alt+Control+2** from anywhere in Windows opens the `2htm` dialog. (While Word has focus, Word keeps Alt+Control+2 for its Heading 2 style.)
- Adds a **Convert via 2htm** entry to the File Explorer right-click menu for all file types. Right-clicking a supported file (or pressing **Shift+F10**) and choosing this entry is the fastest way to convert one file; the result lands beside it. The access key is **2**, matching the desktop hotkey, so from a file in File Explorer the keys are **Shift+F10**, then **2**.

The last page offers two checkboxes. **Launch 2htm now** is checked; **Open the user guide** is not. When you press Finish, a short results box says what was installed and where the logs are. 2htm starts after you close that box, so the box is never hidden behind it.

The program itself is in the `exec` folder of the installation, and this guide and the other documents are in the `help` folder. `ReadMe.htm`, `License.htm` and `2htm.cmd`, which runs the program from a command prompt there, are at the top.

### Updating

In the dialog, press **F11**. 2htm asks GitHub whether a newer version exists. If one does, Yes is the default: press Enter and the new setup program is downloaded and started. If you already have the newest version, No is the default. If the check fails, for example with no internet connection, a message says so.

---

## Running 2htm

### From the dialog (easiest)

Launch `2htm` from any of these:

- The desktop shortcut (or its **Alt+Control+2** hotkey)
- The Start-menu shortcut
- Double-clicking `2htm.exe` in the `exec` folder in File Explorer

The parameter dialog has these controls. Each label has an underlined letter that you can press with **Alt** to jump straight to that control:

- **Source files** [S] — a single file path, a wildcard pattern (e.g., `*.docx`), or several of either separated by spaces. A single path containing spaces does not need quotes — 2htm recognizes the entire trimmed field as one path when it points to an existing file or directory. Quotes are only needed when supplying multiple specs and at least one contains a space.
- **Browse source...** [B] — pick a single source from a file picker
- **Output directory** [O] — where the output is written. Blank means the current working directory.
- **Choose output...** [C] — pick the output directory from a folder picker
- **Strip images** [I] — drop image references from the output
- **Plain text** [P] — produce plain-text `.txt` output instead of HTML
- **Force replacements** [F] — overwrite an existing `<basename>.htm` instead of skipping the input. Without this, 2htm skips an input whose .htm already exists in the output directory.
- **View output** [V] — open the output directory in File Explorer when the run is done
- **Log session** [L] — also write a fresh `2htm.log` in the output directory (or current directory if no output directory is set). A session log is always kept in `%LOCALAPPDATA%\2htm\logs`.
- **Use configuration** [U] — load these field values from the saved configuration at startup, and save them back when you press OK
- **Guide** [G] — open this guide in your browser, then return to the dialog with everything as it was.
- **Help** [H] — list every field with its tip, the dialog's keys, and the version check. F1 also shows Help.
- **Default settings** [D] — clear all fields, uncheck all boxes, and delete the saved configuration if any
- **OK** / **Cancel** — start the run, or cancel without running. Enter is OK; Escape is Cancel.

More keys work anywhere in the dialog: **F1** shows Help, **F11** checks the web for a newer version of 2htm (see Updating), **Control+Enter** is OK from any control, **Shift+F1** says the tip for the field you are on, and **F7** lists the dialog's controls so you can move to one. The dialog is built with the Homer Lbc classes, so its text fields also have the Lbc editing keys, which Help lists.

**Default settings** and **Guide** return you to the dialog; Default settings first clears the fields and deletes any saved configuration.

The Browse source and Choose output pickers open at the directory derived from the corresponding text field's current value when that value points to an existing path; otherwise they open at your Documents folder. With **Use configuration** checked, those text fields are pre-populated from your last session, so the pickers naturally pick up where you left off.

If you press OK with an output directory that does not yet exist, 2htm prompts to create it (default Yes). Choosing No keeps the dialog open with focus on the output field so you can correct it.

When all files have been processed, a final results dialog summarizes what was done.


### From the command line

Open a Command Prompt in the 2htm program folder, or put that folder on your PATH, and run `2htm` with the source as an argument. `2htm.cmd` there runs the program in the `exec` folder.

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

- `-f`, `--force` — overwrite an existing `<basename>.htm` instead of skipping the input.
- `-g`, `--gui-mode` — show the parameter dialog.
- `-h`, `--help` — show usage and exit.
- `-l`, `--log` — also write `2htm.log` (UTF-8 with BOM) in the output directory, replaced each session. A session log is always kept in `%LOCALAPPDATA%\2htm\logs`.
- `-o <folder>`, `--output-dir <folder>` — write output to `<folder>` (made if missing); the default is the current directory.
- `-p`, `--plain-text` — produce plain-text `.txt` output instead of HTML.
- `-s`, `--strip-images` — drop image references from the output.
- `-u`, `--use-configuration` — read saved settings from `%LOCALAPPDATA%\2htm\configs\2htm.inix`.
- `-v`, `--version` — show the version and exit.
- `--view-output` — after the run, open the output directory in File Explorer.

Every option in the GUI corresponds one-to-one with a command-line flag, so a workflow prototyped in the dialog can be translated to a batch file without surprises.

---

## Supported input formats

- `.docx` — Microsoft Word document
- `.md` — Pandoc Markdown file
- `.pdf` — PDF document
- `.pptx` — Microsoft PowerPoint presentation
- `.xlsx` — Microsoft Excel workbook

---|---|
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

When **Use configuration** is checked in the dialog (or `-u` is on the command line), `2htm` reads and writes its settings in:

```
%LOCALAPPDATA%\2htm\configs\2htm.inix
```

It stores the source field, the output directory, and the option checkboxes, one per line under a `[Settings]` heading, such as `ViewOutput=yes`. You can edit it in any text editor; a comment you add is kept when 2htm saves. **Default settings** in the dialog deletes the file.

Settings saved by versions before 1.19, in `%LOCALAPPDATA%\2htm\2htm.ini`, are read when there is no `.inix` yet. The old file is removed the first time the new one is saved.

---

## Log files

2htm always keeps a log of each session, whatever the options:

```
%LOCALAPPDATA%\2htm\logs\2htm-yyyyMMdd-HHmmss.log
```

Each run gets its own file, named for the moment it started, so sorting the folder by name also sorts it by time. The 30 newest are kept. The log starts with the version, where the program ran from, the Windows version, the working folder, the command line and every setting, then records each file converted and any error with its full details. If something goes wrong, zip the `logs` folder and send it.

**Log session** (or `-l`) additionally writes a fresh `2htm.log` in the output directory (or the current directory if none is set), so a log can travel with the output. Any prior `2htm.log` there is replaced.

Both logs are UTF-8 with a byte-order mark, so Notepad opens them correctly.

---

## Notes

- 2htm output preserves heading structure, lists, tables, and image alternative text. The output is a single self-contained `.htm` file with CSS embedded inline; it can be opened in any browser without external dependencies.
- For `.docx`, `.xlsx`, and `.pptx` files, 2htm uses Microsoft Office's COM automation. Office must be installed and runnable; modern (64-bit) Office is the common case.
- For `.md` files, 2htm uses the Markdig library (bundled inside the executable). No Office is needed.
- Plain text mode (`-p` or the dialog's Plain text checkbox) produces a `.txt` file instead. Image lines are stripped and the text is normalized for use as input to a synthesizer or for paste-into-email scenarios.

---

## Uninstalling

Use Installed apps in Windows Settings, or the Uninstall shortcut in the 2htm Start-menu group. The uninstaller removes the program, its logs and the File Explorer entry. Your saved settings, `configs\2htm.inix`, are left in place, as they always have been; delete them by hand if you want 2htm gone completely.

---

## For developers

How 2htm is built, released and structured is in Developer, in this folder.

## License

MIT License. See `License.htm` at the top of the program folder.
