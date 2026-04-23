# 2htm

**Author:** Jamal Mazrui
**License:** MIT (see `License.htm`)

`2htm` is a single, independent Windows executable that converts popular document file types to accessible HTML (WCAG 2.2 AAA) or plain text. It runs on any modern Windows version with Microsoft Office installed. There is no installer, no .NET runtime to download, no sidecar DLLs — just one .exe file that works from any folder.

The same executable runs in two modes with the same set of options: a command-line mode for scripting and batch processing, and a GUI dialog mode for interactive use. A single command-line flag (`-g`) switches between them.

The tool converts each input into a single-file HTML equivalent that can be opened in any modern web browser — Chrome, Edge, Firefox, Safari — with no dependencies, no companion folders, and no special viewer required. The conversion aims for WCAG 2.2 AAA conformance to the extent the source document's structure and content allow. Headings, landmarks, table markup, alt-text propagation, color contrast, and language declaration are all preserved or inferred where possible.

2htm is free to use and free to modify under the permissive MIT license. The source is a single C# file (`2htm.cs`) — C# is Microsoft's flagship application language, well-documented, with a mature free toolchain (Visual Studio Community or Visual Studio Build Tools, both free downloads from Microsoft). Developers can customize or extend the conversion logic without paying for commercial tooling.

This project was developed in collaboration with Claude, Anthropic's AI coding assistant.

The whole 2htm project may be downloaded in a single zip archive from:

<http://GitHub.com/JamalMazrui/2htm/archive/main.zip>

---

## What you need

- Windows 10 or later
- Microsoft Office 2016 or later (Word, Excel, PowerPoint) for converting Office documents and PDFs

Plain text, CSV, JSON, HTML, and Markdown files can be converted without Office installed.

---

## Why 2htm

**Accessibility pipelines.** Organizations publishing documents to the public often need to produce "alternate formats" — versions of content that are more accessible to users with disabilities. Running 2htm as a step in a content pipeline turns Word, Excel, PowerPoint, and PDF files into clean, landmark-rich HTML automatically. The HTML output reads well on screen readers, reflows on small screens, and opens on any device with a browser.

**Single-file portability.** Because 2htm is one executable, it can be dropped into a folder, attached to an email, or stored on a thumb drive. Administrators can deploy it across an organization without installer paperwork. Developers can call it from batch files, scheduled tasks, or CI jobs.

**No vendor lock-in for the output.** A `.htm` file produced by 2htm has no sidecar folder of images or styles. Images are embedded as base64 data URIs; CSS is inlined. The file can be stored, emailed, archived, or served from any web host without worrying about what gets left behind.

**The same interface for GUI and CLI users.** Every option available on the command line is also a field in the GUI dialog. Users who prefer a visual workflow get a keyboard-accessible form (every control has an Alt+Letter hotkey and a tab-order sequence); users who prefer scripting get a conventional POSIX-style command-line interface.

---

## How to use it

Put `2htm.exe` in any folder. Open a Command Prompt in the folder containing the files you want to convert, and run:

```cmd
2htm report.docx
```

The command above creates `report.htm` in the current directory — a fully accessible, single-file HTML document you can open in any browser, share by email, or post to a website.

### Convert one file

```cmd
2htm annual-report.docx
```

### Convert many files at once

```cmd
2htm *.xlsx
```

### Convert every supported file in a folder

```cmd
2htm C:\work\documents
```

A bare folder path processes every file type 2htm recognizes inside that folder. Files with unsupported extensions (images, archives, and so on) are silently skipped.

### Convert to plain text instead of HTML

```cmd
2htm -p report.docx
```

### Open the GUI

```cmd
2htm -g
```

A small dialog lets you pick source files, an output directory, and conversion options. The GUI is designed to work cleanly with screen readers; every field has a keyboard hotkey.

### Get help

```cmd
2htm -h
```

---

## Supported input formats

| Format | Extensions | Notes |
|---|---|---|
| Microsoft Word | `.docx` `.doc` `.rtf` `.odt` | Word automation |
| PDF | `.pdf` | Word 2013+ PDF Reflow |
| Microsoft Excel | `.xlsx` `.xls` | Region-aware tables |
| Microsoft PowerPoint | `.pptx` `.ppt` | One section per slide |
| CSV | `.csv` | Native |
| Web | `.html` `.htm` | Native (cleaned) |
| Markdown | `.md` | Pandoc / CommonMark via Markdig |
| JSON | `.json` | Pretty-printed |
| Text | `.txt` | Native |

---

## Command-line options

| Option | Long form | Description |
|---|---|---|
| `-h` | `--help` | Show usage and exit |
| `-v` | `--version` | Show version and exit |
| `-s` | `--strip-images` | Remove images from output (smaller, faster) |
| `-p` | `--plain-text` | Produce .txt instead of .htm |
| `-f` | `--force` | Overwrite existing output files |
| `-l` | `--log` | Write detailed diagnostics to `2htm.log` |
| `-g` | `--gui-mode` | Launch the dialog |
| `-o <dir>` | `--output-dir <dir>` | Write output to `<dir>` instead of the current directory |
|   | `--view-output` | After conversion, open the output directory in File Explorer |
| `-u` | `--use-configuration` | Read saved defaults from `%LOCALAPPDATA%\2htm\2htm.ini` |

### Examples

Convert every Word and Excel file in two folders, write the output to a third folder, and open that folder in File Explorer when done:

```cmd
2htm -o C:\converted --view-output C:\reports\*.docx C:\finance\*.xlsx
```

Convert a PDF to plain text, overwriting any existing output file:

```cmd
2htm -p -f handbook.pdf
```

Run the GUI with options pre-populated:

```cmd
2htm -g -p -s C:\docs\*.docx
```

---

## GUI mode

Running `2htm -g` opens a small dialog with these controls (keyboard hotkeys in parentheses):

- **Source files** (Alt+S) — a file path, a folder path, or a wildcard pattern. A bare folder processes every supported file in it.
- **Browse source...** (Alt+B) — pick a folder with the Windows folder-chooser dialog.
- **Output directory** (Alt+O) — where converted files are written. Defaults to the source directory.
- **Choose output...** (Alt+C) — pick the output folder.
- **Strip images** (Alt+I) — remove images for smaller, faster output.
- **Plain text** (Alt+P) — write .txt files instead of .htm.
- **Force replacements** (Alt+F) — overwrite existing output.
- **View output** (Alt+V) — open the output directory in File Explorer when done.
- **Use configuration** (Alt+U) — save these settings as defaults for next run.
- **Default settings** (Alt+D) — reset all fields to factory defaults.
- **Help** (Alt+H) — show a quick help message.
- **OK** / **Cancel** — Enter / Esc.

Every option in the GUI corresponds one-to-one with a command-line flag, so a workflow prototyped in the dialog can be translated to a batch file without surprises.

After conversion, a results dialog shows what was converted, skipped, or failed.

---

## Saved configuration (opt-in)

Pass `-u` on the command line, or check **Use configuration** in the GUI, to have 2htm remember the last-used values between runs. Values live in `%LOCALAPPDATA%\2htm\2htm.ini` — a small text file in the standard per-user Windows location.

Without `-u` and without the checkbox, 2htm creates no files of its own. Unchecking the box in the GUI suppresses the write for that run but does not delete any existing configuration file.

**Precedence:** command-line values override saved values, and GUI edits override both.

---

## View output

Pass `--view-output` or check the box in the GUI to open the output directory in File Explorer after conversion. The open fires only if at least one file was actually converted. If an Explorer window is already displaying that directory, 2htm brings it to the foreground instead of spawning a duplicate.

---

## Output

For each input file, a new file is written to the output directory:

- `report.docx` → `report.htm` (or `report.txt` with `-p`)
- Original file is never modified.
- If an output file already exists, the input is skipped unless `-f` is given.

The HTML output is a single standalone file. Images (when kept) are embedded as base64 data URIs, so the `.htm` file can be shared without a sidecar folder. The result passes automated WCAG 2.2 AAA checks in axe-core for landmarks, headings, table structure, alt-text propagation, color contrast, and language declaration.

### Exit codes

| Code | Meaning |
|---|---|
| 0 | All files converted (or help/version shown) |
| 1 | Some files failed |
| 2 | Fatal error (unknown option, unwritable output directory) |

---

## Notes

- Office must be installed and licensed. 2htm drives Word, Excel, and PowerPoint via COM automation; it cannot read these formats without Office.
- 2htm cleans up Office COM servers when it's done, even if conversion fails mid-run. If a workbook or document ever appears hung in Task Manager after a 2htm run, it's a bug — please report it.
- Excel workbooks with very large "used ranges" (hundreds of millions of cells) are handled on a special code path that uses `SpecialCells` instead of reading the full array. This is effective for workbooks that have auto-extended formulas down empty columns.
- PowerPoint automation requires a visible application window (PowerPoint does not support invisible automation). The window is minimized while 2htm works and closed at the end.

---

## Pipelines and integration

Because 2htm is a portable single file with well-defined exit codes, it integrates naturally into automation:

- **Batch scripts** can call `2htm` synchronously and inspect `%ERRORLEVEL%`.
- **Scheduled tasks** can convert folders of documents on a nightly basis to keep an accessible-formats mirror up to date.
- **CI/CD jobs** can turn design docs committed as `.docx` into `.htm` for web publication as part of a build.
- **Content management workflows** can use 2htm as an "alternate format" step, producing accessible HTML versions of public documents alongside the original Office files.

The output file is self-contained, so downstream steps can simply copy or serve the `.htm` file without worrying about dependencies.

---

## Development

This section is for developers who want to build `2htm.exe` from source, or modify the conversion logic.

Because 2htm is released under the MIT license, anyone may use, modify, or redistribute the code for any purpose, including commercial use, as long as the copyright notice is preserved. C# is Microsoft's primary application language, backed by extensive documentation and a mature ecosystem of Office automation examples. The required toolchain (Roslyn, via Visual Studio Community or Visual Studio Build Tools) is free from Microsoft.

### Prerequisites

- Windows 10 or later
- **.NET Framework 4.x** — ships with every supported Windows version (Windows 10 and Windows 11 include it out of the box). No .NET SDK or .NET Core install is needed; the legacy .NET Framework that's already on your machine provides the runtime libraries 2htm links against.
- **A Roslyn C# compiler** — this is the modern compiler that supports current language features. It ships with any of:
  - Visual Studio 2017 or later (any edition, including the free **Community** edition).
  - **Visual Studio Build Tools 2019 or 2022** — a free, smaller download that installs just the compiler (no IDE). During install, select the workload **".NET desktop build tools"**. Download from <https://visualstudio.microsoft.com/downloads/>.

  **Important**: the `csc.exe` bundled with .NET Framework at `%WINDIR%\Microsoft.NET\Framework64\v4.0.30319\csc.exe` is the older pre-Roslyn compiler and cannot build 2htm — it only supports C# 5 and earlier. The build script detects and rejects this compiler; install Roslyn via Visual Studio or the Build Tools.

The build script searches several known Visual Studio install paths automatically.

### Build

Open a Command Prompt in the project folder and run:

```cmd
build2htm.cmd
```

On the first build, the script downloads `Markdig.dll` from nuget.org (the CommonMark parser for Markdown input). Markdig is embedded into `2htm.exe` as a manifest resource, so the resulting executable is a true single file — no sidecar DLLs are needed at runtime.

On subsequent builds, `Markdig.dll` sits next to `2htm.cs` and is reused.

The build targets `x64` to match the 64-bit Office that Microsoft has installed by default since Office 2019. If your Office is 32-bit, edit `build2htm.cmd` and change `/platform:x64` to `/platform:x86` before building. A 64-bit process cannot automate a 32-bit Office COM server and vice versa.

### Source layout

All source lives in a single file: `2htm.cs` (~4,800 lines). The file is organized into several `static class` sections:

- `program` — entry point, argument parsing, conversion dispatch, temp-folder management.
- `logger` — diagnostic logging (opt-in via `-l`).
- `fileIntegrity`, `tempManager` — pre-flight checks and temp file handling.
- `htmlWriter` — WCAG-conformant HTML emitter shared by all converters.
- `comHelper` — COM late-binding helpers for Office automation.
- `wordConverter`, `excelConverter`, `csvConverter`, `pptConverter`, `htmlConverter`, `markdownConverter`, `jsonConverter`, `textConverter`, `textPassthrough` — per-format converters.
- `guiDialog` — WinForms dialog (the `-g` mode).
- `shellHelper` — smart Explorer-window detection for `--view-output`.
- `configManager` — opt-in `.ini` read/write for `-u`.

The code is written in **Camel Type**, a coding style the author developed to convey type information succinctly through variable names themselves, rather than through external type annotations or the reader's memory. Each identifier carries a short prefix signaling its type: `s` for string, `b` for bool, `i` for integer, `n` for real number, `dt` for datetime, `ls` for list, `d` for dictionary, and so on. A reader glancing at `sSource`, `bForce`, or `lsFileArgs` knows the type without cross-referencing a declaration or hovering for IDE tooltips — useful when reading source in a plain text editor, when navigating by screen reader, or when reviewing code in a diff or a printout. Beyond type clarity, Camel Type standardizes capitalization (lower camel case throughout), prefers methods that return meaningful values over void subprocedures, alphabetizes variable declarations at the top of each scope, and requires named constants in place of all magic numbers. The consistency makes the source uniformly easy to scan, audit, and refactor.

One related benefit is readability under a screen reader. Identifiers in lower camel case break naturally into syllables that the screen reader pronounces as distinct words — `sFileName` reads as "s, file, name," not as a single unreadable run. For the same reason, this project uses `readMe.md` rather than the common `README.md` (the all-caps form either gets spelled out letter-by-letter or mis-pronounced as one run-on sound, and also carries an unintended shouting connotation in written English).

Camel Type is language-agnostic; within `2htm.cs` it is applied to C# but the same conventions carry over to any procedural or object-oriented language. For the full C# guidelines see `CamelType_CSharp.md` in this repository; an equivalent document for JavaScript is available as `CamelType_JavaScript.md` in the author's other projects.

### Running from source

There is no "run from source" path — C# requires compilation. After `build2htm.cmd` finishes, run the resulting executable:

```cmd
2htm.exe --help
```

### Development history

This project was developed in collaboration with Anthropic's Claude AI assistant, over a sustained series of design and implementation sessions. The author drove the product decisions — naming, command-line conventions, accessibility priorities, configuration philosophy, GUI layout, and the Camel Type coding style described above — while the AI assisted with C# implementation details, COM automation quirks, Windows API research, and iterative debugging of edge cases (pathological Excel workbooks, PowerPoint shape enumeration, screen-reader keyboard conventions, and so on).

---

## License

MIT License. See `License.htm`.

---

## Download

The whole 2htm project may be downloaded in a single zip archive from:

<http://GitHub.com/JamalMazrui/2htm/archive/main.zip>
