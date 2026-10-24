---
title: "2htm History"
author: "Jamal Mazrui"
---

# 2htm History

## 23 October 2026 -- the Homer launch dialog

- **The dialog is the one every Homer launch app shows,** built by the Homer kit's LaunchDialog (kit 1.89.0). It is in rows called bands, read left to right: Source paths and Browse local source; Plain text and Strip images; Output directory and Choose output directory; Force replacements, Use configuration and View output directory; Default settings and Help; then OK and Cancel alone at the bottom right, as Microsoft's dialog guidance places them.
- **Insert+Tab names the Source paths box again, and the Output directory box is found again.** The kit put a band's label after the band, so the box had no name and Browse seemed to come first.
- **Control+Enter is OK and Escape is Cancel from any control, and F11 checks for a newer version from anywhere.** Help says so, and says where the logs are kept.
- **Logging is automatic.** The Log session checkbox and /log are gone; every run writes %LOCALAPPDATA%\2htm\logs\2htm-yyyyMMdd-HHmmss.log. The Guide button is gone too.
- **Every run ends with the Results box,** whose one OK button has the focus, so Enter, Space or Escape closes it; the output directory then opens if View output directory is checked, and 2htm closes.
- **Switches match the dialog:** a long switch is a control's label and a short one its Alt letter, written with /, - or --, in any case. New: -d, --default-settings; -i for Strip images (-s still works); -v is now View output directory, so the version is /version alone. The older names still work.
- **The progress window is the kit's,** LaunchProgress: "2htm, <status>" in its title, a Status box with the focus, the percentage done and time left ("2 of 5 done, 40%, about 3 min left"), Cancel (Escape) to stop after the current file, and speech only while 2htm is the window in front. Its own window's status label was not reliably spoken, and it had no way out.
- **Force replacements works as in every launch app:** unchecked, a source whose result is already there is skipped and kept, and the Results box says how many and how to replace them; checked, the work is done again and the old result kept until the new one is complete. Its tip is now the kit's standard one.
- **The build needs HomerDev 1.89.0 or later.**

## 20 October 2026 -- Converting through the Homer kit

- **PDFs are read with the Homer kit's own PDF reader,** as in BatCon and bookFido, keeping a tagged PDF's structure and reading an untagged one by its layout. The build fetches PdfPig by itself. It needs HomerDev 1.88.0 or later.
- **The guide describes conversion as it is now:** through the Homer kit, with Office needed only for Word documents without Pandoc and for the older Office formats; Markdig is no longer mentioned.
- **Switches follow the Homer convention:** /force, -force and --force alike, in any case, as every Homer app reads them; the single letters still work, and help shows the /switch form. /view and /output are short for /view-output and /output-dir, as in BatCon.
- **Recordings and speech, as in every Homer app:** a recording converts to a transcript, and a document to spoken audio, through the Homer kit. The build needs HomerDev 1.87.5 or later.
- **2htm converts through the Homer kit,** as every Homer app does, so a file becomes the same web page or text in 2htm as anywhere else. The page form every Homer app now uses came from 2htm: the stylesheet, one heading for the title, the page's parts marked for a screen reader, and pictures carried inside. Word, Excel and PowerPoint files still use Office when Pandoc is not installed.
- **Markdig is no longer part of 2htm;** Markdown becomes a web page through the kit.
- **The build needs HomerDev 1.87.4 or later.**

## 14 October 2026 -- it builds again

- **The build failed on a name out of reach:** the guard against a stuck Office worker ended the run with iExitFatal, which belongs to the program class, not the one the guard is in. It now says program.iExitFatal (from an audit by another AI).

## 14 October 2026 -- a stopped Office worker is proved stopped

- **If Office's worker is still running after Office is ended, the run ends.** The wait for the worker was not checked, so one still running could write after its file was put back. If it has not ended 30 seconds after Office is ended, 2htm logs why and stops, which ends the worker with it (from a second audit by another AI).

## 14 October 2026 -- the developer guide brought up to date

- **The four steps are in the order they run:** build, tidy, push, release, as build_release runs them; the guide listed push before tidy (from an audit by another AI).
- **The developer guide names the kit's classes where they are,** in C:\HomerDev\exec\CSharp; it still named C:\HomerDev\CSharp, a folder the kit no longer has (from an audit by another AI).

## 14 October 2026 -- contents lists, and plainer guides

- **Every help page has a contents list.** The build ran Pandoc without --toc, so a long guide had no list of its sections. Pages are now built with one, and a page that lacks one is built again once to gain it.

## 14 October 2026 -- the announcement in plainer words

- **Announce.md now reads at grade 7.2, down from 13.1,** inside the ninth-grade target for documents shared publicly. Long sentences were split and dense phrases put plainly; every fact, key and link is kept, and first-person passages keep their author's own words. The section about the companion tools reads the same in each announcement that carries it, with the tools' hotkeys in alphabetical order.

## 14 October 2026 -- the newest Pandoc for the help pages

- **The help pages are made by the newest Pandoc, not the first on the PATH.** The build took the first pandoc on the PATH, which was Pandoc 2.19 in C:\bin, while Pandoc 3.12.1 was installed in Program Files. It now asks the kit's newest.ps1, as the kit's app template has done since kit 1.56.0, and the log names every copy found and its version.

## 14 October 2026 -- an Office conversion with a deadline

- **A document Office cannot finish no longer holds 2htm.** An Office call that never returned stopped 2htm with it: no retry, no report. Each attempt now runs on its own thread and has 10 minutes; past that, only the Word, Excel or PowerPoint processes 2htm itself started are ended, so Office windows you had open are left alone, and 2htm says the file was stopped. A timeout is not retried, since a document that hung once would most likely hang again (from an audit by another AI).

## 14 October 2026 -- recovery copies kept, and the program started as you

- **Each conversion keeps its own recovery copy.** The earlier output was moved aside as `<output>.previous`, and an existing one was deleted first, so a second failed conversion erased the first one's recovery copy. Each conversion now gets a dated name of its own, and a failed restore is logged as a warning naming where the copy is (from an audit by another AI).
- **Setup starts 2htm as you,** not with the installer's administrator rights, and creates its logs folder before leaving the launch marker, which a first install never wrote.

## 9 October 2026 -- spoken tutorials

**Ten spoken walks, in the Homer pattern of ten.** Walk 0 is an overview and table of contents; 1, the user interface, each control shown as the screen reader speaks it; 2, installing; 3 to 8, tasks around concrete wants -- a long Word report read in the browser, a PDF made readable, plain text to listen to, slides and spreadsheets, a whole folder, and File Explorer and the command line; and 9, the conclusion, with a glossary in two voices and every way to get help. Each task walk is predicted at three minutes or more by the kit's measured model, and every key and feature it teaches was checked against this guide. The build speaks them; the installer ships them and their audio; the repository carries both.

## 8 October 2026 -- an audit by another AI

ChatGPT audited 2htm and reported 38 findings. Checked against the code, these held and are fixed:

- **A failed reconversion keeps the earlier output.** On any failure the destination was deleted, and with --force that was the previous good output. It is now moved aside first, put back if the conversion fails, and removed only after success.
- **Office documents of your own are safe.** When Office stopped answering, 2htm killed every Word, Excel or PowerPoint without a window, which could include yours. It now records the Office processes it starts, and kills only those.
- **One output, one source.** report.docx and report.md both became report.htm, and the later silently replaced the first; and a conversion could replace another file being converted. Each output now belongs to one source, and the other is skipped and reported.
- **Output goes beside its source, as the guide says.** With no output folder given, output went to the current folder, so a shortcut started in Documents put it there, away from the file converted. The guide's lines that said the current folder are corrected.
- **A missing input is not a success.** A missing file, a pattern matching nothing, or a skipped collision now ends the run with the partial-failure code, not success.
- **A stale copy retired.** A folder named 2htm-main held an old copy of the project, committed by accident; the build removes it.
- **Kit tools** updated from HomerDev 1.63.3.

Left for later, as larger changes: a deadline for a conversion that hangs; the post-install launch that does not drop elevation; and the conversion quality findings -- Excel formulas and headers, PowerPoint order and charts, Word encoding and styles -- each of which needs real Office documents to test.

## Version 1.19.1 (September 2026)

- **Setup.** The Results box at the end of setup is titled "2htm Setup Results", and the finish page uses the Homer wording: the verb first, no "recommended", and "Launch 2htm (desktop hotkey ...)".
- **Built with HomerDev 1.43.19.** The build refreshes the kit's tools under their current names, and the ones that call each other now find each other; `scripts\tidy`, `scripts\check` and `scripts\release` carry the day's fixes, among them a release that publishes a draft and confirms it is GitHub's latest.
- **The dialog is built with the Homer Lbc classes**, like every Homer dialog. The fields, checkboxes and their access keys are unchanged. New with Lbc: Control+Enter is OK from any control, Shift+F1 says a field's tip, F7 lists the controls, the text fields have the Lbc editing keys, and Help lists every field with its tip and ends with the version check. A **Guide** button (Alt+G) opens the full guide and returns to the dialog.

## Version 1.19.0 (September 2026)

### What's new

- **Built on the Homer Development Kit.** 2htm compiles the kit's shared classes (Elevate, Inix, Log, Paths, Say, Web) from `C:\HomerDev\CSharp` and follows the Homer layout: the program in `exec`, the documents in `help`, one log per run in `logs`. Its own dialog and its converters are unchanged.
- **F11 checks for a newer version.** In the dialog, F11 asks GitHub for the latest release. Yes is the default when a newer version exists and No when this one is current; Yes downloads `2htm_setup.exe` and starts it.
- **A session log, always.** Every run writes `%LOCALAPPDATA%\2htm\logs\2htm-yyyyMMdd-HHmmss.log` with the environment, every setting and every file converted; the 30 newest are kept. An error nothing else caught is recorded with its details, and the console names the log. `-l` (Log session) still also writes `2htm.log` beside the output.
- **Settings file moved.** Use configuration now reads and writes the `[Settings]` section of `%LOCALAPPDATA%\2htm\configs\2htm.inix`, keeping any comment a person adds. A `2htm.ini` from an earlier version is read until the new file exists, and removed the first time the new one is saved.
- **Field names spoken once.** The Source files and Output directory fields no longer set an accessible name that repeats their labels. Each label now comes just before its field in tab order, which is how WinForms names the field, so a screen reader says the name once. (Version 1.18.3 had added the repeated names.)
- **Documents.** The ReadMe is a quick start; the full guide is `help\2htm.md` and `.htm`, and Help opens it. New Developer, Hotkeys and this History; Announce describes the current release. The option and format tables are lists.
- **Version from one place.** The version lives in `version.txt`, which the build steps and writes into the program; nothing else carries a number to keep in step.

### Installer

- The program installs to `exec`, the documents other than ReadMe and License to `help`, and `2htm.cmd` at the top runs the program from a command prompt there. The File Explorer entry, Convert via 2htm, now points at `exec\2htm.exe`.
- The finish page offers Launch (checked) and Open the user guide (unchecked). A results box says what was installed and where the logs are, and 2htm starts only after that box is closed.
- A reinstall no longer asks for the folder; it goes where the last one went.
- The uninstaller removes the logs and the File Explorer entry. Your saved settings stay, as before.

### For developers

- `build2htm.cmd` is the kit's C# build template with 2htm's own Markdig section: Markdig 0.18.3 fetched once into `work\nuget` and embedded, with the netstandard facade referenced. Kit version check, `version.txt` seeded and stepped, Roslyn found with vswhere or installed as the Build Tools, the program built into `exec`, the kit's tools refreshed into `scripts`, documents converted, the Homer encoding applied, and every `help` file checked against the installer. It logs every command and exit code to `logs`.
- `RepoFiles.txt` and `LocalFiles.txt` decide what git carries; built programs are no longer in the repository.

## Version 1.18.3

This release brings 2htm's source code into compliance with the project's Camel Type coding standard:

- The Hungarian `o` prefix is now reserved for COM objects only. Variables holding managed .NET objects (StreamWriter, FileStream, Process, Regex, Match, ProcessStartInfo, ZipArchive, FileInfo, etc.) now use the lowercase class name as their prefix per the rule.
- Renames are mechanical and do not change runtime behavior. Examples: `oOut` → `writer`, `oFs` → `fileStream`, `oZip` → `zipArchive`, `oMatch` → `match`, `oResult` → `dialogResult`. COM objects driving Word/Excel/PowerPoint (`oWord`, `oExcel`, `oPpt`, `oWb`, `oDoc`, `oCell`, `oRange`, `oSheet`, `oSlide`, `oShape`, `oTable`, `oChart`, etc.) keep their `o` prefix.
- **Cross-program naming.** Identifier names for shared concepts now match across the three companion tools (urlCheck, extCheck, 2htm). The program-name and version constants are `sProgramName` and `sProgramVersion` (was `sVersion`). New constants `sConfigDirName`, `sConfigFileName`, `sLogFileName` replace the inline string literals. GUI layout constants are now all `iLayout*` (was `iDefault*`). The output-directory parameter is `sOutputDir` everywhere (was `sOutDir` in some signatures). The `logger` surface is uniform: `open`, `close`, `info`, `warn`, `error`, `debug`.
- **Picker initial directory.** The Browse source and Choose output buttons now open at the directory derived from the text-field value when that value points to an existing path (whether the user just typed it or it was loaded from a saved configuration), and at the user's Documents folder otherwise.
- **Friendlier source-field parsing.** When you supply a single path, you no longer need to put quotes around it just because the path contains spaces. 2htm tests the entire trimmed source field as a single spec first; only when it is not a usable single spec does it fall back to space-tokenization. Quotes are only required when supplying multiple specs and at least one contains a space.
- **Source field accessibility.** The Source files and Output directory text boxes now have explicit `AccessibleName` properties so JAWS and NVDA reliably announce each field by its label when focus moves to it, regardless of label-textbox visual layout.
- **Cleaner result messages.** The result MessageBox in GUI mode and the matching console output in CLI mode now show only what the user needs: per-file basenames and the structured summary. The program-name-with-tagline line that used to head the help/usage output is gone (just the version is shown). The `[INFO]` / `[WARN]` / `[ERROR]` prefixes that used to appear on console diagnostic writes have been removed -- they were redundant with the log file's own level columns and made the GUI MessageBox text noisy. The same level data is still recorded in the log file when `-l` is given.
- **Output-directory create prompt.** If you press OK with an output directory that does not yet exist, 2htm asks "Create [path]?" with default Yes. Choosing No keeps the dialog open with focus on the output field.
- **Office automation alerts disabled.** Word, Excel, and PowerPoint application objects created by 2htm now have their `DisplayAlerts` property set to none, plus other prompt-suppression options. Of particular note, `Word.Application.Options.DoNotPromptForConvert = true` suppresses the "Word will now convert your PDF to an editable Word document" dialog that previously locked up 2htm when converting a PDF. `AutomationSecurity = msoAutomationSecurityForceDisable` blocks any macros silently.
- **Progress display fix.** The "Converting" status bar now shows files **already completed** rather than the file being started. When converting a single file, you see "file.pdf — 0 of 1, 0%" while it is being processed (rather than the previous misleading "1 of 1, 100%" while still working).
- **Concise skip message.** The "skipped because output exists" message now uses the input basename rather than the full path, matching the basename-only style of the success line. Full paths still go to the log when `-l` is given.
- **Pre-pruning + structured results summary.** Before the conversion loop runs and before the progress UI opens, the file list is pruned in two passes: (1) unsupported extensions are silently dropped (logged when `-l` is on); (2) files whose output target already exists are dropped unless **Force replacements** is checked — these are counted as "skipped." The progress counter denominator is the post-pruning count, so percentages reflect actual work. The final summary is structured as up to three sections — `Converted N file(s):`, `Failed to convert N file(s):`, and `Skipped N file(s). Check "Force replacements" to overwrite.` — each shown only when its count is non-zero, with singular "file" / plural "files" inflection. Failed entries include a short reason after the basename when one is available (`slides.pptx: file is corrupt`); the full exception and stack trace go to the log when `-l` is on.
- **CLI vs GUI output styles.** In CLI mode (real console attached) basenames print inline as the loop runs — natural progress feedback for the console user. In GUI mode (and right-click invocations) the loop is silent on stdout; the progress status form shows the current file, and the structured summary is the final MessageBox. The structured summary is printed in both modes, but the per-name lists are suppressed in CLI mode (they would just repeat what already scrolled by).
- **Log header.** When `-l` (Log session) is enabled, the log file now begins with a clean header before the timestamped processing notifications: program name and version, a friendly run timestamp (`Run on May 1, 2026 at 2:30 PM`), and a `Parameters:` block listing each setting with both explicit and defaulted values resolved (Source, Output directory, Force replacements, Plain text, Strip images, View output, Use configuration, Log session, GUI mode, Working directory, Command line). The header is followed by the normal timestamped log entries.
- **Markdig fetch logic inlined.** The previous `build2htm.cmd` placed the Markdig-download routine in a separate `:fnFetchMarkdig` subroutine and used `call :fnFetchMarkdig` to invoke it. cmd.exe has a known chunk-boundary bug in its label-search code (the search reads the file in 512/1024-byte chunks; a label at certain byte positions can be missed), which surfaced here as "The system cannot find the batch label specified - fnFetchMarkdig". Inlining the fetch logic eliminates the forward `call :label` and so eliminates the bug. .cmd files in this archive ship with CRLF line endings as additional defense-in-depth.

The icon is now embedded in `2htm.exe` at build time via the `/win32icon` flag, and shortcuts inherit it automatically. A 2htm.ico file ships in the GitHub repo for the installer wizard's own icon (compile-time use), but does not need to ship with the installed program.

The installer (`2htm_setup.exe`):

- Prompts for the installation directory on every run (default: `C:\Program Files\2htm`). The directory page is now explicitly enabled (`DisableDirPage=no` in the .iss); previously it was at the Inno Setup default of `auto`, which silently skipped the page on reinstalls of the same `AppId`. The previous directory is pre-filled, so on a reinstall the user just presses Next to keep the same path.
- Includes a brief MIT-license summary on the welcome page.
- Installs only HTML versions of the documentation (`ReadMe.htm`, `Announce.htm`, `License.htm`); the Markdown counterparts and the source/build/installer scripts live in the GitHub repository.
- The "Launch 2htm now" checkbox on the final page reminds the user that the desktop hotkey is Alt+Ctrl+2.
- Adds a "Convert via **2**htm" entry to the File Explorer right-click menu for all file types. The accelerator letter `2` matches the desktop hotkey accelerator. Uninstall removes the registry entries.

## Earlier versions

For the full revision history see the GitHub repository.
