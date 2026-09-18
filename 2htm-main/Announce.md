# 2htm Release Notes

**Copyright:** © 2026 Jamal Mazrui — Released under the [MIT License](https://opensource.org/license/mit/)
**Project home:** <https://github.com/JamalMazrui/2htm>

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

## About the companion accessibility tools

`urlCheck`, `extCheck`, and `2htm` are a small family of free, MIT-licensed Windows command-line tools written by Jamal Mazrui and shared on GitHub. Each is distributed as a single-file, independent binary executable that runs without an installation step, without a runtime dependency, and without anything in the registry.

- [`urlCheck`](https://github.com/JamalMazrui/urlCheck) — drives Microsoft Edge through Playwright and runs axe-core on each page, producing per-page reports plus a session-level Accessibility Conformance Report (`ACR.xlsx` and `ACR.docx`) covering all 86 WCAG 2.2 success criteria.
- [`extCheck`](https://github.com/JamalMazrui/extCheck) — checks the accessibility of `.docx`, `.xlsx`, `.pptx`, and `.md` files, writing per-file CSV reports of issues found by an extensible rule registry.
- [`2htm`](https://github.com/JamalMazrui/2htm) — converts Office documents and other text formats to clean, accessible HTML using Microsoft's own conversion engines, with options for plain text and image stripping.

The three programs share a deliberately consistent interface and a set of friendly features intended to make them equally usable for the typical Windows user (working through a GUI dialog) and for developers automating tasks (working through the command line). Because the same options are available either way, the same scan or conversion can be reproduced from a script or scheduled task exactly as it was performed by hand.

Common features:

- **Fully accessible CLI and GUI**, following platform conventions for accessible interfaces. Every GUI control has a unique mnemonic hotkey; tab order is logical; status, progress, and result messages are announced consistently to screen readers; help text is available in both modes.
- **Equivalent CLI and GUI behavior.** Every option exposed by one mode is exposed by the other, with the same spelling and the same defaults.
- **Familiar across the family.** The three programs use the same control names, dialog layout, and command-line flag spellings wherever the underlying concept is the same. A user who has learned one is immediately at home in the other two — no re-learning.
- **Optional installer for users who prefer a Windows-native install flow.** Each program ships with a small Inno Setup installer (`<program>_setup.exe`) that puts the executable in Program Files, registers a global desktop hotkey (mnemonic to the program name: Alt+Ctrl+U for urlCheck, Alt+Ctrl+2 for 2htm, Alt+Ctrl+X for extCheck), adds a Start Menu entry, and installs the documentation. Users who prefer no installer can run the executable directly from the .zip.
- **Multiple sources in a single command.** A single invocation accepts any number of files, wildcards, folders, URLs, or list files, processed sequentially.
- **Opt-in configuration recall.** When checked, the program remembers the most recent dialog values to a configuration file under `%LOCALAPPDATA%`, so a frequent task is one click away on the next run.
- **Real-time progress in CLI mode; structured summary in GUI mode.** The console shows files and URLs as they are processed, with a short error reason on failure; the GUI message box at the end of the session shows a categorized list of what was completed, failed, or skipped.
- **Force-replacement and skip-existing behavior.** By default, prior outputs are preserved and re-runs skip work already done. A force flag overrides this for a clean re-run.
- **Optional session log.** A diagnostic log can be written next to the program's outputs for after-the-fact review.

All three are released under the [MIT license](https://opensource.org/license/mit/) — a short, permissive open-source license that permits use, modification, and redistribution for any purpose, including commercial use, with the only requirement being that the original copyright notice and license text are preserved in copies of the software.
