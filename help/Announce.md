---
title: "2htm Announcement"
author: "Jamal Mazrui"
---

# 2htm 1.19.0

2htm converts Word, Excel, PowerPoint, PDF, Markdown and other files to accessible HTML or plain text. Get it from the [2htm releases page](https://github.com/jamalmazrui/2htm/releases).

## What is new in 1.19.0

- **Update from inside the program.** Press F11 in the dialog to check GitHub for a newer version. If there is one, Enter downloads and starts its installer.
- **A log of every session.** Each run keeps its own log in `%LOCALAPPDATA%\2htm\logs`, so a problem can be explained afterwards even when Log session was off.
- **Settings in the Homer place.** Use configuration now saves to `%LOCALAPPDATA%\2htm\configs\2htm.inix`. Settings from earlier versions are carried over.
- **A tidier installation.** The program is in the `exec` folder and the documents in `help`, as in every Homer Tools program. The installer's results box comes before 2htm starts.

The full list of changes is in History.

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
- **Optional installer for users who prefer a Windows-native install flow.** Each program ships with a small Inno Setup installer (`<program>_setup.exe`) that puts the executable in Program Files, registers a global desktop hotkey (mnemonic to the program name: Alt+Control+2 for 2htm, Alt+Control+X for extCheck, Alt+Control+Shift+U for urlCheck, Alt+Control+U for urlFido), adds a Start Menu entry, and installs the documentation. Users who prefer no installer can run the executable directly from the .zip.
- **Multiple sources in a single command.** A single invocation accepts any number of files, wildcards, folders, URLs, or list files, processed sequentially.
- **Opt-in configuration recall.** When checked, the program remembers the most recent dialog values to a configuration file under `%LOCALAPPDATA%`, so a frequent task is one click away on the next run.
- **Real-time progress in CLI mode; structured summary in GUI mode.** The console shows files and URLs as they are processed, with a short error reason on failure; the GUI message box at the end of the session shows a categorized list of what was completed, failed, or skipped.
- **Force-replacement and skip-existing behavior.** By default, prior outputs are preserved and re-runs skip work already done. A force flag overrides this for a clean re-run.
- **Optional session log.** A diagnostic log can be written next to the program's outputs for after-the-fact review.

All three are released under the [MIT license](https://opensource.org/license/mit/) — a short, permissive open-source license that permits use, modification, and redistribution for any purpose, including commercial use, with the only requirement being that the original copyright notice and license text are preserved in copies of the software.
