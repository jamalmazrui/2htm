# 2htm Release Notes

## Version 1.18.3

This release brings 2htm's source code into compliance with the project's Camel Type coding standard:

- The Hungarian `o` prefix is now reserved for COM objects only. Variables holding managed .NET objects (StreamWriter, FileStream, Process, Regex, Match, ProcessStartInfo, ZipArchive, FileInfo, etc.) now use the lowercase class name as their prefix per the rule.
- Renames are mechanical and do not change runtime behavior. Examples: `oOut` → `writer`, `oFs` → `fileStream`, `oZip` → `zipArchive`, `oMatch` → `match`, `oResult` → `dialogResult`. COM objects driving Word/Excel/PowerPoint (`oWord`, `oExcel`, `oPpt`, `oWb`, `oDoc`, `oCell`, `oRange`, `oSheet`, `oSlide`, `oShape`, `oTable`, `oChart`, etc.) keep their `o` prefix.
- **Cross-program naming.** Identifier names for shared concepts now match across the three companion tools (urlCheck, extCheck, 2htm). The program-name and version constants are `sProgramName` and `sProgramVersion` (was `sVersion`). New constants `sConfigDirName`, `sConfigFileName`, `sLogFileName` replace the inline string literals. GUI layout constants are now all `iLayout*` (was `iDefault*`). The output-directory parameter is `sOutputDir` everywhere (was `sOutDir` in some signatures). The `logger` surface is uniform: `open`, `close`, `info`, `warn`, `error`, `debug`.
- **Picker initial directory.** The Browse source and Choose output buttons now open at the directory derived from the text-field value when that value points to an existing path (whether the user just typed it or it was loaded from a saved configuration), and at the user's Documents folder otherwise.
- **Markdig fetch logic inlined.** The previous `build2htm.cmd` placed the Markdig-download routine in a separate `:fnFetchMarkdig` subroutine and used `call :fnFetchMarkdig` to invoke it. cmd.exe has a known chunk-boundary bug in its label-search code (the search reads the file in 512/1024-byte chunks; a label at certain byte positions can be missed), which surfaced here as "The system cannot find the batch label specified - fnFetchMarkdig". Inlining the fetch logic eliminates the forward `call :label` and so eliminates the bug. .cmd files in this archive ship with CRLF line endings as additional defense-in-depth.

The icon is now embedded in `2htm.exe` at build time via the `/win32icon` flag, and shortcuts inherit it automatically. A 2htm.ico file ships in the GitHub repo for the installer wizard's own icon (compile-time use), but does not need to ship with the installed program.

The installer (`2htm_setup.exe`):

- Prompts for the installation directory (default: `C:\Program Files\2htm`).
- Includes a brief MIT-license summary on the welcome page.
- Installs only HTML versions of the documentation (`ReadMe.htm`, `Announce.htm`, `License.htm`); the Markdown counterparts and the source/build/installer scripts live in the GitHub repository.
- The "Launch 2htm now" checkbox on the final page reminds the user that the desktop hotkey is Alt+Ctrl+2.

## Earlier versions

For the full revision history see the GitHub repository.
