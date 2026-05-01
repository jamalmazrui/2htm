// =====================================================================
// 2htm.cs - Convert documents to accessible HTML.
//
// Supported input formats:
//   Office:     .docx .doc .rtf .odt (via Word COM)
//   PDF:        .pdf                  (via Word PDF Reflow, Word 2013+)
//   Workbooks:  .xlsx .xls            (via Excel COM)
//   CSV:        .csv                  (native RFC 4180 parser)
//   Slides:     .pptx .ppt            (via PowerPoint COM)
//   Web:        .html .htm            (native)
//   Markdown:   .md                   (native minimal subset)
//   JSON:       .json                 (native recursive-descent parser)
//   Text:       .txt                  (native)
//
// Build with build2htm.cmd.
//
// Coding style: Camel Type (Hungarian prefix, lowerCamel throughout).
// Keeps COM variables typed as `dynamic` throughout for late-bound
// IDispatch marshalling. Defers casts to narrow try/catch scopes so
// failures are attributable to specific operations.
// =====================================================================

using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Dynamic;
using System.Globalization;
using System.IO;
using System.IO.Compression;
using System.Reflection;
using System.Runtime.InteropServices;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading;
using System.Xml;
using Markdig;

namespace twoHtm
{
    // -----------------------------------------------------------------
    // Program entry point.
    // -----------------------------------------------------------------
    public static class program
    {
        public const int iExitOk = 0;
        public const int iExitPartial = 1;
        public const int iExitFatal = 2;

        public const string sProgramName = "2htm";
        public const string sProgramVersion = "1.18.3";
        public const string sConfigDirName = "2htm";
        public const string sConfigFileName = "2htm.ini";
        public const string sLogFileName = "2htm.log";

        // Global flags set by command-line switches. All converters
        // read these directly rather than having them threaded
        // through every method signature.
        //
        // bStripImages: set by --strip-images / -s. Converters must
        // not emit any <img> tags or image-related markup, and must
        // not leave behind references to images (no broken src
        // attributes, no captions without images). In plain-text
        // mode this flag has no effect because text output has no
        // image concept.
        //
        // bPlainText: set by --plain-text / -p. Output file is .txt
        // instead of .htm, and each converter produces readable
        // unformatted text instead of HTML.
        //
        // bForce: set by --force / -f. Overwrites an existing target
        // file instead of skipping it. Without this flag, a file is
        // only converted when its target does not yet exist.
        //
        // bLog: set by --log / -l, or by the GUI Log session
        // checkbox. Creates 2htm.log in the output directory if one
        // was specified (-o or the GUI Output field), or in the
        // current directory otherwise. Streams detailed diagnostic
        // information to it (in addition to the normal console
        // output). Each session starts with a fresh file -- any
        // prior log is overwritten. Intended to help debug
        // conversion failures; upload the log file when reporting
        // an issue.
        public static bool bStripImages = false;
        public static bool bPlainText = false;
        public static bool bForce = false;
        public static bool bLog = false;
        public static bool bGuiMode = false;

        // When true, the program was launched from File Explorer
        // with arguments (typically via the "Convert with 2htm"
        // right-click menu) and its console window has been
        // hidden. Any error output produced during the run should
        // be surfaced as a MessageBox at the end, since the user
        // has no visible console to see stderr in.
        public static bool bHideConsoleMode = false;
        // When true, open the output directory in the default
        // file manager after conversion (--view-output). The open
        // fires only if at least one file was actually converted,
        // and reuses an already-open Explorer window on that
        // directory rather than launching a duplicate.
        public static bool bViewOutput = false;

        // When true, load saved defaults from
        // %LOCALAPPDATA%\2htm\2htm.ini on startup (before the GUI
        // dialog opens) and, on OK-click, write the dialog's
        // current values back. Off by default — 2htm leaves no
        // filesystem footprint of its own unless the user opts in
        // via -u / --use-configuration or the GUI checkbox. The
        // config is never auto-deleted; unchecking the box at OK
        // time simply suppresses the write for this run.
        public static bool bUseConfig = false;

        // Parallel "was this set on the command line" flags. When
        // true, the saved config file must NOT overwrite the
        // command-line-supplied value during the pre-GUI load.
        // Command line always wins over saved config. GUI edits
        // (which happen after config load) always win over both.
        public static bool bStripImagesFromCli = false;
        public static bool bPlainTextFromCli = false;
        public static bool bForceFromCli = false;
        public static bool bLogFromCli = false;
        public static bool bViewOutputFromCli = false;
        public static bool bOutputDirFromCli = false;
        public static bool bSourceFromCli = false;

        // Output directory. When empty, outputs go to the current
        // working directory (the historical behavior). Set by -o or
        // --output-dir. The path is stored verbatim; expansion to
        // full path is deferred to convertOne. An absent directory
        // is treated as fatal at run-time.
        public static string sOutputDir = "";

        // [STAThread] is REQUIRED for Office COM automation.
        // Without it, the process runs in the MTA apartment. Word
        // happens to tolerate MTA for basic operations, but Excel
        // (especially UsedRange.Value2) and PowerPoint (shape
        // iteration, Slide.Export) will disconnect mid-operation
        // with HRESULT 0x80010108 (RPC_E_DISCONNECTED) or
        // 0x80010114 (OLE_E_OBJNOTCONNECTED). Setting [STAThread]
        // on Main initializes the main thread as STA, which is the
        // apartment Office COM servers expect.
        //
        // IMPORTANT: Main must not reference any type from an embedded
        // assembly (e.g., Markdig). The CLR JITs a method just before
        // executing it, and any type references in that method trigger
        // assembly resolution. We register the AssemblyResolve handler
        // FIRST, then delegate the real work to run() so that any
        // markdown processing (which pulls in Markdig types) is
        // compiled after our handler is in place.
        [STAThread]
        public static int Main(string[] asArgs)
        {
            AppDomain.CurrentDomain.AssemblyResolve += resolveEmbeddedAssembly;
            return run(asArgs);
        }

        // Returns true when this process appears to have been
        // launched by a GUI shell (File Explorer double-click,
        // Start-menu shortcut, pinned taskbar button, zip-archive
        // extraction-and-run) rather than from a command-line
        // shell (cmd.exe, PowerShell, Windows Terminal).
        //
        // Technique: GetConsoleProcessList returns the process IDs
        // attached to the current console. When a shell launches a
        // console-subsystem executable, the new process inherits
        // the shell's console, giving a count >= 2 (shell plus us,
        // at minimum). When the shell creates a fresh console for
        // a double-clicked executable, we're the only process on
        // that console, so the count is 1.
        //
        // A count of 0 means no console is attached at all (service
        // context, some scheduled-task configurations). In that
        // case we fall through to command-line behavior, which
        // simply prints usage and exits — the right outcome for a
        // non-interactive invocation that didn't supply arguments.
        [System.Runtime.InteropServices.DllImport("kernel32.dll", SetLastError = true)]
        private static extern uint GetConsoleProcessList(
            [System.Runtime.InteropServices.Out] uint[] aiProcessIds,
            uint iCount);

        [System.Runtime.InteropServices.DllImport("kernel32.dll")]
        private static extern IntPtr GetConsoleWindow();

        [System.Runtime.InteropServices.DllImport("user32.dll")]
        private static extern bool ShowWindow(IntPtr hWnd, int nCmdShow);

        private const int iSwHide = 0;

        private static bool isLaunchedFromGui()
        {
            try {
                var aiList = new uint[16];
                uint iCount = GetConsoleProcessList(aiList, (uint)aiList.Length);
                return iCount == 1;
            } catch {
                // If the API is unavailable for any reason, prefer
                // the safer behavior: assume command-line mode.
                return false;
            }
        }

        // Hides this process's console window. Only safe to call
        // when we own our console (i.e., isLaunchedFromGui()
        // returned true). In the GUI-launch case the console was
        // created by Windows specifically for us; hiding it removes
        // an extra window that would otherwise confuse the user —
        // they see only the 2htm dialog, no mysterious command
        // prompt behind it, no extra taskbar entry, no stray Alt+Tab
        // target.
        //
        // IMPORTANT: never call this when the console was inherited
        // from a parent shell (cmd.exe, PowerShell). Hiding an
        // inherited console would remove the user's shell window
        // from the screen, which is not our call to make. That is
        // why this runs ONLY on the auto-detected GUI-launch path,
        // NOT when -g was supplied from a command line.
        private static void hideOwnConsoleWindow()
        {
            try {
                IntPtr hwnd = GetConsoleWindow();
                if (hwnd != IntPtr.Zero) ShowWindow(hwnd, iSwHide);
            } catch {
                // Non-fatal: if we can't hide the console, the GUI
                // still works, the console just stays visible.
            }
        }

        private static int run(string[] asArgs)
        {
            tempManager.sweepStale();

            // If the program was launched with no arguments, either
            // show command-line usage (when invoked from an actual
            // shell) or fall through into GUI mode (when invoked by
            // double-clicking in File Explorer, launching from a
            // Start-menu shortcut, or extracting a copy from a zip
            // and running it). The GetConsoleProcessList API
            // distinguishes the two cases cleanly: a process
            // launched from cmd.exe or PowerShell shares its parent
            // shell's console (count >= 2), whereas a process
            // launched by the shell creating a new console for it
            // (double-click, context-menu verb) is the only process
            // attached to that console (count == 1). See Raymond
            // Chen's write-up of this technique for context.
            //
            // The console-hiding logic applies whenever we own our
            // console, regardless of whether arguments are present.
            // That handles TWO user-visible cases:
            //   (a) Double-click from Explorer with no arguments:
            //       auto-GUI mode triggers, dialog appears, no
            //       mysterious blank console window.
            //   (b) Right-click "Convert with 2htm" on a file in
            //       Explorer: arguments ARE present (the filename),
            //       so CLI mode runs, but we still hide the
            //       console flash that would otherwise occur.
            //       Any error output is captured and shown as a
            //       MessageBox at the end of the run.
            bool bOwnConsole = isLaunchedFromGui();
            if (asArgs.Length == 0) {
                if (bOwnConsole) {
                    bGuiMode = true;
                    hideOwnConsoleWindow();
                } else {
                    printUsage();
                    return iExitOk;
                }
            } else if (bOwnConsole) {
                // CLI launched by Explorer (context-menu verb).
                // Hide the console flash and mark that we should
                // report any conversion errors via MessageBox.
                bHideConsoleMode = true;
                hideOwnConsoleWindow();
            }

            // First pass: recognize flags and separate them from
            // file/wildcard arguments. Flags starting with "-" are
            // consumed here; any "-"-argument not recognized as a
            // flag is an error (we do not silently fall through to
            // wildcard expansion for those, since a typo in a flag
            // could otherwise be interpreted as a pattern).
            //
            // Indexed loop (not foreach) because -o/--output-dir
            // consumes the following argument as its value.
            var lsFileArgs = new List<string>();
            for (int i = 0; i < asArgs.Length; i++) {
                string sArg = asArgs[i];
                if (sArg == "-h" || sArg == "--help" || sArg == "/?" || sArg == "-?") {
                    printUsage();
                    return iExitOk;
                }
                if (sArg == "-v" || sArg == "--version") {
                    Console.WriteLine(program.sProgramName + " " + sProgramVersion);
                    return iExitOk;
                }
                if (sArg == "-s" || sArg == "--strip-images") {
                    bStripImages = true; bStripImagesFromCli = true;
                    continue;
                }
                if (sArg == "-p" || sArg == "--plain-text") {
                    bPlainText = true; bPlainTextFromCli = true;
                    continue;
                }
                if (sArg == "-f" || sArg == "--force") {
                    bForce = true; bForceFromCli = true;
                    continue;
                }
                if (sArg == "-l" || sArg == "--log") {
                    bLog = true; bLogFromCli = true;
                    continue;
                }
                if (sArg == "-g" || sArg == "--gui-mode") {
                    bGuiMode = true;
                    continue;
                }
                if (sArg == "--view-output") {
                    bViewOutput = true; bViewOutputFromCli = true;
                    continue;
                }
                if (sArg == "-u" || sArg == "--use-configuration") {
                    bUseConfig = true;
                    continue;
                }
                if (sArg == "-o" || sArg == "--output-dir") {
                    // Expects a following argument giving the
                    // directory. If missing or empty, treat as fatal
                    // so the user learns the correct usage.
                    if (i + 1 >= asArgs.Length ||
                        string.IsNullOrWhiteSpace(asArgs[i + 1])) {
                        Console.Error.WriteLine("[ERROR] " + sArg +
                            " requires a directory argument.");
                        Console.Error.WriteLine("Run '2htm --help' for usage.");
                        return iExitFatal;
                    }
                    sOutputDir = asArgs[++i];
                    bOutputDirFromCli = true;
                    continue;
                }
                if (sArg.StartsWith("-") && !File.Exists(sArg) && !Directory.Exists(sArg)) {
                    Console.Error.WriteLine("[ERROR] Unknown option: " + sArg);
                    Console.Error.WriteLine("Run '2htm --help' for usage.");
                    return iExitFatal;
                }
                lsFileArgs.Add(sArg);
                bSourceFromCli = true;
            }

            // Configuration file (opt-in, per-user INI at
            // %LOCALAPPDATA%\2htm\2htm.ini):
            //
            //   - In CLI mode: the file is loaded ONLY when -u /
            //     --use-configuration is passed. A user who never
            //     passes -u and has never created a config file
            //     has zero filesystem footprint from 2htm beyond
            //     the conversions they explicitly request.
            //
            //   - In GUI mode: if the file exists, it is loaded
            //     automatically and bUseConfig is treated as
            //     implicitly true. The "Use configuration"
            //     checkbox in the dialog will show as checked,
            //     reflecting that a saved config is in effect.
            //     The user can uncheck the box to stop refreshing
            //     the saved values on future OK-clicks; the file
            //     stays on disk until the user deletes it
            //     manually.
            //
            // This asymmetry keeps the CLI zero-footprint promise
            // while giving GUI users a natural "remember my
            // settings" experience.
            if (bGuiMode && !bUseConfig && configManager.configExists()) {
                bUseConfig = true;
            }
            if (bUseConfig) {
                configManager.loadInto(lsFileArgs);
            }

            // In GUI mode, invoke the dialog before any file
            // processing. The dialog uses the already-parsed values
            // (file args, sOutputDir, bStripImages, bPlainText,
            // bForce) as defaults; its OK-path copies the user's
            // choices back into the same globals. Cancel-path exits
            // without converting anything.
            if (bGuiMode) {
                string sSource = string.Join(" ",
                    lsFileArgs.ConvertAll(s => s.Contains(" ") ? "\"" + s + "\"" : s));
                string sOut = sOutputDir;
                bool bStrip = bStripImages;
                bool bPlain = bPlainText;
                bool bForceLocal = bForce;
                bool bView = bViewOutput;
                bool bLogLocal = bLog;
                bool bUseCfg = bUseConfig;
                if (!guiDialog.show(ref sSource, ref sOut, ref bStrip,
                    ref bPlain, ref bForceLocal, ref bView, ref bLogLocal, ref bUseCfg)) {
                    return iExitOk;
                }
                bStripImages = bStrip;
                bPlainText = bPlain;
                bForce = bForceLocal;
                bViewOutput = bView;
                bLog = bLogLocal;
                bUseConfig = bUseCfg;
                sOutputDir = sOut;
                // Re-parse the source field. It may contain multiple
                // paths or patterns separated by whitespace, and
                // quoted paths for names with spaces.
                lsFileArgs = splitSourceField(sSource);

                // If the user left the Use configuration box
                // checked, save the dialog's current values as the
                // next-run defaults. If unchecked, do nothing;
                // never auto-delete (the user's filesystem, their
                // call).
                if (bUseConfig) {
                    configManager.save(sSource, sOutputDir, bStripImages,
                        bPlainText, bForce, bViewOutput, bLog);
                }
            } else if (lsFileArgs.Count == 0) {
                printUsage();
                return iExitOk;
            }

            // Validate the output directory. Create it if it doesn't
            // exist; fail fast if that fails.
            if (!string.IsNullOrEmpty(sOutputDir)) {
                try {
                    if (!Directory.Exists(sOutputDir))
                        Directory.CreateDirectory(sOutputDir);
                } catch (Exception ex) {
                    string sErr = "Output directory '" + sOutputDir +
                        "' does not exist and cannot be created: " + ex.Message;
                    Console.Error.WriteLine("[ERROR] " + sErr);
                    if (bGuiMode) showFinalMessage(sErr);
                    return iExitFatal;
                }
            }

            if (bLog) logger.open(sOutputDir);

            // In GUI mode, capture all console output so we can
            // show it to the user at the end via MessageBox. In
            // hide-console mode (CLI launched from Explorer's
            // context menu, console hidden), capture stderr only —
            // stdout's basename-progress noise is not worth
            // interrupting the user with on a successful run, but
            // errors must still be surfaced somehow since the user
            // has no visible console to read them from.
            TextWriter writerOriginalOut = Console.Out;
            TextWriter writerOriginalErr = Console.Error;
            StringWriter stringWriterOut = null;
            StringWriter stringWriterErr = null;
            if (bGuiMode) {
                stringWriterOut = new StringWriter();
                Console.SetOut(stringWriterOut);
                Console.SetError(stringWriterOut);
            } else if (bHideConsoleMode) {
                stringWriterErr = new StringWriter();
                Console.SetError(stringWriterErr);
            }

            int iExitCode = iExitOk;
            try {
                // Write the run header to the log: program version,
                // friendly start time, and the resolved parameter
                // list (showing both explicit and defaulted values).
                // The parameter list mirrors the GUI dialog controls
                // so the user can map a logged run to the equivalent
                // dialog state.
                var lsParams = new List<KeyValuePair<string, string>>();
                lsParams.Add(new KeyValuePair<string, string>("Source",
                    lsFileArgs.Count == 0
                        ? "(none)"
                        : string.Join(" ", lsFileArgs.ConvertAll(s => s.Contains(" ") ? "\"" + s + "\"" : s))));
                lsParams.Add(new KeyValuePair<string, string>("Output directory",
                    string.IsNullOrEmpty(sOutputDir)
                        ? Directory.GetCurrentDirectory() + " (default: working dir)"
                        : sOutputDir));
                lsParams.Add(new KeyValuePair<string, string>("Force replacements", bForce.ToString().ToLowerInvariant()));
                lsParams.Add(new KeyValuePair<string, string>("Plain text",         bPlainText.ToString().ToLowerInvariant()));
                lsParams.Add(new KeyValuePair<string, string>("Strip images",       bStripImages.ToString().ToLowerInvariant()));
                lsParams.Add(new KeyValuePair<string, string>("View output",        bViewOutput.ToString().ToLowerInvariant()));
                lsParams.Add(new KeyValuePair<string, string>("Use configuration",  bUseConfig.ToString().ToLowerInvariant()));
                lsParams.Add(new KeyValuePair<string, string>("Log session",        bLog.ToString().ToLowerInvariant()));
                lsParams.Add(new KeyValuePair<string, string>("GUI mode",           bGuiMode.ToString().ToLowerInvariant()));
                lsParams.Add(new KeyValuePair<string, string>("Working directory",  Directory.GetCurrentDirectory()));
                lsParams.Add(new KeyValuePair<string, string>("Command line",       string.Join(" ", asArgs)));
                logger.header(program.sProgramName, sProgramVersion, lsParams);

                var lsFiles = expandWildcards(lsFileArgs.ToArray());
                if (lsFiles.Count == 0) {
                    Console.Error.WriteLine("[INFO] No matching files.");
                    logger.info("No matching files; exiting.");
                    iExitCode = iExitOk;
                } else {
                    logger.info("Matched " + lsFiles.Count + " file(s).");

                    // ---- Pre-prune ----
                    //
                    // Two passes happen BEFORE the conversion loop
                    // and BEFORE the progress UI opens, so the
                    // counter denominator reflects only files that
                    // will actually be processed:
                    //
                    //   1. Drop files whose extension 2htm cannot
                    //      handle (silent; logged).
                    //   2. If --force is NOT set, drop files whose
                    //      target output already exists. These are
                    //      counted as "skipped" and surface in the
                    //      results summary with a Force-replacements
                    //      hint.
                    List<string> lsToConvert;
                    List<string> lsSkippedExisting;
                    prePrune(lsFiles, out lsToConvert, out lsSkippedExisting);

                    // ---- Conversion loop ----
                    var lsConverted = new List<string>();
                    var lsFailed = new List<program.failure>();
                    if (lsToConvert.Count > 0) {
                        // GUI mode (or right-click) shows a status
                        // form during the loop and the structured
                        // summary at the end as a MessageBox; no
                        // inline per-file console writes -- the
                        // captured stdout becomes the MessageBox
                        // text, and we want only the summary there.
                        //
                        // Pure CLI mode (a real console attached)
                        // prints each basename inline as work
                        // begins, so the user sees progress as the
                        // run proceeds; on success the line is
                        // terminated with a newline, on failure
                        // ": <reason>" is appended on the same line.
                        // The structured summary is also printed at
                        // the end of the run.
                        bool bGuiOrHidden = bGuiMode || bHideConsoleMode;
                        if (bGuiOrHidden) {
                            guiProgress.runConversions(lsToConvert,
                                out lsConverted, out lsFailed);
                        } else {
                            runConversionLoop(lsToConvert, null,
                                /* bInlineConsole: */ true,
                                out lsConverted, out lsFailed);
                        }
                    }

                    // ---- Structured results summary ----
                    //
                    // Three sections, each printed only when its
                    // count is non-zero. Singular "file" when the
                    // count is 1; plural "files" otherwise. In GUI
                    // mode the captured stdout becomes the final
                    // MessageBox; in CLI mode it follows the inline
                    // basenames printed during the loop.
                    int iConverted = lsConverted.Count;
                    int iFailed = lsFailed.Count;
                    int iSkippedExisting = lsSkippedExisting.Count;
                    bool bGuiOrHidden2 = bGuiMode || bHideConsoleMode;

                    if (iConverted > 0) {
                        // In CLI mode the basenames were already
                        // printed inline during the loop. Print the
                        // section header but suppress the per-name
                        // list (it would just repeat what scrolled
                        // by). In GUI/hidden mode include the names
                        // since nothing else has been written.
                        Console.WriteLine("Converted " + iConverted + " " +
                            (iConverted == 1 ? "file" : "files") + ":");
                        if (bGuiOrHidden2) {
                            foreach (var sName in lsConverted)
                                Console.WriteLine(sName);
                        }
                    }
                    if (iFailed > 0) {
                        if (iConverted > 0) Console.WriteLine();
                        Console.WriteLine("Failed to convert " + iFailed + " " +
                            (iFailed == 1 ? "file" : "files") + ":");
                        if (bGuiOrHidden2) {
                            // In GUI/hidden mode list each failure on
                            // its own line as "basename: reason" (or
                            // just basename if no reason).
                            foreach (var oFail in lsFailed) {
                                if (string.IsNullOrEmpty(oFail.sReason))
                                    Console.WriteLine(oFail.sBase);
                                else
                                    Console.WriteLine(oFail.sBase + ": " + oFail.sReason);
                            }
                        }
                        // In CLI mode the failure line was already
                        // written inline. Don't repeat.
                    }
                    if (iSkippedExisting > 0) {
                        if (iConverted > 0 || iFailed > 0) Console.WriteLine();
                        Console.WriteLine("Skipped " + iSkippedExisting + " " +
                            (iSkippedExisting == 1 ? "file" : "files") +
                            ". Check \"Force replacements\" to overwrite.");
                    }
                    if (iConverted == 0 && iFailed == 0 && iSkippedExisting == 0) {
                        Console.WriteLine("No supported files to convert.");
                    }

                    logger.info("Done. " + iConverted + " converted, " +
                        iSkippedExisting + " skipped, " + iFailed + " failed.");
                    iExitCode = iFailed == 0 ? iExitOk : iExitPartial;

                    // --view-output: open the output directory in
                    // the shell, but only if at least one file was
                    // actually converted. Skipped-only runs should
                    // not pop an Explorer window.
                    if (bViewOutput && iConverted > 0) {
                        string sDirToOpen = string.IsNullOrEmpty(sOutputDir)
                            ? Directory.GetCurrentDirectory()
                            : sOutputDir;
                        shellHelper.openFolderSmart(sDirToOpen);
                    }
                }
            } finally {
                logger.close();
                if (bGuiMode) {
                    Console.SetOut(writerOriginalOut);
                    Console.SetError(writerOriginalErr);
                    string sCaptured = stringWriterOut != null ? stringWriterOut.ToString() : "";
                    showFinalMessage(string.IsNullOrWhiteSpace(sCaptured)
                        ? "Done. No output."
                        : sCaptured);
                } else if (bHideConsoleMode) {
                    Console.SetError(writerOriginalErr);
                    string sErrText = stringWriterErr != null ? stringWriterErr.ToString() : "";
                    // Only interrupt the user if something actually
                    // went wrong. Silent success is the right UX
                    // for a right-click-to-convert action: the
                    // output file appears, the console flash never
                    // happened, done. On error, surface the
                    // stderr text so the user knows why no output
                    // appeared.
                    if (!string.IsNullOrWhiteSpace(sErrText) &&
                        iExitCode != iExitOk) {
                        try {
                            System.Windows.Forms.MessageBox.Show(sErrText,
                                "2htm — Conversion error",
                                System.Windows.Forms.MessageBoxButtons.OK,
                                System.Windows.Forms.MessageBoxIcon.Warning);
                        } catch { }
                    }
                    // Successful silent conversion from a right-
                    // click invocation: the output file appears
                    // in the folder, no dialog, no console. The
                    // user confirms completion by finding the
                    // new file in the same folder as the input.
                }
            }
            return iExitCode;
        }

        // Splits the GUI's "Source files" textbox contents back into
        // individual arguments. Handles simple space separation and
        // double-quoted paths with embedded spaces.
        public static List<string> splitSourceField(string sField)
        {
            // Friendlier parsing rules:
            //   1. Trim the input.
            //   2. Strip a single layer of surrounding double quotes.
            //   3. Test the entire (unquoted) trimmed field as a single
            //      spec (existing file, existing directory, or wildcard
            //      pattern matching at least one file). If usable,
            //      return it as one token.
            //   4. Otherwise fall back to space-tokenization, honoring
            //      "..." segments so the user can mix paths-with-spaces
            //      and ones without.
            // The user only needs to use quotes when supplying multiple
            // specs and at least one contains a space.
            var ls = new List<string>();
            if (string.IsNullOrWhiteSpace(sField)) return ls;
            string sTrimmed = sField.Trim();
            string sUnquoted = sTrimmed;
            if (sUnquoted.Length >= 2 && sUnquoted[0] == '"' && sUnquoted[sUnquoted.Length - 1] == '"')
                sUnquoted = sUnquoted.Substring(1, sUnquoted.Length - 2).Trim();

            if (isUsableSingleSpec(sUnquoted)) {
                ls.Add(sUnquoted);
                return ls;
            }

            var sb = new StringBuilder();
            bool bInQuote = false;
            foreach (char c in sTrimmed) {
                if (c == '"') { bInQuote = !bInQuote; continue; }
                if (!bInQuote && (c == ' ' || c == '\t')) {
                    if (sb.Length > 0) { ls.Add(sb.ToString()); sb.Clear(); }
                    continue;
                }
                sb.Append(c);
            }
            if (sb.Length > 0) ls.Add(sb.ToString());
            return ls;
        }

        /// <summary>
        /// Return true if sSpec, taken whole, is a usable file
        /// specification (existing file, existing directory, or
        /// wildcard pattern that matches at least one file). Used by
        /// splitSourceField to recognize a single path with embedded
        /// spaces vs multiple space-separated paths.
        /// </summary>
        private static bool isUsableSingleSpec(string sSpec)
        {
            if (string.IsNullOrEmpty(sSpec)) return false;
            try {
                if (System.IO.File.Exists(sSpec)) return true;
                if (System.IO.Directory.Exists(sSpec)) return true;
                if (sSpec.IndexOfAny(new[] { '*', '?' }) >= 0) {
                    string sDir = System.IO.Path.GetDirectoryName(sSpec);
                    if (string.IsNullOrEmpty(sDir))
                        sDir = System.IO.Directory.GetCurrentDirectory();
                    string sPattern = System.IO.Path.GetFileName(sSpec);
                    if (System.IO.Directory.Exists(sDir) && !string.IsNullOrEmpty(sPattern)) {
                        try {
                            string[] aMatched = System.IO.Directory.GetFiles(sDir, sPattern);
                            if (aMatched != null && aMatched.Length > 0) return true;
                        } catch { }
                    }
                }
            } catch { }
            return false;
        }

        // Displays a modal MessageBox with the captured console
        // output. Large output is scrollable via a multi-line
        // TextBox inside a small Form; small output uses the
        // native MessageBox.
        private static void showFinalMessage(string sText)
        {
            // MessageBox handles ~1000 chars gracefully; longer than
            // that and it truncates. Switch to a custom scrollable
            // form for anything bigger than 800 chars or more than
            // 15 lines.
            bool bLong = sText.Length > 800 ||
                sText.Split('\n').Length > 15;
            if (!bLong) {
                System.Windows.Forms.MessageBox.Show(sText,
                    "2htm — Results",
                    System.Windows.Forms.MessageBoxButtons.OK,
                    System.Windows.Forms.MessageBoxIcon.Information);
                return;
            }

            var frm = new System.Windows.Forms.Form();
            frm.Text = "2htm — Results";
            frm.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen;
            frm.ClientSize = new System.Drawing.Size(600, 400);
            frm.FormBorderStyle = System.Windows.Forms.FormBorderStyle.Sizable;
            frm.MinimizeBox = false;
            frm.MaximizeBox = true;
            frm.ShowInTaskbar = false;
            frm.Font = System.Drawing.SystemFonts.MessageBoxFont;

            var txt = new System.Windows.Forms.TextBox();
            txt.Multiline = true;
            txt.ReadOnly = true;
            txt.ScrollBars = System.Windows.Forms.ScrollBars.Vertical;
            txt.Text = sText;
            txt.Dock = System.Windows.Forms.DockStyle.Fill;
            txt.Font = new System.Drawing.Font(System.Drawing.FontFamily.GenericMonospace,
                9.0f);
            frm.Controls.Add(txt);

            var pnl = new System.Windows.Forms.Panel();
            pnl.Height = 40;
            pnl.Dock = System.Windows.Forms.DockStyle.Bottom;
            frm.Controls.Add(pnl);

            var btn = new System.Windows.Forms.Button();
            btn.Text = "OK";
            btn.DialogResult = System.Windows.Forms.DialogResult.OK;
            btn.Size = new System.Drawing.Size(100, 26);
            btn.Anchor = System.Windows.Forms.AnchorStyles.Top |
                System.Windows.Forms.AnchorStyles.Right;
            btn.Location = new System.Drawing.Point(
                pnl.ClientSize.Width - btn.Width - 12, 7);
            pnl.Controls.Add(btn);
            frm.AcceptButton = btn;
            frm.CancelButton = btn;

            frm.ShowDialog();
        }

        // Loads assemblies that have been embedded as manifest
        // resources into this EXE. The csc /resource: switch at build
        // time embeds Markdig.dll with the logical name "Markdig.dll".
        // At runtime the CLR raises AssemblyResolve when it cannot
        // find a referenced assembly on disk; this handler reads the
        // bytes from the embedded resource and loads it in-memory.
        private static Assembly resolveEmbeddedAssembly(object sender, ResolveEventArgs args)
        {
            try {
                var assemblyName = new AssemblyName(args.Name);
                string sResource = assemblyName.Name + ".dll";
                Assembly assembly = Assembly.GetExecutingAssembly();
                using (Stream stream = assembly.GetManifestResourceStream(sResource)) {
                    if (stream == null) return null;
                    var binBytes = new byte[stream.Length];
                    int iRead = 0;
                    while (iRead < binBytes.Length) {
                        int iJust = stream.Read(binBytes, iRead, binBytes.Length - iRead);
                        if (iJust <= 0) break;
                        iRead += iJust;
                    }
                    return Assembly.Load(binBytes);
                }
            } catch {
                return null;
            }
        }

        private static void printUsage()
        {
            Console.WriteLine(program.sProgramName + " " + sProgramVersion + " - convert documents to accessible HTML");
            Console.WriteLine();
            Console.WriteLine("Usage: 2htm [options] <file-or-wildcard> [<file-or-wildcard> ...]");
            Console.WriteLine();
            Console.WriteLine("  One or more file arguments may be given. Each may be:");
            Console.WriteLine("    - a literal filename");
            Console.WriteLine("    - a wildcard pattern such as *.xlsx");
            Console.WriteLine("    - a folder path (equivalent to <folder>\\*.* with");
            Console.WriteLine("      unsupported extensions filtered out)");
            Console.WriteLine("  Files matching more than one argument are processed only once.");
            Console.WriteLine();
            Console.WriteLine("Examples:");
            Console.WriteLine("  2htm report.docx");
            Console.WriteLine("  2htm *.xlsx *.pptx");
            Console.WriteLine("  2htm --plain-text C:\\docs\\*.pdf reports\\*.docx");
            Console.WriteLine("  2htm --force --strip-images slides\\*.pptx");
            Console.WriteLine();
            Console.WriteLine("Options:");
            Console.WriteLine("  -h, --help, -?, /?   Show this help and exit.");
            Console.WriteLine("  -v, --version        Show the version number and exit.");
            Console.WriteLine("  -f, --force          Overwrite the target file if it already");
            Console.WriteLine("                       exists. Default is to skip.");
            Console.WriteLine("  -s, --strip-images   Strip images from the output. No <img>");
            Console.WriteLine("                       tags, no broken references, no alt text.");
            Console.WriteLine("                       Use for large image-heavy documents.");
            Console.WriteLine("  -p, --plain-text     Produce .txt instead of .htm. Output is");
            Console.WriteLine("                       readable plain text (no HTML markup).");
            Console.WriteLine("  -o <dir>, --output-dir <dir>");
            Console.WriteLine("                       Write output files to <dir> instead of");
            Console.WriteLine("                       the current directory. <dir> is created");
            Console.WriteLine("                       if it does not exist.");
            Console.WriteLine("  -g, --gui-mode       Launch a dialog to enter parameters. Any");
            Console.WriteLine("                       options given on the command line become");
            Console.WriteLine("                       the defaults. After conversion, a message");
            Console.WriteLine("                       box displays the per-file results. GUI mode");
            Console.WriteLine("                       is also entered automatically when 2htm is");
            Console.WriteLine("                       launched with no arguments from a GUI shell");
            Console.WriteLine("                       such as File Explorer (e.g., by double-");
            Console.WriteLine("                       clicking the executable or extracting it");
            Console.WriteLine("                       from a zip archive and running it).");
            Console.WriteLine("  --view-output        After conversion, open the output");
            Console.WriteLine("                       directory in the default file manager");
            Console.WriteLine("                       (typically File Explorer). Fires only if");
            Console.WriteLine("                       at least one file was actually converted;");
            Console.WriteLine("                       reuses an already-open window if one is");
            Console.WriteLine("                       already displaying that directory.");
            Console.WriteLine("  -u, --use-configuration");
            Console.WriteLine("                       Read saved defaults from");
            Console.WriteLine("                       %LOCALAPPDATA%\\2htm\\2htm.ini. Any other");
            Console.WriteLine("                       options supplied on the command line take");
            Console.WriteLine("                       precedence over saved values. In GUI mode,");
            Console.WriteLine("                       the Use configuration checkbox also");
            Console.WriteLine("                       controls whether the current dialog values");
            Console.WriteLine("                       are written back as next-run defaults.");
            Console.WriteLine("                       Without -u (and no checkbox), 2htm");
            Console.WriteLine("                       creates no files of its own.");
            Console.WriteLine("  -l, --log            Write detailed diagnostics to 2htm.log");
            Console.WriteLine("                       (UTF-8 with BOM) in the output directory if");
            Console.WriteLine("                       one was specified, or the current working");
            Console.WriteLine("                       directory otherwise. Any prior 2htm.log in");
            Console.WriteLine("                       that location is overwritten so the file");
            Console.WriteLine("                       always reflects only the current session.");
            Console.WriteLine("                       Useful for reporting conversion problems.");
            Console.WriteLine();
            Console.WriteLine("Supported input formats:");
            Console.WriteLine("  Word:       .docx .doc .rtf .odt");
            Console.WriteLine("  PDF:        .pdf   (via Word 2013+ PDF Reflow)");
            Console.WriteLine("  Excel:      .xlsx .xls");
            Console.WriteLine("  CSV:        .csv   (native)");
            Console.WriteLine("  PowerPoint: .pptx .ppt");
            Console.WriteLine("  Web:        .html .htm (native)");
            Console.WriteLine("  Markdown:   .md    (Markdig CommonMark + YAML front matter)");
            Console.WriteLine("  JSON:       .json  (native)");
            Console.WriteLine("  Text:       .txt   (native)");
            Console.WriteLine();
            Console.WriteLine("Output files are written to the current working directory by");
            Console.WriteLine("default (.htm, or .txt when --plain-text is given), or to the");
            Console.WriteLine("directory given by -o / --output-dir if one is specified. Existing");
            Console.WriteLine("output files are skipped unless --force is given. The tool never");
            Console.WriteLine("overwrites a source file with its own converted output.");
            Console.WriteLine();
            Console.WriteLine("Exit codes:");
            Console.WriteLine("  0  all files converted (or help/version shown)");
            Console.WriteLine("  1  some files failed");
            Console.WriteLine("  2  fatal error (e.g., unknown option)");
        }

        // Expands one or more command-line file arguments into a
        // sorted, deduplicated list of absolute file paths. Each
        // argument may be:
        //
        //   - A literal filename (absolute or relative to the
        //     current directory). Emitted verbatim if it exists.
        //   - A wildcard pattern using * or ? (e.g., "*.xlsx",
        //     "report_????.csv", "data\*.xml"). Expanded via
        //     Directory.GetFiles.
        //
        // Multiple arguments are all walked; their results are
        // combined into a single list. A file that matches more
        // than one pattern (e.g., "*.xlsx" and "*q3*") is included
        // only once. Results are sorted case-insensitively by full
        // path for deterministic output.
        //
        // Office lock files (names starting with "~$") are silently
        // skipped — they are bookkeeping artifacts created when a
        // workbook is open in Excel, not files a user intends to
        // convert.
        // Extensions the tool knows how to convert. Used to filter
        // the results of a bare-directory expansion so the user
        // doesn't see spurious "Unsupported extension" errors for
        // every .png, .zip, .log, etc. that happens to be in the
        // folder. An explicit wildcard like "*.png" or an explicit
        // filename is NOT filtered — if the user specifically asked
        // for a file, the failure is informative.
        private static readonly HashSet<string> hsSupportedExts =
            new HashSet<string>(StringComparer.OrdinalIgnoreCase) {
                ".docx", ".doc", ".rtf", ".odt", ".pdf",
                ".xlsx", ".xls",
                ".pptx", ".ppt",
                ".csv",
                ".html", ".htm",
                ".md",
                ".json",
                ".txt"
            };

        private static List<string> expandWildcards(string[] asArgs)
        {
            var hsResult = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
            foreach (var sArgRaw in asArgs) {
                // If the argument is a bare directory path (no
                // wildcard, points to an existing directory), treat
                // it as "<dir>\*.*" — convert every supported file
                // inside that directory. This makes
                //     2htm C:\docs
                //     2htm -g   (with "C:\docs" typed into the GUI)
                // behave identically to
                //     2htm C:\docs\*.*
                // and filter the expansion so only known-supported
                // extensions become work items.
                string sArg = sArgRaw;
                bool bBareDir = false;
                if (sArg.IndexOfAny(new[] { '*', '?' }) < 0 &&
                    Directory.Exists(sArg)) {
                    sArg = Path.Combine(sArg, "*.*");
                    bBareDir = true;
                }

                string sDir = Path.GetDirectoryName(sArg);
                string sPattern = Path.GetFileName(sArg);
                if (string.IsNullOrEmpty(sDir)) sDir = Directory.GetCurrentDirectory();
                if (!Directory.Exists(sDir)) {
                    Console.Error.WriteLine("[WARN] Directory not found: " + sDir);
                    continue;
                }
                if (sPattern.IndexOfAny(new[] { '*', '?' }) < 0) {
                    string sFull = Path.GetFullPath(sArg);
                    if (File.Exists(sFull) && !isLockFile(sFull)) hsResult.Add(sFull);
                    else if (!File.Exists(sFull))
                        Console.Error.WriteLine("[WARN] File not found: " + sArg);
                } else {
                    string[] aMatches;
                    try {
                        aMatches = Directory.GetFiles(sDir, sPattern);
                    } catch (Exception ex) {
                        Console.Error.WriteLine("[WARN] Cannot enumerate '" + sArg + "': " + ex.Message);
                        continue;
                    }
                    if (aMatches.Length == 0)
                        Console.Error.WriteLine("[WARN] No files match: " + sArg);
                    foreach (var sMatch in aMatches) {
                        if (isLockFile(sMatch)) continue;
                        // For a bare-directory expansion, skip files
                        // whose extension we don't know how to
                        // handle. An explicit wildcard like "*.zip"
                        // keeps the file in — if the user asked for
                        // zips by name, let the converter raise
                        // "Unsupported extension" visibly.
                        if (bBareDir) {
                            string sExt = Path.GetExtension(sMatch);
                            if (!hsSupportedExts.Contains(sExt)) continue;
                        }
                        hsResult.Add(Path.GetFullPath(sMatch));
                    }
                }
            }
            var lsResult = new List<string>(hsResult);
            lsResult.Sort(StringComparer.OrdinalIgnoreCase);
            return lsResult;
        }

        private static bool isLockFile(string sPath)
        {
            string sName = Path.GetFileName(sPath);
            return !string.IsNullOrEmpty(sName) && sName.StartsWith("~$");
        }

        // Compute the output path for an input file. Single source of
        // truth for the rule "<input-basename-without-ext>.<htm|txt>"
        // in the effective output directory. Used by prePrune to
        // detect "output already exists" and by convertOne to know
        // where to write.
        private static string computeOutputPath(string sInPath)
        {
            string sOutExt = bPlainText ? ".txt" : ".htm";
            string sEffectiveOutputDir = string.IsNullOrEmpty(sOutputDir)
                ? Directory.GetCurrentDirectory()
                : sOutputDir;
            return Path.Combine(sEffectiveOutputDir,
                Path.GetFileNameWithoutExtension(sInPath) + sOutExt);
        }

        // Pre-prune the input list before the conversion loop runs.
        // Two passes:
        //
        //   1. Drop files whose extension 2htm cannot convert. These
        //      are dropped silently -- they do not appear in any
        //      result section since the user cannot un-skip them by
        //      checking "Force replacements". The log records each
        //      drop when -l is given.
        //
        //   2. If --force is NOT set, drop files whose target output
        //      already exists. These ARE counted as "skipped" and
        //      surface in the final results MessageBox along with a
        //      hint about the Force replacements checkbox.
        //
        // Also drops files where the input path equals the output
        // path (the self-overwrite guard, e.g. a .txt input in
        // --plain-text mode that would write back to itself). These
        // are silent like the unsupported drop -- the user cannot
        // un-skip them by toggling Force.
        //
        // Pre-pruning happens BEFORE the progress UI opens so the
        // counter denominator reflects only files that will actually
        // be processed.
        public static void prePrune(List<string> lsFiles,
            out List<string> lsToConvert,
            out List<string> lsSkippedExisting)
        {
            lsToConvert = new List<string>();
            lsSkippedExisting = new List<string>();
            foreach (var sFile in lsFiles) {
                string sExt = Path.GetExtension(sFile).ToLowerInvariant();
                if (!hsSupportedExts.Contains(sExt)) {
                    logger.info("Skipped (unsupported extension " + sExt + "): " + sFile);
                    continue;
                }
                string sOutPath = computeOutputPath(sFile);
                // Self-overwrite guard.
                if (string.Equals(Path.GetFullPath(sFile),
                                  Path.GetFullPath(sOutPath),
                                  StringComparison.OrdinalIgnoreCase)) {
                    logger.info("Skipped (cannot overwrite input with its own output): " + sFile);
                    continue;
                }
                if (File.Exists(sOutPath) && !bForce) {
                    lsSkippedExisting.Add(sFile);
                    logger.info("Skipped (" + Path.GetFileName(sOutPath) +
                        " already exists; use --force to overwrite): " + sFile);
                    continue;
                }
                lsToConvert.Add(sFile);
            }
        }

        // Converts a single input file to its output path. Throws on
        // any conversion failure. The caller (runConversionLoop) is
        // responsible for the per-file try/catch and for choosing
        // what to do with errors (typically: log and add to the
        // failed list).
        //
        // Pre-pruning (prePrune) has already filtered out unsupported
        // extensions, self-overwrite cases, and "target exists
        // without --force" cases, so anything reaching convertOne is
        // expected to actually convert.
        private static void convertOne(string sInPath)
        {
            string sExt = Path.GetExtension(sInPath).ToLowerInvariant();
            string sOutPath = computeOutputPath(sInPath);

            logger.info("Converting: " + sInPath + " -> " + sOutPath);

            // Pre-flight integrity check for Office formats. Saves us
            // a slow COM round-trip and a confusing RPC error message
            // when the file is clearly corrupt before Office even sees it.
            string sIntegrityError = fileIntegrity.check(sInPath, sExt);
            if (sIntegrityError != null)
                throw new InvalidDataException(sIntegrityError);

            // Run the conversion. If anything throws, delete any
            // partial output file the converter may have left
            // behind, then rethrow. We never want to leave a stale
            // or half-written output file that would mislead the
            // user into thinking conversion succeeded.
            bool bSucceeded = false;
            try {
                if (bPlainText) {
                    convertOneToText(sExt, sInPath, sOutPath);
                } else {
                    convertOneToHtml(sExt, sInPath, sOutPath);
                }
                bSucceeded = true;
                logger.info("Converted ok: " + sInPath);
            } finally {
                if (!bSucceeded) {
                    try {
                        if (File.Exists(sOutPath)) {
                            File.Delete(sOutPath);
                            logger.info("Deleted partial output: " + sOutPath);
                        }
                    } catch (Exception ex) {
                        logger.info("Could not delete partial output " + sOutPath + ": " + ex.Message);
                    }
                }
            }
        }

        // Runs the conversion loop over a pre-pruned list of files.
        // The caller is expected to have already filtered out
        // unsupported extensions, self-overwrite cases, and
        // "target exists without --force" cases via prePrune. Shared
        // by CLI mode (called directly from run()) and GUI mode
        // (called by guiProgress.runConversions). The onStarting
        // callback, if non-null, is invoked once per file before
        // convertOne runs, with the basename and a 1-based index
        // and total. In GUI mode it updates the status label; in
        // CLI mode it is null.
        //
        // Trim and shorten a message for inline display next to a
        // basename. Office COM exceptions can produce multi-paragraph
        // text; we want a single short line. Returns "" if the input
        // is null or empty.
        public static string firstLine(string s)
        {
            const int iMaxLen = 120;
            if (string.IsNullOrEmpty(s)) return "";
            int i = s.IndexOfAny(new[] { '\r', '\n' });
            if (i >= 0) s = s.Substring(0, i);
            s = s.Trim();
            if (s.Length > iMaxLen) s = s.Substring(0, iMaxLen - 3) + "...";
            return s;
        }

        // A single failure record: basename and a short reason. The
        // reason is whatever firstLine() gave us; the full exception
        // and stack trace are in the log. The structured summary at
        // the end uses this to render "basename: reason" on one line.
        public class failure
        {
            public string sBase;
            public string sReason;
            public failure(string sB, string sR) { sBase = sB; sReason = sR; }
        }

        // Runs the conversion loop. The caller has already pre-pruned
        // (unsupported extensions and target-exists cases dropped).
        // The onStarting callback updates the progress UI in GUI
        // mode; in pure CLI mode it is null and the loop instead
        // prints each basename inline as the file is processed.
        //
        // The bInlineConsole parameter selects between the two
        // output styles for the per-file progress:
        //   true  -- CLI mode: print basename to stdout as work
        //            begins; on success terminate with "\n"; on
        //            failure append ": <reason>\n".
        //   false -- GUI mode (and right-click): print nothing
        //            inline. The progress form shows the current
        //            file; the structured summary at the end shows
        //            the per-file outcomes.
        //
        // Outputs lsConverted (basenames of successes) and lsFailed
        // (failure records: basename + short reason). Their sum is
        // exactly lsToConvert.Count.
        public static void runConversionLoop(List<string> lsToConvert,
            Action<string, int, int> onStarting,
            bool bInlineConsole,
            out List<string> lsConverted,
            out List<failure> lsFailed)
        {
            lsConverted = new List<string>();
            lsFailed = new List<failure>();
            int iIndex = 0;
            foreach (var sFile in lsToConvert) {
                iIndex++;
                string sBase = Path.GetFileName(sFile);
                if (onStarting != null) {
                    try {
                        onStarting(sBase, iIndex, lsToConvert.Count);
                    } catch {
                        // Callback failure must not abort the run.
                    }
                }
                if (bInlineConsole) Console.Write(sBase);
                try {
                    convertOne(sFile);
                    lsConverted.Add(sBase);
                    if (bInlineConsole) Console.WriteLine();
                } catch (Exception ex) {
                    string sReason = firstLine(ex.Message);
                    lsFailed.Add(new failure(sBase, sReason));
                    if (bInlineConsole) Console.WriteLine(": " + sReason);
                    logger.error("Conversion failed: " + sFile);
                    logger.error("Exception: " + ex.GetType().FullName + ": " + ex.Message);
                    if (ex.InnerException != null)
                        logger.error("Inner: " + ex.InnerException.GetType().FullName +
                            ": " + ex.InnerException.Message);
                    logger.error("Stack trace:\r\n" + ex.StackTrace);
                }
            }
        }

        private static void convertOneToHtml(string sExt, string sInPath, string sOutPath)
        {
            switch (sExt) {
                case ".docx": case ".doc": case ".rtf": case ".odt": case ".pdf":
                    wordConverter.convert(sInPath, sOutPath); break;
                case ".xlsx": case ".xls":
                    excelConverter.convert(sInPath, sOutPath); break;
                case ".pptx": case ".ppt":
                    pptConverter.convert(sInPath, sOutPath); break;
                case ".csv":
                    csvConverter.convert(sInPath, sOutPath);
                    break;
                case ".html": case ".htm":
                    htmlConverter.convert(sInPath, sOutPath);
                    break;
                case ".md":
                    markdownConverter.convert(sInPath, sOutPath);
                    break;
                case ".json":
                    jsonConverter.convert(sInPath, sOutPath);
                    break;
                case ".txt":
                    textConverter.convert(sInPath, sOutPath);
                    break;
                default:
                    throw new NotSupportedException("Unsupported extension: " + sExt);
            }
        }

        private static void convertOneToText(string sExt, string sInPath, string sOutPath)
        {
            switch (sExt) {
                case ".docx": case ".doc": case ".rtf": case ".odt": case ".pdf":
                    wordConverter.convertToText(sInPath, sOutPath); break;
                case ".xlsx": case ".xls":
                    excelConverter.convertToText(sInPath, sOutPath); break;
                case ".pptx": case ".ppt":
                    pptConverter.convertToText(sInPath, sOutPath); break;
                case ".csv":
                    textPassthrough.copy(sInPath, sOutPath);
                    break;
                case ".html": case ".htm":
                    htmlConverter.convertToText(sInPath, sOutPath);
                    break;
                case ".md":
                    markdownConverter.convertToText(sInPath, sOutPath);
                    break;
                case ".json":
                    jsonConverter.convertToText(sInPath, sOutPath);
                    break;
                case ".txt":
                    textPassthrough.copy(sInPath, sOutPath);
                    break;
                default:
                    throw new NotSupportedException("Unsupported extension: " + sExt);
            }
        }
    }

    // -----------------------------------------------------------------
    // Simple WinForms entry dialog shown in GUI mode (-g / --gui-mode).
    // Collects the same inputs the command line accepts, preloaded with
    // whatever values the command line already supplied. On OK the
    // values are copied back into program's globals; on Cancel the
    // application exits without converting anything.
    //
    // Accessibility notes:
    //   - Labels use AutomationId via Label.Text so screen readers
    //     announce the field name when focus lands on the associated
    //     text box (the label comes immediately before its control in
    //     both tab order and visual order).
    //   - Access keys (alt-prefixed) are set on all interactive
    //     controls except OK and Cancel, which use Enter and Esc per
    //     Windows UX guidelines.
    //   - The form is center-screen, single-instance, modal. It does
    //     not minimize or resize — screen reader navigation benefits
    //     from a stable control geometry.
    // -----------------------------------------------------------------
    public static class guiDialog
    {
        public static bool show(ref string sSource, ref string sOutputDir,
            ref bool bStrip, ref bool bPlain, ref bool bForce, ref bool bView,
            ref bool bLog, ref bool bUseCfg)
        {
            // Build the form.
            var frm = new System.Windows.Forms.Form();
            frm.Text = program.sProgramName;
            frm.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedDialog;
            frm.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen;
            frm.MaximizeBox = false;
            frm.MinimizeBox = false;
            frm.ShowInTaskbar = false;
            frm.ClientSize = new System.Drawing.Size(560, 220);
            frm.Font = System.Drawing.SystemFonts.MessageBoxFont;

            // Route F1 to the same action as clicking the Help
            // button: show the help MessageBox with an option to
            // launch the full HTML documentation. KeyPreview lets
            // the form see the keystroke before whatever child
            // control currently has focus consumes it. F1 is the
            // standard Windows help shortcut and is expected by
            // keyboard-driven users, including screen-reader users.
            frm.KeyPreview = true;
            frm.KeyDown += (s, e) => {
                if (e.KeyCode == System.Windows.Forms.Keys.F1) {
                    e.Handled = true;
                    e.SuppressKeyPress = true;
                    showHelpMessage();
                }
            };

            // Layout constants in pixels. These deliberately reflect
            // Windows desktop conventions: ~7 px gutter, ~11 px row
            // spacing, button width ~120 px.
            const int iLayoutLeft = 12;
            const int iLayoutRight = 12;
            const int iLayoutTop = 12;
            const int iLayoutGap = 7;
            const int iLayoutRowGap = 11;
            const int iLayoutLabelWidth = 110;
            const int iLayoutButtonWidth = 130;
            const int iLayoutButtonHeight = 26;
            const int iLayoutTextHeight = 23;

            int iFormW = frm.ClientSize.Width;
            int iTextX = iLayoutLeft + iLayoutLabelWidth + iLayoutGap;
            int iTextW = iFormW - iTextX - iLayoutGap - iLayoutButtonWidth - iLayoutRight;
            int iBtnX = iFormW - iLayoutRight - iLayoutButtonWidth;

            // --- Row 1: Source files ---
            int y = iLayoutTop;
            var lblSource = new System.Windows.Forms.Label();
            lblSource.Text = "&Source files:";
            lblSource.AutoSize = false;
            lblSource.Location = new System.Drawing.Point(iLayoutLeft, y + 3);
            lblSource.Size = new System.Drawing.Size(iLayoutLabelWidth, iLayoutTextHeight);
            lblSource.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
            frm.Controls.Add(lblSource);

            var txtSource = new System.Windows.Forms.TextBox();
            // Default when nothing was supplied on the command line
            // and no saved config filled the field: the user's
            // Documents folder. This is the Microsoft-recommended
            // starting point for file operations in a user-facing
            // GUI (it's what every Office file dialog opens to by
            // default). A command-line user invoking 2htm -g from a
            // specific directory on the command line still gets
            // their own path if they passed one.
            txtSource.Text = string.IsNullOrWhiteSpace(sSource)
                ? defaultSourceForGui()
                : sSource;
            txtSource.Location = new System.Drawing.Point(iTextX, y);
            txtSource.Size = new System.Drawing.Size(iTextW, iLayoutTextHeight);
            txtSource.TabIndex = 0;
            frm.Controls.Add(txtSource);

            var btnBrowseSource = new System.Windows.Forms.Button();
            btnBrowseSource.Text = "&Browse source...";
            btnBrowseSource.Location = new System.Drawing.Point(iBtnX, y - 1);
            btnBrowseSource.Size = new System.Drawing.Size(iLayoutButtonWidth, iLayoutButtonHeight);
            btnBrowseSource.TabIndex = 1;
            btnBrowseSource.UseVisualStyleBackColor = true;
            // Click handler wired below, after txtOut is declared
            // (so the handler can reference it to auto-fill).
            frm.Controls.Add(btnBrowseSource);

            // --- Row 2: Output directory ---
            y += iLayoutTextHeight + iLayoutRowGap;
            var lblOut = new System.Windows.Forms.Label();
            lblOut.Text = "&Output directory:";
            lblOut.AutoSize = false;
            lblOut.Location = new System.Drawing.Point(iLayoutLeft, y + 3);
            lblOut.Size = new System.Drawing.Size(iLayoutLabelWidth, iLayoutTextHeight);
            lblOut.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
            frm.Controls.Add(lblOut);

            var txtOut = new System.Windows.Forms.TextBox();
            // If no output directory is supplied (from CLI or saved
            // config), pre-populate from the source textbox's
            // current value rather than leaving the user staring
            // at a blank field. This makes the "blank = same as
            // source" semantics visible upfront and gives the user
            // a starting point to edit from. If the source can't
            // be resolved to a directory (multiple patterns, bad
            // path), leave the field blank.
            if (!string.IsNullOrWhiteSpace(sOutputDir)) {
                txtOut.Text = sOutputDir;
            } else {
                txtOut.Text = deriveOutputDirFromSource(txtSource.Text);
            }
            txtOut.Location = new System.Drawing.Point(iTextX, y);
            txtOut.Size = new System.Drawing.Size(iTextW, iLayoutTextHeight);
            txtOut.TabIndex = 2;
            frm.Controls.Add(txtOut);

            var btnBrowseOut = new System.Windows.Forms.Button();
            btnBrowseOut.Text = "&Choose output...";
            btnBrowseOut.Location = new System.Drawing.Point(iBtnX, y - 1);
            btnBrowseOut.Size = new System.Drawing.Size(iLayoutButtonWidth, iLayoutButtonHeight);
            btnBrowseOut.TabIndex = 3;
            btnBrowseOut.UseVisualStyleBackColor = true;
            btnBrowseOut.Click += (s, e) => {
                string sPicked = pickFolder("Choose the output directory",
                    txtOut.Text);
                if (sPicked != null) txtOut.Text = sPicked;
            };
            frm.Controls.Add(btnBrowseOut);

            // Wire the source-browse handler now that txtOut is in
            // scope. When the user picks a new source folder, the
            // output field auto-follows IF the user hasn't
            // explicitly edited it — detected by checking whether
            // the current output value matches what the OLD source
            // would have auto-derived. If the user typed something
            // else in there, leave it alone.
            btnBrowseSource.Click += (s, e) => {
                string sPicked = pickFolder("Choose a folder of files to convert",
                    txtSource.Text);
                if (sPicked != null) {
                    string sOldDerived = deriveOutputDirFromSource(txtSource.Text);
                    bool bOutputMirrorsSource =
                        string.IsNullOrWhiteSpace(txtOut.Text) ||
                        string.Equals(txtOut.Text.Trim(), sOldDerived,
                            StringComparison.OrdinalIgnoreCase);
                    txtSource.Text = sPicked;
                    if (bOutputMirrorsSource) {
                        txtOut.Text = deriveOutputDirFromSource(sPicked);
                    }
                }
            };

            // --- Row 3+4: Option checkboxes (2x2 grid) ---
            //   [ Strip images ]   [ Plain text ]
            //   [ Force replacements ] [ View output ]
            // A 2x2 grid gives each label room to display fully and
            // aligns checkboxes predictably for screen-reader
            // traversal.
            y += iLayoutTextHeight + iLayoutRowGap * 2;
            int iChkW = (iFormW - iLayoutLeft - iLayoutRight) / 2;
            var chkStrip = new System.Windows.Forms.CheckBox();
            chkStrip.Text = "Strip &images";
            chkStrip.Checked = bStrip;
            chkStrip.Location = new System.Drawing.Point(iLayoutLeft, y);
            chkStrip.Size = new System.Drawing.Size(iChkW, iLayoutTextHeight);
            chkStrip.TabIndex = 4;
            frm.Controls.Add(chkStrip);

            var chkPlain = new System.Windows.Forms.CheckBox();
            chkPlain.Text = "&Plain text";
            chkPlain.Checked = bPlain;
            chkPlain.Location = new System.Drawing.Point(iLayoutLeft + iChkW, y);
            chkPlain.Size = new System.Drawing.Size(iChkW, iLayoutTextHeight);
            chkPlain.TabIndex = 5;
            frm.Controls.Add(chkPlain);

            y += iLayoutTextHeight + iLayoutRowGap;
            var chkForce = new System.Windows.Forms.CheckBox();
            chkForce.Text = "&Force replacements";
            chkForce.Checked = bForce;
            chkForce.Location = new System.Drawing.Point(iLayoutLeft, y);
            chkForce.Size = new System.Drawing.Size(iChkW, iLayoutTextHeight);
            chkForce.TabIndex = 6;
            frm.Controls.Add(chkForce);

            var chkView = new System.Windows.Forms.CheckBox();
            chkView.Text = "&View output";
            chkView.Checked = bView;
            chkView.Location = new System.Drawing.Point(iLayoutLeft + iChkW, y);
            chkView.Size = new System.Drawing.Size(iChkW, iLayoutTextHeight);
            chkView.TabIndex = 7;
            frm.Controls.Add(chkView);

            // Third row: Log session and Use configuration. Both
            // are "meta" options that affect persistence/diagnostics
            // rather than the conversion itself, so they sit
            // together below the conversion-control checkboxes.
            y += iLayoutTextHeight + iLayoutRowGap;
            var chkLog = new System.Windows.Forms.CheckBox();
            chkLog.Text = "&Log session";
            chkLog.Checked = bLog;
            chkLog.Location = new System.Drawing.Point(iLayoutLeft, y);
            chkLog.Size = new System.Drawing.Size(iChkW, iLayoutTextHeight);
            chkLog.TabIndex = 8;
            frm.Controls.Add(chkLog);

            var chkUseCfg = new System.Windows.Forms.CheckBox();
            chkUseCfg.Text = "&Use configuration";
            chkUseCfg.Checked = bUseCfg;
            chkUseCfg.Location = new System.Drawing.Point(iLayoutLeft + iChkW, y);
            chkUseCfg.Size = new System.Drawing.Size(iChkW, iLayoutTextHeight);
            chkUseCfg.TabIndex = 9;
            frm.Controls.Add(chkUseCfg);

            // --- Bottom row: commit buttons per Windows UX ---
            // Help and Default settings on the left (they don't
            // commit or cancel the dialog), OK and Cancel on the
            // right. This matches Microsoft's UX guidance for
            // secondary dialogs.
            y += iLayoutTextHeight + iLayoutRowGap * 2;
            var btnHelp = new System.Windows.Forms.Button();
            btnHelp.Text = "&Help";
            btnHelp.Location = new System.Drawing.Point(iLayoutLeft, y);
            btnHelp.Size = new System.Drawing.Size(iLayoutButtonWidth, iLayoutButtonHeight);
            btnHelp.TabIndex = 10;
            btnHelp.UseVisualStyleBackColor = true;
            btnHelp.Click += (s, e) => showHelpMessage();
            frm.Controls.Add(btnHelp);

            var btnDefaults = new System.Windows.Forms.Button();
            btnDefaults.Text = "&Default settings";
            btnDefaults.Location = new System.Drawing.Point(
                iLayoutLeft + iLayoutButtonWidth + iLayoutGap, y);
            btnDefaults.Size = new System.Drawing.Size(iLayoutButtonWidth, iLayoutButtonHeight);
            btnDefaults.TabIndex = 11;
            btnDefaults.UseVisualStyleBackColor = true;
            // Default settings: full reset. Resets the dialog's
            // fields to factory defaults AND deletes the saved
            // configuration file (plus the 2htm directory under
            // %LOCALAPPDATA% if it becomes empty). This is the
            // explicit way for a user to fully opt out of the
            // footprint — "start over, leave no trace." The
            // deletion happens immediately on click, not on OK; a
            // user who clicks Default settings and then Cancel has
            // still cleared the saved config.
            btnDefaults.Click += (s, e) => {
                string sDefault = defaultSourceForGui();
                txtSource.Text = sDefault;
                txtOut.Text = deriveOutputDirFromSource(sDefault);
                chkStrip.Checked = false;
                chkPlain.Checked = false;
                chkForce.Checked = false;
                chkView.Checked = false;
                chkLog.Checked = false;
                chkUseCfg.Checked = false;
                configManager.eraseAll();
            };
            frm.Controls.Add(btnDefaults);

            var btnOk = new System.Windows.Forms.Button();
            btnOk.Text = "OK";
            btnOk.DialogResult = System.Windows.Forms.DialogResult.OK;
            btnOk.Location = new System.Drawing.Point(
                iFormW - iLayoutRight - 2 * iLayoutButtonWidth - iLayoutGap, y);
            btnOk.Size = new System.Drawing.Size(iLayoutButtonWidth, iLayoutButtonHeight);
            btnOk.TabIndex = 12;
            btnOk.UseVisualStyleBackColor = true;
            // Validate output directory before allowing the dialog to
            // close. If the user has typed a non-existent directory,
            // prompt to create it (default Yes). On No or creation
            // failure, set DialogResult = None so the dialog stays
            // open. WinForms invokes Click handlers BEFORE the
            // automatic close, so this hook runs first.
            btnOk.Click += (s, e) => {
                string sOutCandidate = (txtOut.Text ?? "").Trim();
                if (sOutCandidate.Length >= 2 && sOutCandidate[0] == '"' && sOutCandidate[sOutCandidate.Length - 1] == '"')
                    sOutCandidate = sOutCandidate.Substring(1, sOutCandidate.Length - 2).Trim();
                if (string.IsNullOrEmpty(sOutCandidate)) return;
                try {
                    if (System.IO.Directory.Exists(sOutCandidate)) return;
                } catch { return; }
                System.Windows.Forms.DialogResult dr = System.Windows.Forms.MessageBox.Show(frm,
                    "Create " + sOutCandidate + "?",
                    program.sProgramName,
                    System.Windows.Forms.MessageBoxButtons.YesNo,
                    System.Windows.Forms.MessageBoxIcon.Question,
                    System.Windows.Forms.MessageBoxDefaultButton.Button1);
                if (dr != System.Windows.Forms.DialogResult.Yes) {
                    frm.DialogResult = System.Windows.Forms.DialogResult.None;
                    txtOut.Focus();
                    return;
                }
                try {
                    System.IO.Directory.CreateDirectory(sOutCandidate);
                } catch (Exception ex) {
                    System.Windows.Forms.MessageBox.Show(frm,
                        "Could not create directory:\r\n" + sOutCandidate + "\r\n\r\n" + ex.Message,
                        program.sProgramName,
                        System.Windows.Forms.MessageBoxButtons.OK,
                        System.Windows.Forms.MessageBoxIcon.Warning);
                    frm.DialogResult = System.Windows.Forms.DialogResult.None;
                    txtOut.Focus();
                }
            };
            frm.Controls.Add(btnOk);

            var btnCancel = new System.Windows.Forms.Button();
            btnCancel.Text = "Cancel";
            btnCancel.DialogResult = System.Windows.Forms.DialogResult.Cancel;
            btnCancel.Location = new System.Drawing.Point(iBtnX, y);
            btnCancel.Size = new System.Drawing.Size(iLayoutButtonWidth, iLayoutButtonHeight);
            btnCancel.TabIndex = 13;
            btnCancel.UseVisualStyleBackColor = true;
            frm.Controls.Add(btnCancel);

            // Wire Enter = OK and Esc = Cancel per Windows guidelines.
            frm.AcceptButton = btnOk;
            frm.CancelButton = btnCancel;

            // Adjust form height to the last control's bottom + margin.
            frm.ClientSize = new System.Drawing.Size(iFormW,
                y + iLayoutButtonHeight + iLayoutTop);

            // Show it.
            var dialogResult = frm.ShowDialog();
            if (dialogResult != System.Windows.Forms.DialogResult.OK) return false;

            // Hand values back.
            sSource = txtSource.Text.Trim();
            sOutputDir = txtOut.Text.Trim();
            // If the output directory is still blank at submission
            // time, derive a sensible default from the source field.
            // For a single entry that is a folder or a
            // wildcard/file, use the containing folder; for multiple
            // entries, leave blank (caller falls back to the current
            // working directory).
            if (string.IsNullOrEmpty(sOutputDir))
                sOutputDir = deriveOutputDirFromSource(sSource);
            bStrip = chkStrip.Checked;
            bPlain = chkPlain.Checked;
            bForce = chkForce.Checked;
            bView = chkView.Checked;
            bLog = chkLog.Checked;
            bUseCfg = chkUseCfg.Checked;
            return true;
        }

        // Sensible starting path for the Source files textbox when
        // nothing is supplied on the command line and no saved
        // configuration is in play. Uses the user's Documents
        // folder — Microsoft's recommended starting point for
        // user-facing file operations (it's what every Office file
        // dialog defaults to). If that folder can't be resolved,
        // falls back to the current working directory.
        private static string defaultSourceForGui()
        {
            try {
                string sDocs = Environment.GetFolderPath(
                    Environment.SpecialFolder.MyDocuments);
                if (!string.IsNullOrEmpty(sDocs) &&
                    System.IO.Directory.Exists(sDocs)) return sDocs;
            } catch { }
            return System.IO.Directory.GetCurrentDirectory();
        }

        // Given whatever the user typed into the Source files field,
        // return a single directory that should serve as the output
        // default, or empty string if one can't be confidently
        // inferred. Rules:
        //   - Empty source → empty output (caller falls back to CWD).
        //   - A single existing directory → that directory.
        //   - A single wildcard pattern → directory containing it.
        //   - A single existing file → that file's parent directory.
        //   - Multiple entries, or anything ambiguous → empty
        //     (output-dir remains blank; caller falls back to CWD).
        private static string deriveOutputDirFromSource(string sSourceField)
        {
            if (string.IsNullOrWhiteSpace(sSourceField)) return "";
            var ls = program.splitSourceField(sSourceField);
            if (ls.Count != 1) return "";
            string sOne = ls[0];
            try {
                if (System.IO.Directory.Exists(sOne)) return sOne;
                if (sOne.IndexOfAny(new[] { '*', '?' }) >= 0) {
                    string sDir = System.IO.Path.GetDirectoryName(sOne);
                    if (!string.IsNullOrEmpty(sDir) &&
                        System.IO.Directory.Exists(sDir))
                        return sDir;
                    return "";
                }
                if (System.IO.File.Exists(sOne)) {
                    string sDir = System.IO.Path.GetDirectoryName(sOne);
                    return sDir ?? "";
                }
            } catch { }
            return "";
        }

        // Opens a FolderBrowserDialog initialized at the given path
        // (or a reasonable default if it is empty or invalid).
        // Returns null if the user cancels.
        private static string pickFolder(string sDesc, string sSeed)
        {
            // Pick a sensible initial folder. The strategy follows
            // Microsoft's guidance: start at the user's most recent
            // intent (the seed -- typically the current text in the
            // Source files or Output directory field) when one is
            // usable, otherwise fall back to the user's Documents
            // folder.
            //
            // The seed may be:
            //   - empty
            //   - a folder path
            //   - a wildcard pattern (e.g. C:\docs\*.xlsx)
            //   - a single file path
            //   - a space-separated list of any of the above
            //   - a quoted path containing spaces
            //   - a non-existent path the user typed manually
            string sInitial = "";
            try {
                string sCandidate = (sSeed ?? "").Trim();
                if (!string.IsNullOrEmpty(sCandidate)) {
                    // Inspect the first space-separated token only.
                    int iSpace = sCandidate.IndexOf(' ');
                    if (iSpace >= 0) sCandidate = sCandidate.Substring(0, iSpace);
                    sCandidate = sCandidate.Trim('"');
                    // Wildcards: strip the basename and inspect the parent.
                    if (sCandidate.IndexOfAny(new[] { '*', '?' }) >= 0)
                        sCandidate = System.IO.Path.GetDirectoryName(sCandidate) ?? "";
                    if (!string.IsNullOrEmpty(sCandidate)) {
                        if (System.IO.Directory.Exists(sCandidate)) {
                            sInitial = System.IO.Path.GetFullPath(sCandidate);
                        } else if (System.IO.File.Exists(sCandidate)) {
                            sInitial = System.IO.Path.GetDirectoryName(
                                System.IO.Path.GetFullPath(sCandidate));
                        } else {
                            string sParent = System.IO.Path.GetDirectoryName(sCandidate);
                            if (!string.IsNullOrEmpty(sParent) &&
                                System.IO.Directory.Exists(sParent))
                                sInitial = System.IO.Path.GetFullPath(sParent);
                        }
                    }
                }
            } catch { sInitial = ""; }

            // Documents fallback when no usable seed.
            if (string.IsNullOrEmpty(sInitial)) {
                try {
                    sInitial = Environment.GetFolderPath(
                        Environment.SpecialFolder.MyDocuments);
                } catch { sInitial = ""; }
            }

            using (var dlg = new System.Windows.Forms.FolderBrowserDialog()) {
                dlg.Description = sDesc;
                dlg.ShowNewFolderButton = true;
                if (!string.IsNullOrEmpty(sInitial)) dlg.SelectedPath = sInitial;
                if (dlg.ShowDialog() == System.Windows.Forms.DialogResult.OK)
                    return dlg.SelectedPath;
            }
            return null;
        }

        private static void showHelpMessage()
        {
            string sMsg =
                "2htm converts documents (Word, Excel, PowerPoint, PDF, " +
                "Markdown, and more) to accessible HTML or plain text.\r\n\r\n" +
                "Fill in the source files to convert, pick any options, " +
                "and press OK. Leave the output directory blank to write " +
                "results to the current folder.\r\n\r\n" +
                "Options:\r\n" +
                "  Strip images - omit images from the output\r\n" +
                "  Plain text - write .txt instead of .htm\r\n" +
                "  Force replacements - overwrite existing output files\r\n" +
                "  View output - open the output folder when done\r\n" +
                "  Log session - write 2htm.log (replacing any prior log) " +
                "to the output folder, or to the current folder if no " +
                "output folder is set\r\n" +
                "  Use configuration - remember these settings for next time, " +
                "in %LOCALAPPDATA%\\2htm\\2htm.ini\r\n\r\n" +
                "Press Cancel to exit without converting.\r\n\r\n" +
                "Open the full documentation in your browser?";
            var dialogResult = System.Windows.Forms.MessageBox.Show(sMsg,
                "2htm — Help",
                System.Windows.Forms.MessageBoxButtons.YesNo,
                System.Windows.Forms.MessageBoxIcon.Information,
                System.Windows.Forms.MessageBoxDefaultButton.Button2);
            if (dialogResult == System.Windows.Forms.DialogResult.Yes) {
                launchReadMe();
            }
        }

        // Opens readMe.htm in the user's default browser. readMe.htm
        // lives alongside 2htm.exe (generated by the installer at
        // install time, or put there by the user when deploying
        // 2htm standalone). If the file is missing, fall back to
        // readMe.md, and if that's also missing, show a polite
        // notice rather than a system-level "file not found" error.
        private static void launchReadMe()
        {
            string sExeDir = Path.GetDirectoryName(
                System.Reflection.Assembly.GetExecutingAssembly().Location);
            string sHtm = Path.Combine(sExeDir, "readMe.htm");
            string sMd = Path.Combine(sExeDir, "readMe.md");
            string sTarget = File.Exists(sHtm)
                ? sHtm
                : (File.Exists(sMd) ? sMd : null);
            if (sTarget == null) {
                System.Windows.Forms.MessageBox.Show(
                    "Documentation (readMe.htm or readMe.md) was not found " +
                    "in the 2htm install folder:\r\n\r\n" + sExeDir + "\r\n\r\n" +
                    "If 2htm was installed via the installer, reinstall it. " +
                    "If you deployed 2htm.exe manually, place readMe.htm " +
                    "(or readMe.md) in the same folder.",
                    "2htm — Documentation not found",
                    System.Windows.Forms.MessageBoxButtons.OK,
                    System.Windows.Forms.MessageBoxIcon.Warning);
                return;
            }
            try {
                var processStartInfo = new System.Diagnostics.ProcessStartInfo(sTarget) {
                    UseShellExecute = true
                };
                System.Diagnostics.Process.Start(processStartInfo);
            } catch (Exception ex) {
                System.Windows.Forms.MessageBox.Show(
                    "Could not open the documentation:\r\n\r\n" + ex.Message,
                    "2htm — Error",
                    System.Windows.Forms.MessageBoxButtons.OK,
                    System.Windows.Forms.MessageBoxIcon.Warning);
            }
        }
    }

    // -----------------------------------------------------------------
    // Progress dialog shown during GUI-mode conversion runs. Displays
    // a single status line with the current file's basename and a
    // running "N of M (P%)" indicator.
    //
    // Implementation: single-threaded. The form is shown non-modally
    // and the conversion loop runs on this same (UI) thread. Between
    // files, Application.DoEvents() pumps the Windows message queue
    // so the form repaints and screen readers observe the status-
    // label change. This keeps the code simple and avoids the STA /
    // COM-apartment problems that come with background-thread
    // automation of Office. The tradeoff is that during the
    // conversion of any single very slow file, the UI cannot
    // repaint — but that's a property of the file, not of the
    // threading model, and for typical batches of reasonable-sized
    // documents the per-file status updates give adequate feedback.
    //
    // Accessibility: the status label carries AccessibleName and
    // AccessibleRole so screen readers announce each change. The
    // form has no buttons, so there is no re-entrancy concern from
    // DoEvents pumping a stray click.
    // -----------------------------------------------------------------
    // -----------------------------------------------------------------
    // guiProgress: a small modeless status form shown during the file-
    // conversion loop in GUI mode (and in right-click invocations).
    // The same three-method API (open / update / close) is used by
    // extCheck. The label uses AccessibleRole.StatusBar so screen
    // readers announce text changes; Application.DoEvents() pumps the
    // message queue so the new text actually paints between files.
    //
    // The displayed count reflects files ALREADY COMPLETED, not the
    // file being started. When starting file 1 of 5, the label shows
    // "file.pdf -- 0 of 5, 0%". The percent and count advance only
    // after the file finishes. This avoids the confusion of seeing
    // "100%" while the (only) file is still being processed -- a
    // real user complaint.
    //
    // The form has no buttons, so DoEvents pumping is safe (no stray
    // click can re-enter the conversion code).
    // -----------------------------------------------------------------
    public static class guiProgress
    {
        private static System.Windows.Forms.Form frm;
        private static System.Windows.Forms.Label lblStatus;

        public static void open(int iTotal)
        {
            frm = new System.Windows.Forms.Form();
            frm.Text = "2htm \u2014 Converting";
            frm.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedDialog;
            frm.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen;
            frm.MaximizeBox = false;
            frm.MinimizeBox = false;
            frm.ControlBox = false;
            frm.ShowInTaskbar = true;
            frm.ClientSize = new System.Drawing.Size(480, 92);
            frm.Font = System.Drawing.SystemFonts.MessageBoxFont;

            var lblIntro = new System.Windows.Forms.Label();
            lblIntro.Text = "Converting files. Please wait...";
            lblIntro.AutoSize = false;
            lblIntro.Location = new System.Drawing.Point(14, 14);
            lblIntro.Size = new System.Drawing.Size(452, 22);
            lblIntro.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
            frm.Controls.Add(lblIntro);

            lblStatus = new System.Windows.Forms.Label();
            lblStatus.Text = "Starting...";
            lblStatus.AutoSize = false;
            lblStatus.Location = new System.Drawing.Point(14, 42);
            lblStatus.Size = new System.Drawing.Size(452, 22);
            lblStatus.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
            lblStatus.AccessibleName = "Conversion status";
            lblStatus.AccessibleRole = System.Windows.Forms.AccessibleRole.StatusBar;
            frm.Controls.Add(lblStatus);

            frm.Show();
            System.Windows.Forms.Application.DoEvents();
        }

        public static void update(string sBase, int iIndex, int iTotal)
        {
            if (frm == null || lblStatus == null) return;
            // iIndex is 1-based; we display "iIndex - 1" as the
            // completed count, so the percentage and count reflect
            // work DONE, not work in progress.
            int iCompleted = iIndex - 1;
            int iPercent = iTotal > 0 ? (iCompleted * 100 / iTotal) : 0;
            lblStatus.Text = sBase + " \u2014 " + iCompleted + " of " + iTotal +
                ", " + iPercent + "%";
            System.Windows.Forms.Application.DoEvents();
        }

        public static void close()
        {
            if (frm == null) return;
            try { frm.Close(); frm.Dispose(); } catch { }
            frm = null;
            lblStatus = null;
        }

        // Wrapper that opens the progress form, runs the conversion
        // loop with the update callback, and closes the form. The
        // call site in run() uses this; passes a pre-pruned list
        // and gets back two lists of basenames (converted + failed).
        public static void runConversions(List<string> lsToConvert,
            out List<string> lsConverted,
            out List<program.failure> lsFailed)
        {
            open(lsToConvert.Count);
            try {
                Action<string, int, int> onStarting = (sBase, iIndex, iTotal) =>
                    update(sBase, iIndex, iTotal);
                program.runConversionLoop(lsToConvert, onStarting,
                    /* bInlineConsole: */ false,
                    out lsConverted, out lsFailed);
            } finally {
                close();
            }
        }
    }

    // -----------------------------------------------------------------
    // Diagnostic logger. Off by default; enabled with --log / -l or
    // by checking "Log session" in the GUI dialog.
    //
    // When enabled, writes a UTF-8 file named 2htm.log (with BOM) to
    // the output directory if one was specified (-o or the GUI Output
    // directory field), or to the current directory otherwise. Each
    // line is prefixed with an ISO-8601 timestamp and severity level.
    // The log stream is flushed after every write so that if the
    // process crashes, the log captures everything up to the crash
    // point. Each session starts with a fresh file -- any prior
    // 2htm.log in the same location is overwritten, so the file
    // always reflects only the current run.
    //
    // All methods no-op silently when the log is not open, so call
    // sites can log unconditionally without guarding each call.
    //
    // Thread-safety: 2htm is single-threaded with respect to the
    // main conversion flow, so no locking is required. Office COM
    // callbacks run on the main STA thread.
    // -----------------------------------------------------------------
    public static class logger
    {
        private static StreamWriter writer = null;

        public static void open(string sDir = "")
        {
            if (writer != null) return;
            // Resolve the directory: use sDir if non-empty and it is
            // (or can be) an existing directory; otherwise fall back
            // to the current directory. We deliberately do NOT create
            // the directory ourselves; that's the responsibility of
            // the conversion code which has the same writability
            // expectations.
            string sLogDir;
            try {
                if (!string.IsNullOrWhiteSpace(sDir) && Directory.Exists(sDir)) {
                    sLogDir = Path.GetFullPath(sDir);
                } else {
                    sLogDir = Directory.GetCurrentDirectory();
                }
            } catch {
                sLogDir = Directory.GetCurrentDirectory();
            }
            string sPath = Path.Combine(sLogDir, program.sLogFileName);
            try {
                // UTF-8 WITH BOM so the user's editor/screen reader
                // knows the encoding without guessing. append:false
                // truncates any prior log so the file reflects only
                // the current session.
                writer = new StreamWriter(sPath, append: false, encoding: new UTF8Encoding(true));
                writer.AutoFlush = true;
            } catch (Exception ex) {
                Console.Error.WriteLine("[WARN] Could not open log file '" + sPath + "': " +
                    ex.Message + ". Continuing without a log.");
                writer = null;
            }
        }

        public static void close()
        {
            if (writer == null) return;
            try {
                writer.WriteLine(stamp("INFO") + " Log closed.");
                writer.Flush();
                writer.Close();
            } catch { }
            writer = null;
        }

        public static void info(string sMsg)  { write("INFO", sMsg); }
        public static void warn(string sMsg)  { write("WARN", sMsg); }
        public static void error(string sMsg) { write("ERROR", sMsg); }
        public static void debug(string sMsg) { write("DEBUG", sMsg); }

        // Write the run header to the top of the log: program name
        // and version, the friendly run-start timestamp, and the
        // resolved parameter list. Emits raw lines (no per-line
        // timestamp/level prefix) so the header reads as a clean
        // banner. The processing notifications that follow use the
        // standard timestamped format via info/warn/error/debug.
        //
        // dParams is a List<KeyValuePair<string,string>> rather than
        // a Dictionary so the caller controls the order in which
        // parameters appear.
        public static void header(string sName, string sVersion,
            List<KeyValuePair<string, string>> dParams)
        {
            if (writer == null) return;
            try {
                writer.WriteLine("=== " + sName + " " + sVersion + " ===");
                writer.WriteLine("Run on " + friendlyTime(DateTime.Now));
                if (dParams != null && dParams.Count > 0) {
                    writer.WriteLine("Parameters:");
                    int iPad = 0;
                    foreach (var oKv in dParams)
                        if (oKv.Key.Length > iPad) iPad = oKv.Key.Length;
                    foreach (var oKv in dParams)
                        writer.WriteLine("  " + oKv.Key.PadRight(iPad) + " : " + oKv.Value);
                }
                writer.WriteLine("===");
            } catch { }
        }

        // Render a DateTime in a friendly form, e.g.,
        // "May 1, 2026 at 2:30 PM". Uses invariant culture for
        // stable output across machines.
        public static string friendlyTime(DateTime dt)
        {
            return dt.ToString("MMMM d, yyyy", CultureInfo.InvariantCulture) +
                " at " +
                dt.ToString("h:mm tt", CultureInfo.InvariantCulture);
        }

        private static void write(string sLevel, string sMsg)
        {
            if (writer == null) return;
            try {
                writer.WriteLine(stamp(sLevel) + " " + sMsg);
            } catch { }
        }

        private static string stamp(string sLevel)
        {
            return DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss.fff",
                CultureInfo.InvariantCulture) + " " + sLevel.PadRight(5);
        }
    }

    // -----------------------------------------------------------------
    // File integrity checks. Modern Office formats (.docx/.xlsx/.pptx)
    // are ZIP archives with specific marker files inside. Legacy
    // formats (.doc/.xls/.ppt) are OLE Compound Files with a specific
    // magic number. Returns null if OK, or a reason-string if bad.
    // -----------------------------------------------------------------
    public static class fileIntegrity
    {
        // OLE Compound File magic: D0 CF 11 E0 A1 B1 1A E1
        private static readonly byte[] binOleMagic =
            { 0xD0, 0xCF, 0x11, 0xE0, 0xA1, 0xB1, 0x1A, 0xE1 };

        public static string check(string sPath, string sExtLower)
        {
            try {
                var fileInfo = new FileInfo(sPath);
                if (!fileInfo.Exists) return "File does not exist.";
                if (fileInfo.Length == 0) return "File is empty (zero bytes).";

                switch (sExtLower) {
                    case ".docx": case ".xlsx": case ".pptx":
                        return checkOpenXml(sPath, sExtLower);
                    case ".doc": case ".xls": case ".ppt":
                        return checkOleCompound(sPath);
                    default:
                        return null; // no integrity check for other formats
                }
            } catch (Exception ex) {
                return "Integrity check failed: " + ex.Message;
            }
        }

        private static string checkOpenXml(string sPath, string sExtLower)
        {
            try {
                using (var zipArchive = ZipFile.OpenRead(sPath)) {
                    // All OOXML files must contain [Content_Types].xml.
                    bool bHasContentTypes = false;
                    foreach (var zipEntry in zipArchive.Entries) {
                        if (string.Equals(zipEntry.FullName, "[Content_Types].xml",
                            StringComparison.OrdinalIgnoreCase)) {
                            bHasContentTypes = true;
                            break;
                        }
                    }
                    if (!bHasContentTypes)
                        return "File appears corrupted: Open XML structure missing [Content_Types].xml. " +
                               "Try opening it manually in Office to see if it can be repaired.";
                }
                return null;
            } catch (InvalidDataException) {
                return "File appears corrupted: not a valid Open XML (ZIP) archive. " +
                       "Try opening it manually in Office to see if it can be repaired.";
            }
        }

        private static string checkOleCompound(string sPath)
        {
            try {
                byte[] binHead = new byte[8];
                using (var fileStream = File.OpenRead(sPath)) {
                    int iRead = fileStream.Read(binHead, 0, 8);
                    if (iRead < 8)
                        return "File appears corrupted: too short to be a valid Office document.";
                }
                for (int i = 0; i < 8; i++) {
                    if (binHead[i] != binOleMagic[i])
                        return "File appears corrupted: missing OLE Compound File signature. " +
                               "Try opening it manually in Office to see if it can be repaired.";
                }
                return null;
            } catch (Exception ex) {
                return "Could not read file for integrity check: " + ex.Message;
            }
        }
    }

    // -----------------------------------------------------------------
    // Per-run temp directory with crash-safe startup sweep.
    // -----------------------------------------------------------------
    public static class tempManager
    {
        public const string sTempPrefix = "2htm_";
        public const string sProcessName = "2htm";
        private static readonly Regex regexDir =
            new Regex(@"^2htm_(?<pid>\d+)_[0-9a-fA-F]{32}$", RegexOptions.Compiled);

        public static string newRunDir()
        {
            int iPid = Process.GetCurrentProcess().Id;
            string sGuid = Guid.NewGuid().ToString("N");
            string sDir = Path.Combine(Path.GetTempPath(), sTempPrefix + iPid + "_" + sGuid);
            Directory.CreateDirectory(sDir);
            return sDir;
        }

        public static void tryDelete(string sDir)
        {
            if (string.IsNullOrEmpty(sDir)) return;
            try { if (Directory.Exists(sDir)) Directory.Delete(sDir, true); }
            catch { }
        }

        public static void sweepStale()
        {
            try {
                string sTemp = Path.GetTempPath();
                if (!Directory.Exists(sTemp)) return;
                foreach (string sDir in Directory.GetDirectories(sTemp, sTempPrefix + "*")) {
                    try {
                        Match match = regexDir.Match(Path.GetFileName(sDir));
                        if (!match.Success) continue;
                        int iPid;
                        if (!int.TryParse(match.Groups["pid"].Value, out iPid)) continue;
                        if (isOurProcessAlive(iPid)) continue;
                        Directory.Delete(sDir, true);
                    } catch { }
                }
            } catch { }
        }

        private static bool isOurProcessAlive(int iPid)
        {
            try {
                Process process = Process.GetProcessById(iPid);
                try {
                    return string.Equals(process.ProcessName, sProcessName,
                        StringComparison.OrdinalIgnoreCase);
                } finally { process.Dispose(); }
            } catch (ArgumentException) { return false; }
            catch { return true; }
        }
    }

    // -----------------------------------------------------------------
    // HTML generation helpers.
    // -----------------------------------------------------------------
    public static class htmlWriter
    {
        public const string sDefaultDocType = "<!DOCTYPE html>";
        public const string sDefaultCharset = "utf-8";
        public const string sGeneratorName = "2htm";

        public const string sBaseCss =
            "html{font-size:100%}" +
            "body{font-family:sans-serif;max-width:80ch;margin:1em auto;padding:0 1em;" +
                "line-height:1.6;color:#000;text-align:left}" +
            "p{margin:1em 0}" +
            "h1,h2,h3,h4,h5,h6{line-height:1.3;margin-top:1.5em}" +
            "a{color:#00529b;text-decoration:underline}" +
            "a:hover,a:focus{text-decoration:underline;text-decoration-thickness:2px}" +
            "*:focus{outline:3px solid #00529b;outline-offset:2px}" +
            ".visually-hidden{position:absolute !important;clip:rect(1px,1px,1px,1px);" +
                "width:1px;height:1px;overflow:hidden;white-space:nowrap}" +
            "table{border-collapse:collapse;margin:1em 0}" +
            "th,td{border:1px solid #595959;padding:.4em .6em;text-align:left;vertical-align:top}" +
            "th{background:#e8e8e8}" +
            "td.num,th.num{text-align:right;font-variant-numeric:tabular-nums}" +
            "caption{caption-side:top;font-weight:bold;text-align:left;padding:.25em 0;font-size:1.05em}" +
            "nav ol{list-style:decimal;padding-left:2em}" +
            "figure{margin:1em 0}" +
            "img{max-width:100%;height:auto}" +
            "aside.slide-notes{border-left:4px solid #595959;padding-left:1em;margin:1em 0}" +
            ".byline{font-style:italic;color:#333;margin:0 0 1em 0}" +
            "p.subtitle{font-style:italic;color:#333;margin:0 0 1em 0}" +
            "pre{background:#f4f4f4;border:1px solid #595959;padding:.5em;overflow-x:auto;" +
                "font-family:monospace;white-space:pre-wrap;word-break:break-word}" +
            "code{font-family:monospace;background:#f4f4f4;padding:0 .2em;border:1px solid #d0d0d0}" +
            "pre code{background:transparent;border:0;padding:0}" +
            ".json-value{font-family:monospace}" +
            ".json-value.num{color:#004d00}" +
            ".json-value.str::before{content:'\"'}" +
            ".json-value.str::after{content:'\"'}" +
            ".json-value.null,.json-value.bool{font-style:italic}" +
            "dl.json-obj{margin-left:1.5em;border-left:2px solid #d0d0d0;padding-left:1em}" +
            "ol.json-arr{margin-left:1.5em}";

        public static string escape(string sIn)
        {
            if (string.IsNullOrEmpty(sIn)) return "";
            var sb = new StringBuilder(sIn.Length);
            foreach (char c in sIn) {
                switch (c) {
                    case '&': sb.Append("&amp;"); break;
                    case '<': sb.Append("&lt;"); break;
                    case '>': sb.Append("&gt;"); break;
                    case '"': sb.Append("&quot;"); break;
                    default:  sb.Append(c); break;
                }
            }
            return sb.ToString();
        }

        public static void writeHead(TextWriter writer, Dictionary<string, string> dMeta,
            string sDocTitle = null)
        {
            string sLang = dMeta.TryGetValue("language", out var sL) && !string.IsNullOrWhiteSpace(sL)
                ? sL : "en";
            string sMetaTitle = dMeta.TryGetValue("title", out var sT) && !string.IsNullOrWhiteSpace(sT)
                ? sT : "Converted document";

            // If a subtitle is present, combine it into the visible
            // title as "Title: Subtitle" (the conventional book
            // style). This affects both the <title> element in the
            // head and the <h1> emitted below. If an explicit
            // sDocTitle was supplied by the caller (e.g., the
            // Word/Excel converters), it overrides the derived title.
            if (dMeta.TryGetValue("subtitle", out var sSub) && !string.IsNullOrWhiteSpace(sSub)) {
                sMetaTitle = sMetaTitle + ": " + sSub;
            }
            if (string.IsNullOrWhiteSpace(sDocTitle)) sDocTitle = sMetaTitle;

            writer.WriteLine(sDefaultDocType);
            writer.WriteLine("<html lang=\"" + escape(sLang) + "\">");
            writer.WriteLine("<head>");
            writer.WriteLine("<meta charset=\"" + sDefaultCharset + "\">");
            writer.WriteLine("<meta name=\"viewport\" content=\"width=device-width,initial-scale=1\">");
            writer.WriteLine("<meta name=\"generator\" content=\"" + escape(sGeneratorName) + "\">");
            writer.WriteLine("<title>" + escape(sMetaTitle) + "</title>");

            writeMetaTag(writer, "author", dMeta);
            writeMetaTag(writer, "description", dMeta);
            writeMetaTag(writer, "keywords", dMeta);

            writeDc(writer, "DC.title", dMeta, "title");
            writeDc(writer, "DC.creator", dMeta, "author");
            writeDc(writer, "DC.subject", dMeta, "subject");
            writeDc(writer, "DC.description", dMeta, "description");
            writeDc(writer, "DC.date.created", dMeta, "created");
            writeDc(writer, "DC.date.modified", dMeta, "modified");
            writeDc(writer, "DC.language", dMeta, "language");

            writer.WriteLine("<style>" + sBaseCss + "</style>");
            writer.WriteLine("</head>");
            writer.WriteLine("<body>");

            writer.WriteLine("<h1 id=\"doc-title\">" + escape(sDocTitle) + "</h1>");
            writeMetaHeader(writer, dMeta);
        }

        private static void writeMetaHeader(TextWriter writer, Dictionary<string, string> dMeta)
        {
            // Author and Created date render as a visible byline
            // paragraph ("By Author \u2014 Date"). This reads naturally
            // as an opening line and is what most documents want.
            // The byline fields are NOT repeated in any subsequent
            // properties block.
            bool bAuthor  = hasVal(dMeta, "author");
            bool bCreated = hasVal(dMeta, "created");
            bool bByline  = bAuthor || bCreated;

            if (bByline) writeByline(writer, dMeta, bAuthor, bCreated);

            // The remaining document properties (subject, keywords,
            // description, last-modified date) are rendered as plain
            // paragraphs, one per property, in natural English rather
            // than wrapped in a labeled "Document properties" section.
            // A reader will see "Subject: ...", "Last modified: ...",
            // etc., which is readable by itself and does not require
            // the extra structural chrome.
            writeMetaParagraph(writer, dMeta, "subject",     "Subject");
            writeMetaParagraph(writer, dMeta, "description", "Description");
            writeMetaParagraph(writer, dMeta, "keywords",    "Keywords");
            writeMetaDateParagraph(writer, dMeta, "modified", "Last modified");
        }

        private static void writeMetaParagraph(TextWriter writer, Dictionary<string, string> d,
            string sKey, string sLabel)
        {
            if (!hasVal(d, sKey)) return;
            writer.WriteLine("<p>" + escape(sLabel) + ": " + escape(d[sKey]) + ".</p>");
        }

        private static void writeMetaDateParagraph(TextWriter writer, Dictionary<string, string> d,
            string sKey, string sLabel)
        {
            if (!hasVal(d, sKey)) return;
            string sIso = d[sKey];
            string sDisplay = formatIsoDateForDisplay(sIso);
            writer.WriteLine("<p>" + escape(sLabel) + ": " +
                "<time datetime=\"" + escape(sIso) + "\">" + escape(sDisplay) + "</time>.</p>");
        }

        // Converts an ISO 8601 datetime (e.g., "2022-05-19T08:49:58")
        // to a human-friendly date ("2022-05-19"). The datetime= attr
        // on the <time> element still gets the full ISO string for
        // machine consumption; this helper governs only what's read
        // aloud or displayed to a human. If parsing fails, returns
        // the input unchanged.
        public static string formatIsoDateForDisplay(string sIso)
        {
            if (string.IsNullOrWhiteSpace(sIso)) return sIso;
            // If the string has a 'T' (ISO datetime), take just the
            // date portion. Otherwise return as-is.
            int iT = sIso.IndexOf('T');
            return (iT > 0) ? sIso.Substring(0, iT) : sIso;
        }

        // Emits a byline paragraph combining author and date.
        private static void writeByline(TextWriter writer, Dictionary<string, string> dMeta,
            bool bAuthor, bool bCreated)
        {
            var sb = new StringBuilder();
            sb.Append("<p class=\"byline\">");
            if (bAuthor) {
                sb.Append("By ");
                sb.Append(escape(dMeta["author"]));
                if (bCreated) sb.Append(" \u2014 ");
            }
            if (bCreated) {
                string sIso = dMeta["created"];
                string sDisplay = formatIsoDateForDisplay(sIso);
                sb.Append("<time datetime=\"").Append(escape(sIso)).Append("\">");
                sb.Append(escape(sDisplay));
                sb.Append("</time>");
            }
            sb.Append("</p>");
            writer.WriteLine(sb.ToString());
        }

        private static bool hasVal(Dictionary<string, string> d, string sKey) =>
            d.TryGetValue(sKey, out var sV) && !string.IsNullOrWhiteSpace(sV);

        public static void writeFoot(TextWriter writer)
        {
            writer.WriteLine("</body>");
            writer.WriteLine("</html>");
        }

        private static void writeMetaTag(TextWriter writer, string sName, Dictionary<string, string> dMeta)
        {
            if (dMeta.TryGetValue(sName, out var sVal) && !string.IsNullOrWhiteSpace(sVal))
                writer.WriteLine("<meta name=\"" + sName + "\" content=\"" + escape(sVal) + "\">");
        }

        private static void writeDc(TextWriter writer, string sDcName, Dictionary<string, string> dMeta,
            string sKey)
        {
            if (dMeta.TryGetValue(sKey, out var sVal) && !string.IsNullOrWhiteSpace(sVal))
                writer.WriteLine("<meta name=\"" + sDcName + "\" content=\"" + escape(sVal) + "\">");
        }

        public static string shiftAndClampHeadings(string sHtmlFragment, int iTargetMin)
        {
            if (string.IsNullOrEmpty(sHtmlFragment)) return sHtmlFragment;

            var regexHead = new Regex(@"<(?<slash>/?)h(?<level>[1-6])(?<rest>[^>]*)>",
                RegexOptions.IgnoreCase);
            var matches = regexHead.Matches(sHtmlFragment);
            if (matches.Count == 0) return sHtmlFragment;

            int iMin = 6;
            foreach (Match m in matches) {
                int iLvl = int.Parse(m.Groups["level"].Value);
                if (iLvl < iMin) iMin = iLvl;
            }

            int iShift = iTargetMin - iMin;

            return regexHead.Replace(sHtmlFragment, m => {
                int iLvl = int.Parse(m.Groups["level"].Value);
                int iNew = iLvl + iShift;
                if (iNew < iTargetMin) iNew = iTargetMin;
                if (iNew > 6) iNew = 6;
                return "<" + m.Groups["slash"].Value + "h" + iNew + m.Groups["rest"].Value + ">";
            });
        }

        // Handles <img> tags in a body HTML fragment. Word's filtered
        // HTML writes images to a sibling folder and references them
        // by relative path; this method either embeds those images as
        // base64 data URLs (portable, self-contained) or strips the
        // tags entirely (when --strip-images is active, or when a
        // referenced file is missing or is an unsupported format).
        //
        // The base directory for resolving relative paths is sBaseDir
        // (typically the temp folder where Word wrote its output).
        // The method never throws on per-image failures; broken
        // references are silently removed because the user asked for
        // no broken links under any circumstance.
        public static string inlineOrStripImages(string sHtmlFragment, string sBaseDir)
        {
            if (string.IsNullOrEmpty(sHtmlFragment)) return sHtmlFragment;

            var regexImg = new Regex(@"<img\b[^>]*>", RegexOptions.IgnoreCase);
            var regexSrc = new Regex(@"\bsrc\s*=\s*(?:""([^""]*)""|'([^']*)')",
                RegexOptions.IgnoreCase);

            return regexImg.Replace(sHtmlFragment, m => {
                string sTag = m.Value;

                // When --strip-images is active, strip every <img> tag
                // regardless of whether its file exists.
                if (program.bStripImages) return "";

                var matchSrc = regexSrc.Match(sTag);
                if (!matchSrc.Success) return ""; // no src at all
                string sRel = matchSrc.Groups[1].Success
                    ? matchSrc.Groups[1].Value
                    : matchSrc.Groups[2].Value;
                if (string.IsNullOrWhiteSpace(sRel)) return "";

                // If the src is already a data URL, keep the tag as-is.
                if (sRel.StartsWith("data:", StringComparison.OrdinalIgnoreCase))
                    return sTag;

                // If the src is absolute (http:, https:, file:, etc.),
                // strip rather than embed. We do not fetch external
                // resources; a document that relies on them would
                // produce broken output in offline scenarios.
                if (Regex.IsMatch(sRel, @"^[a-zA-Z][a-zA-Z0-9+\-.]*:"))
                    return "";

                // Decode URL-encoded characters that Word may have
                // produced (spaces as %20, etc.).
                string sDecoded;
                try { sDecoded = Uri.UnescapeDataString(sRel); }
                catch { return ""; }

                string sFull = Path.Combine(sBaseDir, sDecoded);
                if (!File.Exists(sFull)) return "";

                string sMime = mimeForExtension(Path.GetExtension(sFull));
                if (sMime == null) return ""; // unsupported format

                byte[] binBytes;
                try { binBytes = File.ReadAllBytes(sFull); }
                catch { return ""; }

                string sB64 = Convert.ToBase64String(binBytes);
                string sNewSrc = "data:" + sMime + ";base64," + sB64;

                // Replace just the src attribute value; preserve
                // everything else (alt, width, height, etc.).
                string sResult = regexSrc.Replace(sTag,
                    mm => {
                        string sQuote = mm.Value.Contains("\"") ? "\"" : "'";
                        return "src=" + sQuote + sNewSrc + sQuote;
                    },
                    1);
                return sResult;
            });
        }

        // Maps a file extension (including the leading dot) to an
        // IANA MIME type suitable for a data URL. Returns null for
        // anything not in the whitelist; unknown types get stripped
        // rather than embedded with a guessed type.
        private static string mimeForExtension(string sExt)
        {
            if (string.IsNullOrEmpty(sExt)) return null;
            switch (sExt.ToLowerInvariant()) {
                case ".png":  return "image/png";
                case ".jpg":
                case ".jpeg": return "image/jpeg";
                case ".gif":  return "image/gif";
                case ".bmp":  return "image/bmp";
                case ".svg":  return "image/svg+xml";
                case ".webp": return "image/webp";
                default:      return null;
            }
        }

        // Writes a 2-D value array as an HTML <table>, with header
        // markup appropriate to the detected structure.
        //
        //   bColHeaders  first row is a header row; those cells get
        //                <th scope="col">
        //   bRowHeaders  first column is a header column; those cells
        //                get <th scope="row">
        //
        // When both flags are true and the top-left cell is a corner
        // (row 1, column 1), the corner gets <th> with no scope (the
        // WCAG-recommended pattern for a corner cell between both
        // header axes).
        //
        // When neither flag is true, writeTable is the wrong tool —
        // callers should emit paragraphs instead. We still produce a
        // valid table in that case (role="presentation" so screen
        // readers do not apply tabular semantics to it).
        public static void writeTable(TextWriter writer, object[,] aValues,
            string sCaption, string sSummaryId, string sSummaryText,
            bool bColHeaders, bool bRowHeaders)
        {
            int iR0 = aValues.GetLowerBound(0);
            int iC0 = aValues.GetLowerBound(1);
            int iRows = aValues.GetLength(0);
            int iCols = aValues.GetLength(1);

            bool[] abColIsNumeric = new bool[iCols];
            int iBodyStart = bColHeaders ? 1 : 0;
            for (int c = 0; c < iCols; c++)
                abColIsNumeric[c] = isColumnNumeric(aValues, iRows, iR0, iC0, c, iBodyStart);

            if (!string.IsNullOrEmpty(sSummaryId)) {
                writer.WriteLine("<p id=\"" + sSummaryId + "\" class=\"visually-hidden\">" +
                    escape(sSummaryText) + "</p>");
            }

            string sTableAttrs = string.IsNullOrEmpty(sSummaryId) ? "" :
                " aria-describedby=\"" + sSummaryId + "\"";
            if (!bColHeaders && !bRowHeaders)
                sTableAttrs += " role=\"presentation\"";
            writer.WriteLine("<table" + sTableAttrs + ">");
            if (!string.IsNullOrEmpty(sCaption))
                writer.WriteLine("<caption>" + escape(sCaption) + "</caption>");

            if (bColHeaders) {
                writer.WriteLine("<thead><tr>");
                for (int c = 0; c < iCols; c++) {
                    object oVal = aValues[iR0, iC0 + c];
                    string sCls = abColIsNumeric[c] ? " class=\"num\"" : "";
                    // If both axes have headers, the top-left cell is
                    // a corner between them; WCAG recommends emitting
                    // it as a bare <th> with no scope.
                    if (c == 0 && bRowHeaders)
                        writer.WriteLine("<th" + sCls + ">" + cellToHtml(oVal) + "</th>");
                    else
                        writer.WriteLine("<th scope=\"col\"" + sCls + ">" + cellToHtml(oVal) + "</th>");
                }
                writer.WriteLine("</tr></thead>");
            }

            writer.WriteLine("<tbody>");
            for (int r = iBodyStart; r < iRows; r++) {
                writer.WriteLine("<tr>");
                for (int c = 0; c < iCols; c++) {
                    object oVal = aValues[iR0 + r, iC0 + c];
                    bool bNum = abColIsNumeric[c] && isNumeric(oVal);
                    string sCls = bNum ? " class=\"num\"" : "";
                    if (c == 0 && bRowHeaders)
                        writer.WriteLine("<th scope=\"row\"" + sCls + ">" + cellToHtml(oVal) + "</th>");
                    else
                        writer.WriteLine("<td" + sCls + ">" + cellToHtml(oVal) + "</td>");
                }
                writer.WriteLine("</tr>");
            }
            writer.WriteLine("</tbody>");
            writer.WriteLine("</table>");
        }

        private static bool isColumnNumeric(object[,] aValues, int iRows, int iR0, int iC0,
            int iCol, int iBodyStart)
        {
            int iNum = 0, iNonEmpty = 0;
            for (int r = iBodyStart; r < iRows; r++) {
                object oVal = aValues[iR0 + r, iC0 + iCol];
                if (oVal == null) continue;
                iNonEmpty++;
                if (isNumeric(oVal)) iNum++;
            }
            if (iNonEmpty == 0) return false;
            return iNum * 2 >= iNonEmpty;
        }

        public static bool isNumeric(object oVal)
        {
            if (oVal == null) return false;
            if (oVal is double || oVal is int || oVal is long ||
                oVal is decimal || oVal is float || oVal is short) return true;
            if (oVal is string s)
                return double.TryParse(s, NumberStyles.Any, CultureInfo.InvariantCulture, out _);
            return false;
        }

        public static string cellToHtml(object oVal)
        {
            if (oVal == null) return "";
            if (oVal is DateTime dt) {
                string sIso = dt.ToString("yyyy-MM-dd", CultureInfo.InvariantCulture);
                return "<time datetime=\"" + sIso + "\">" + escape(sIso) + "</time>";
            }
            if (oVal is double nD)
                return escape(nD.ToString("R", CultureInfo.InvariantCulture));
            return escape(Convert.ToString(oVal, CultureInfo.InvariantCulture));
        }
    }

    // -----------------------------------------------------------------
    // Late-binding COM helpers.
    // -----------------------------------------------------------------
    public static class comHelper
    {
        public const int iHrRpcCallFailed = unchecked((int)0x800706BE);
        public const int iHrRpcServerUnavailable = unchecked((int)0x800706BA);
        public const int iHrRpcDisconnected = unchecked((int)0x80010108);
        public const int iHrObjectNotConnected = unchecked((int)0x80010114);

        public static bool isRpcFailure(Exception ex)
        {
            // Walk the inner-exception chain: comHelper.op wraps COM
            // failures in InvalidOperationException for diagnostic
            // labelling, so the original COMException may be nested
            // one or more levels deep. The retry logic must recognize
            // transient RPC failures regardless of wrapping.
            for (Exception cur = ex; cur != null; cur = cur.InnerException) {
                int h = cur is COMException ce ? ce.HResult : Marshal.GetHRForException(cur);
                if (h == iHrRpcCallFailed || h == iHrRpcServerUnavailable ||
                    h == iHrRpcDisconnected || h == iHrObjectNotConnected)
                    return true;
            }
            return false;
        }

        // Wraps a COM operation with a descriptive label so that if it
        // fails, the user sees WHICH operation failed, not just the
        // bare HRESULT.
        public static T op<T>(string sLabel, Func<T> fn)
        {
            try { return fn(); }
            catch (Exception ex) {
                throw new InvalidOperationException(sLabel + " failed: " + ex.Message, ex);
            }
        }

        public static void op(string sLabel, Action fn)
        {
            try { fn(); }
            catch (Exception ex) {
                throw new InvalidOperationException(sLabel + " failed: " + ex.Message, ex);
            }
        }

        public static dynamic createApp(string sProgId)
        {
            Type type = Type.GetTypeFromProgID(sProgId);
            if (type == null)
                throw new InvalidOperationException(
                    "Office component '" + sProgId + "' is not installed or not registered. " +
                    "This file cannot be converted without it.");
            dynamic oApp = Activator.CreateInstance(type);
            silenceAlerts(sProgId, oApp);
            return oApp;
        }

        /// <summary>
        /// Silence interactive alerts on a freshly-created Office
        /// Application object so unattended automation does not stall
        /// on a dialog. This handles the common Word/Excel/PowerPoint
        /// alert sources, including:
        ///
        ///   Word
        ///     - DisplayAlerts = wdAlertsNone (0): no on-screen alerts.
        ///     - Options.ConfirmConversions = false: no encoding prompt
        ///       when opening text files of unknown encoding.
        ///     - Options.DoNotPromptForConvert = true: no "Word will
        ///       convert this PDF to an editable Word document" prompt
        ///       (this was the dialog that previously locked up 2htm
        ///       when converting a PDF). Available since Word 2013.
        ///     - AutomationSecurity = msoAutomationSecurityForceDisable
        ///       (3): macros in opened files are blocked silently
        ///       instead of prompting.
        ///     - Visible = false: window is hidden.
        ///
        ///   Excel
        ///     - DisplayAlerts = false
        ///     - AskToUpdateLinks = false: no "update external links"
        ///       prompt when opening a workbook with external refs.
        ///     - AlertBeforeOverwriting = false
        ///     - AutomationSecurity = msoAutomationSecurityForceDisable
        ///     - Visible = false
        ///
        ///   PowerPoint
        ///     - DisplayAlerts = ppAlertsNone (1)
        ///     - AutomationSecurity = msoAutomationSecurityForceDisable
        ///     - PowerPoint cannot run with Visible = false; the window
        ///       is left visible (Microsoft documents this limitation).
        ///
        /// Each setter is wrapped in its own try/catch because not all
        /// versions of Office expose every property, and some of the
        /// settings throw harmlessly when set on certain editions.
        /// </summary>
        private static void silenceAlerts(string sProgId, dynamic oApp)
        {
            // Office .mso constants. Hard-coded to avoid pulling in
            // an Office interop reference at compile time.
            const int iMsoAutomationSecurityForceDisable = 3;
            const int iWdAlertsNone = 0;
            const int iPpAlertsNone = 1;

            string sLowered = (sProgId ?? "").ToLowerInvariant();
            try { oApp.AutomationSecurity = iMsoAutomationSecurityForceDisable; } catch { }

            if (sLowered.StartsWith("word.")) {
                try { oApp.DisplayAlerts = iWdAlertsNone; } catch { }
                try { oApp.Visible = false; } catch { }
                try { oApp.Options.ConfirmConversions = false; } catch { }
                try { oApp.Options.DoNotPromptForConvert = true; } catch { }
                try { oApp.Options.SaveNormalPrompt = false; } catch { }
                try { oApp.Options.WarnBeforeSavingPrintingSendingMarkup = false; } catch { }
                try { oApp.Options.UpdateLinksAtOpen = false; } catch { }
                try { oApp.Options.CheckGrammarAsYouType = false; } catch { }
                try { oApp.Options.CheckSpellingAsYouType = false; } catch { }
            }
            else if (sLowered.StartsWith("excel.")) {
                try { oApp.DisplayAlerts = false; } catch { }
                try { oApp.Visible = false; } catch { }
                try { oApp.AskToUpdateLinks = false; } catch { }
                try { oApp.AlertBeforeOverwriting = false; } catch { }
                try { oApp.ScreenUpdating = false; } catch { }
                try { oApp.EnableEvents = false; } catch { }
            }
            else if (sLowered.StartsWith("powerpoint.")) {
                try { oApp.DisplayAlerts = iPpAlertsNone; } catch { }
                // PowerPoint requires a visible window to render slides,
                // so we do NOT set Visible = false. Leave it as-is.
            }
        }

        public static void release(object oObj)
        {
            if (oObj == null) return;
            try { if (Marshal.IsComObject(oObj)) Marshal.ReleaseComObject(oObj); }
            catch { }
        }

        public static void killOrphanOfficeProcesses(string sProcessName)
        {
            try {
                foreach (Process process in Process.GetProcessesByName(sProcessName)) {
                    try {
                        if (process.MainWindowHandle == IntPtr.Zero) {
                            process.Kill();
                            process.WaitForExit(3000);
                        }
                    } catch { }
                    finally { process.Dispose(); }
                }
            } catch { }
        }

        public static Dictionary<string, string> readBuiltInDocProps(dynamic oDoc)
        {
            var dRet = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);
            string[] asNames = {
                "Title", "Author", "Subject", "Keywords", "Comments",
                "Last author", "Creation date", "Last save time"
            };
            dynamic oProps = null;
            try { oProps = oDoc.BuiltinDocumentProperties; } catch { return dRet; }

            foreach (var sName in asNames) {
                try {
                    dynamic oProp = oProps[sName];
                    object oVal = oProp.Value;
                    if (oVal != null) {
                        string sKey = mapPropName(sName);
                        dRet[sKey] = formatPropValue(oVal);
                    }
                } catch { }
            }
            return dRet;
        }

        private static string mapPropName(string sOfficeName)
        {
            switch (sOfficeName) {
                case "Title": return "title";
                case "Author": return "author";
                case "Subject": return "subject";
                case "Keywords": return "keywords";
                case "Comments": return "description";
                case "Last author": return "last_author";
                case "Creation date": return "created";
                case "Last save time": return "modified";
                default: return sOfficeName.ToLowerInvariant();
            }
        }

        private static string formatPropValue(object oVal)
        {
            if (oVal is DateTime dt) return dt.ToString("yyyy-MM-ddTHH:mm:ss", CultureInfo.InvariantCulture);
            return Convert.ToString(oVal, CultureInfo.InvariantCulture);
        }

        // Retry only for transient RPC failures. Deterministic failures
        // propagate after the first attempt. Returns the 1-based attempt
        // number that succeeded, so the caller can log it.
        public static int retryOfficeOp(Action action, string sProcessName, int iMaxAttempts = 3)
        {
            int iAttempt = 0;
            int iWaitMs = 1500;
            while (true) {
                iAttempt++;
                try { action(); return iAttempt; }
                catch (Exception ex) when (isRpcFailure(ex) && iAttempt < iMaxAttempts) {
                    // Retry is normal recovery; log only, no console
                    // chatter to keep per-file output concise.
                    logger.warn(sProcessName + " RPC failure on attempt " + iAttempt +
                        " of " + iMaxAttempts + ": " + ex.Message);
                    if (ex.InnerException != null)
                        logger.warn("  inner: " + ex.InnerException.GetType().FullName +
                            ": " + ex.InnerException.Message);
                    killOrphanOfficeProcesses(sProcessName);
                    Thread.Sleep(iWaitMs);
                    iWaitMs *= 2;
                }
            }
        }
    }

    // -----------------------------------------------------------------
    // Word.
    // -----------------------------------------------------------------
    public static class wordConverter
    {
        public const int iWdFormatFilteredHtml = 10;
        public const int iWdDoNotSaveChanges = 0;
        public const int iWdAlertsNone = 0;

        public static void convert(string sInPath, string sOutPath)
        {
            int iAttempt = comHelper.retryOfficeOp(
                () => convertAttempt(sInPath, sOutPath), "WINWORD", 3);
            if (iAttempt > 1) logger.info("Converted on attempt " + iAttempt + ": " + sOutPath);
        }

        private static void convertAttempt(string sInPath, string sOutPath)
        {
            dynamic oApp = null, oDoc = null;
            string sTempDir = tempManager.newRunDir();
            string sTempHtm = Path.Combine(sTempDir, "word.htm");
            bool bIsPdf = string.Equals(Path.GetExtension(sInPath), ".pdf",
                StringComparison.OrdinalIgnoreCase);
            bool bWritten = false;
            try {
                oApp = comHelper.op("Word: createApp", () => comHelper.createApp("Word.Application"));
                comHelper.op("Word: set Visible=false", () => { oApp.Visible = false; });
                comHelper.op("Word: set DisplayAlerts", () => { oApp.DisplayAlerts = iWdAlertsNone; });

                if (bIsPdf) {
                    try { oApp.Options.DoNotPromptForConvert = true; } catch { }
                    try { oApp.Options.ConfirmConversions = false; } catch { }
                    oDoc = openPdf(oApp, sInPath);
                } else {
                    oDoc = openOfficeDoc(oApp, sInPath);
                }

                var dMeta = comHelper.readBuiltInDocProps(oDoc);
                string sDocTitle = pickTitle(dMeta, sInPath);

                comHelper.op("Word: SaveAs2 filtered HTML",
                    () => { oDoc.SaveAs2(FileName: sTempHtm, FileFormat: iWdFormatFilteredHtml); });
                comHelper.op("Word: Close",
                    () => { oDoc.Close(iWdDoNotSaveChanges); });
                oDoc = null;

                if (!File.Exists(sTempHtm))
                    throw new InvalidOperationException(
                        "Word did not produce an HTML file. The source may be empty or unreadable.");

                string sBody = extractBody(File.ReadAllText(sTempHtm, Encoding.UTF8));
                sBody = htmlWriter.inlineOrStripImages(sBody, sTempDir);
                sBody = htmlWriter.shiftAndClampHeadings(sBody, 2);
                writeFinal(sOutPath, dMeta, sDocTitle, sBody);

                // writeFinal closed its StreamWriter cleanly: the
                // file is complete on disk. Teardown errors below
                // should not invalidate this success.
                bWritten = true;
            } finally {
                try { if (oDoc != null) { try { oDoc.Close(iWdDoNotSaveChanges); } catch { } comHelper.release(oDoc); } } catch { }
                try { if (oApp != null) { try { oApp.Quit(); } catch { } comHelper.release(oApp); } } catch { }
                try { tempManager.tryDelete(sTempDir); } catch { }
            }

            if (!bWritten) {
                try { if (File.Exists(sOutPath)) File.Delete(sOutPath); } catch { }
            }
        }

        private static dynamic openOfficeDoc(dynamic oApp, string sInPath)
        {
            return comHelper.op("Word: Documents.Open",
                () => (object)oApp.Documents.Open(
                    FileName: sInPath,
                    ReadOnly: true,
                    AddToRecentFiles: false,
                    Visible: false));
        }

        private static dynamic openPdf(dynamic oApp, string sInPath)
        {
            try {
                return oApp.Documents.Open(
                    FileName: sInPath,
                    ConfirmConversions: false,
                    ReadOnly: true,
                    AddToRecentFiles: false,
                    Visible: false);
            } catch (Exception ex) {
                throw new InvalidOperationException(
                    "Word could not open this PDF. Word 2013 or later is required " +
                    "for PDF Reflow. Inner error: " + ex.Message);
            }
        }

        private static string pickTitle(Dictionary<string, string> dMeta, string sInPath)
        {
            if (dMeta.TryGetValue("title", out var s) && !string.IsNullOrWhiteSpace(s)) return s;
            return Path.GetFileNameWithoutExtension(sInPath);
        }

        private static string extractBody(string sHtm)
        {
            var match = Regex.Match(sHtm, @"<body[^>]*>(?<b>.*?)</body>",
                RegexOptions.Singleline | RegexOptions.IgnoreCase);
            string sBody = match.Success ? match.Groups["b"].Value : sHtm;
            sBody = Regex.Replace(sBody, @"<style.*?</style>", "", RegexOptions.Singleline | RegexOptions.IgnoreCase);
            sBody = Regex.Replace(sBody, @"<script.*?</script>", "", RegexOptions.Singleline | RegexOptions.IgnoreCase);
            sBody = Regex.Replace(sBody, @"<o:p>.*?</o:p>", "", RegexOptions.Singleline | RegexOptions.IgnoreCase);
            sBody = Regex.Replace(sBody, @"<xml>.*?</xml>", "", RegexOptions.Singleline | RegexOptions.IgnoreCase);
            sBody = Regex.Replace(sBody, @"\s+class=""Mso[^""]*""", "", RegexOptions.IgnoreCase);
            return sBody;
        }

        private static void writeFinal(string sOutPath, Dictionary<string, string> dMeta,
            string sDocTitle, string sBody)
        {
            using (var writer = new StreamWriter(sOutPath, false, new UTF8Encoding(false))) {
                htmlWriter.writeHead(writer, dMeta, sDocTitle);
                writer.WriteLine("<main aria-labelledby=\"doc-title\">");
                writer.WriteLine(sBody);
                writer.WriteLine("</main>");
                htmlWriter.writeFoot(writer);
            }
        }

        // Plain-text conversion path. Uses Word's SaveAs2 with
        // wdFormatUnicodeText (value 7), which writes a UTF-16
        // little-endian .txt file with BOM. We read that back in
        // and rewrite as UTF-8 without BOM so the final output
        // matches the encoding convention used by the HTML path.
        public const int iWdFormatUnicodeText = 7;

        public static void convertToText(string sInPath, string sOutPath)
        {
            int iAttempt = comHelper.retryOfficeOp(
                () => convertToTextAttempt(sInPath, sOutPath), "WINWORD", 3);
            if (iAttempt > 1) logger.info("Converted on attempt " + iAttempt + ": " + sOutPath);
        }

        private static void convertToTextAttempt(string sInPath, string sOutPath)
        {
            dynamic oApp = null, oDoc = null;
            string sTempDir = tempManager.newRunDir();
            string sTempTxt = Path.Combine(sTempDir, "word.txt");
            bool bIsPdf = string.Equals(Path.GetExtension(sInPath), ".pdf",
                StringComparison.OrdinalIgnoreCase);
            bool bWritten = false;
            try {
                oApp = comHelper.op("Word: createApp", () => comHelper.createApp("Word.Application"));
                comHelper.op("Word: set Visible=false", () => { oApp.Visible = false; });
                comHelper.op("Word: set DisplayAlerts", () => { oApp.DisplayAlerts = iWdAlertsNone; });

                if (bIsPdf) {
                    try { oApp.Options.DoNotPromptForConvert = true; } catch { }
                    try { oApp.Options.ConfirmConversions = false; } catch { }
                    oDoc = openPdf(oApp, sInPath);
                } else {
                    oDoc = openOfficeDoc(oApp, sInPath);
                }

                comHelper.op("Word: SaveAs2 Unicode text",
                    () => { oDoc.SaveAs2(FileName: sTempTxt, FileFormat: iWdFormatUnicodeText); });
                comHelper.op("Word: Close",
                    () => { oDoc.Close(iWdDoNotSaveChanges); });
                oDoc = null;

                if (!File.Exists(sTempTxt))
                    throw new InvalidOperationException(
                        "Word did not produce a text file. The source may be empty or unreadable.");

                // Word writes UTF-16 LE with BOM. Read as such, then
                // rewrite as UTF-8 with BOM to match the text-output
                // encoding convention used across this tool.
                string sContent = File.ReadAllText(sTempTxt, Encoding.Unicode);
                File.WriteAllText(sOutPath, sContent, new UTF8Encoding(true));
                bWritten = true;
            } finally {
                try { if (oDoc != null) { try { oDoc.Close(iWdDoNotSaveChanges); } catch { } comHelper.release(oDoc); } } catch { }
                try { if (oApp != null) { try { oApp.Quit(); } catch { } comHelper.release(oApp); } } catch { }
                try { tempManager.tryDelete(sTempDir); } catch { }
            }

            if (!bWritten) {
                try { if (File.Exists(sOutPath)) File.Delete(sOutPath); } catch { }
            }
        }
    }

    // -----------------------------------------------------------------
    // Excel. Every COM call labelled for diagnostic output on failure.
    // -----------------------------------------------------------------
    public static class excelConverter
    {
        public static void convert(string sInPath, string sOutPath)
        {
            int iAttempt = comHelper.retryOfficeOp(
                () => convertAttempt(sInPath, sOutPath), "EXCEL", 3);
            if (iAttempt > 1) logger.info("Converted on attempt " + iAttempt + ": " + sOutPath);
        }

        private static void convertAttempt(string sInPath, string sOutPath)
        {
            dynamic oApp = null, oWb = null;
            bool bWritten = false;
            try {
                oApp = comHelper.op("Excel: createApp", () => comHelper.createApp("Excel.Application"));
                comHelper.op("Excel: set Visible=false", () => { oApp.Visible = false; });
                comHelper.op("Excel: set DisplayAlerts=false", () => { oApp.DisplayAlerts = false; });
                comHelper.op("Excel: set ScreenUpdating=false", () => { oApp.ScreenUpdating = false; });
                try { oApp.AskToUpdateLinks = false; } catch { }
                try { oApp.EnableEvents = false; } catch { }

                oWb = comHelper.op("Excel: Workbooks.Open",
                    () => (object)oApp.Workbooks.Open(
                        Filename: sInPath,
                        ReadOnly: true,
                        UpdateLinks: 0,
                        AddToMru: false));

                var dMeta = comHelper.readBuiltInDocProps(oWb);
                string sDocTitle = pickTitle(dMeta, sInPath);

                using (var writer = new StreamWriter(sOutPath, false, new UTF8Encoding(false))) {
                    htmlWriter.writeHead(writer, dMeta, sDocTitle);
                    // No synthetic navigation: Excel workbooks have
                    // no notion of a table-of-contents in the source.
                    // Sheets are accessed naturally via the <h2>
                    // section headings below.

                    writer.WriteLine("<main aria-labelledby=\"doc-title\">");
                    int iCount = comHelper.op("Excel: Worksheets.Count",
                        () => (int)oWb.Worksheets.Count);
                    for (int i = 1; i <= iCount; i++) {
                        dynamic oSheet = null;
                        try {
                            int iLocal = i;
                            oSheet = comHelper.op("Excel: Worksheets[" + iLocal + "]",
                                () => (object)oWb.Worksheets[iLocal]);
                            writeSheet(writer, oSheet, i);
                        } finally {
                            comHelper.release(oSheet);
                        }
                    }
                    writer.WriteLine("</main>");

                    htmlWriter.writeFoot(writer);
                }
                bWritten = true;
            } finally {
                try { if (oWb != null) { try { oWb.Close(false); } catch { } comHelper.release(oWb); } } catch { }
                try { if (oApp != null) { try { oApp.Quit(); } catch { } comHelper.release(oApp); } } catch { }
            }

            if (!bWritten) {
                try { if (File.Exists(sOutPath)) File.Delete(sOutPath); } catch { }
            }
        }

        private static string pickTitle(Dictionary<string, string> dMeta, string sInPath)
        {
            if (dMeta.TryGetValue("title", out var s) && !string.IsNullOrWhiteSpace(s)) return s;
            return Path.GetFileNameWithoutExtension(sInPath);
        }

        // Describes a rectangular region of cells on a sheet,
        // together with the header structure detected for that
        // region and the 2-D value array containing its contents.
        private class sheetRegion
        {
            public int iRow;            // 1-indexed top row on sheet
            public int iCol;            // 1-indexed left column on sheet
            public int iRows;           // row count
            public int iCols;           // column count
            public object[,] aValues;   // values, (0,0) = top-left cell
            public bool bColHeaders;    // row 1 qualifies as column headers
            public bool bRowHeaders;    // col 1 qualifies as row headers
        }

        private static void writeSheet(TextWriter writer, dynamic oSheet, int iNum)
        {
            string sName = comHelper.op("Excel: Sheet[" + iNum + "].Name",
                () => Convert.ToString(oSheet.Name, CultureInfo.InvariantCulture));
            bool bHasName = !string.IsNullOrWhiteSpace(sName);
            string sHeading = bHasName
                ? "<span class=\"visually-hidden\">Sheet </span>" + iNum +
                  ": " + htmlWriter.escape(sName.Trim())
                : "Sheet " + iNum;
            writer.WriteLine("<section id=\"sheet-" + iNum + "\" aria-labelledby=\"sheet-" + iNum + "-h\">");
            writer.WriteLine("<h2 id=\"sheet-" + iNum + "-h\">" + sHeading + "</h2>");

            logger.info("Sheet " + iNum + " (" + (bHasName ? sName : "unnamed") + "): discovering regions");
            var lsRegions = discoverRegions(oSheet);
            logger.info("Sheet " + iNum + ": found " + lsRegions.Count + " region(s)");
            if (lsRegions.Count == 0) {
                writer.WriteLine("<p><em>(empty sheet)</em></p>");
                writer.WriteLine("</section>");
                return;
            }

            int iIdx = 0;
            foreach (var oRegion in lsRegions) {
                iIdx++;
                logger.info("Sheet " + iNum + " region " + iIdx +
                    ": " + oRegion.iRows + "x" + oRegion.iCols +
                    " at row " + oRegion.iRow + ", col " + oRegion.iCol +
                    " (colHeaders=" + oRegion.bColHeaders +
                    ", rowHeaders=" + oRegion.bRowHeaders + ")");
                writeRegion(writer, oRegion, sName, iNum, iIdx, lsRegions.Count);
            }
            writer.WriteLine("</section>");
        }

        // Discovers disjoint rectangular regions on the sheet, using
        // Excel's native Range.CurrentRegion detection. Returns the
        // regions sorted by top-left position (row first, then column).
        //
        // Algorithm (matches xlStruct.vbs conceptually):
        //   1. Walk every cell in UsedRange.
        //   2. For each non-empty cell, compute cell.CurrentRegion.
        //   3. Collect distinct regions (deduplicated by address).
        //   4. Drop regions fully enclosed within any larger region.
        //
        // Each surviving region has its 2-D value array fetched and
        // its header structure classified. The result is a list of
        // sheetRegion objects ready for rendering.
        private static List<sheetRegion> discoverRegions(dynamic oSheet)
        {
            var lsResult = new List<sheetRegion>();
            dynamic oUsed = null, oLastRegion = null;
            // Map region address ("$A$1:$D$14") to the Range object
            // so we can iterate unique regions later without calling
            // CurrentRegion on the same seed cell twice.
            // Maximum worksheet cell count for which we'll attempt
            // the standard UsedRange.Value2 fetch + full-array walk.
            //
            // Informed by the .NET Framework CLR's object-size limit
            // of 2 GB (default; gcAllowVeryLargeObjects is NOT
            // enabled in this build). An object[,] of 64-bit
            // references at 200 million elements is ~1.6 GB,
            // comfortably under the cap with room for the array
            // metadata and live VARIANTs that Excel marshals to us.
            //
            // Above this threshold we don't attempt to allocate the
            // full array. We use SpecialCells (constants + formulas)
            // to find areas that actually contain data, and seed
            // CurrentRegion from the top-left of each area. This
            // works correctly regardless of what UsedRange reports,
            // so it handles sheets with phantom UsedRange (stray
            // formatting on cells far from real data) as well as
            // sheets whose real used area happens to be large.
            //
            // The threshold is not an arbitrary performance switch:
            // below it the fast path is guaranteed to be allocatable;
            // above it the fast path is guaranteed to throw
            // OutOfMemoryException. Both paths produce equivalent
            // region output on any sheet where the fast path would
            // have succeeded.
            const long iMaxCellsForValueFetch = 200_000_000;

            var dRegions = new Dictionary<string, dynamic>(StringComparer.OrdinalIgnoreCase);
            try {
                oUsed = oSheet.UsedRange;
                int iRowsUsed = 0, iColsUsed = 0;
                try { iRowsUsed = (int)oUsed.Rows.Count; } catch (Exception ex) {
                    logger.error("  UsedRange.Rows.Count failed: " + ex.Message);
                }
                try { iColsUsed = (int)oUsed.Columns.Count; } catch (Exception ex) {
                    logger.error("  UsedRange.Columns.Count failed: " + ex.Message);
                }
                logger.info("  UsedRange: " + iRowsUsed + " rows x " + iColsUsed + " columns");
                if (iRowsUsed == 0 || iColsUsed == 0) {
                    logger.info("  UsedRange is empty; no regions to discover.");
                    return lsResult;
                }

                int iUsedRow0 = 1, iUsedCol0 = 1;
                try { iUsedRow0 = (int)oUsed.Row;    } catch { }
                try { iUsedCol0 = (int)oUsed.Column; } catch { }
                logger.info("  UsedRange origin: row " + iUsedRow0 + ", col " + iUsedCol0);

                long iCellCount = (long)iRowsUsed * iColsUsed;
                if (iCellCount <= iMaxCellsForValueFetch) {
                    seedRegionsFromValueWalk(oSheet, oUsed, iRowsUsed, iColsUsed,
                        iUsedRow0, iUsedCol0, dRegions);
                } else {
                    logger.info("  UsedRange has " + iCellCount +
                        " cells, above CLR-safe threshold of " +
                        iMaxCellsForValueFetch +
                        "; using SpecialCells seed path.");
                    seedRegionsFromSpecialCells(oSheet, dRegions);
                }

                logger.info("  Seed pass produced " + dRegions.Count + " distinct region(s).");

                // Drop regions enclosed in another (rare but the
                // xlStruct algorithm guards against it).
                var lsKeys = new List<string>(dRegions.Keys);
                var lsKept = new List<string>();
                foreach (var sInner in lsKeys) {
                    bool bEnclosed = false;
                    var oInner = dRegions[sInner];
                    int iInR = (int)oInner.Row, iInC = (int)oInner.Column;
                    int iInRN = (int)oInner.Rows.Count, iInCN = (int)oInner.Columns.Count;
                    foreach (var sOuter in lsKeys) {
                        if (sOuter.Equals(sInner, StringComparison.OrdinalIgnoreCase)) continue;
                        var oOuter = dRegions[sOuter];
                        int iOuR = (int)oOuter.Row, iOuC = (int)oOuter.Column;
                        int iOuRN = (int)oOuter.Rows.Count, iOuCN = (int)oOuter.Columns.Count;
                        if (iInR >= iOuR && iInR + iInRN <= iOuR + iOuRN &&
                            iInC >= iOuC && iInC + iInCN <= iOuC + iOuCN) {
                            bEnclosed = true;
                            break;
                        }
                    }
                    if (!bEnclosed) lsKept.Add(sInner);
                }
                if (lsKept.Count != dRegions.Count)
                    logger.info("  After enclosure filter: " + lsKept.Count +
                        " of " + dRegions.Count + " region(s) kept.");

                // Build sheetRegion records. Sort by top-left position.
                var lsBuilt = new List<sheetRegion>();
                foreach (var sAddr in lsKept) {
                    var oRegion = dRegions[sAddr];
                    var regionBuilt = buildRegion(oRegion);
                    if (regionBuilt != null) lsBuilt.Add(regionBuilt);
                }
                lsBuilt.Sort((a, b) => {
                    int d = a.iRow.CompareTo(b.iRow);
                    return d != 0 ? d : a.iCol.CompareTo(b.iCol);
                });
                lsResult.AddRange(lsBuilt);
            } catch (Exception ex) {
                logger.error("discoverRegions outer exception: " +
                    ex.GetType().FullName + ": " + ex.Message);
                if (ex.InnerException != null)
                    logger.error("  inner: " + ex.InnerException.GetType().FullName +
                        ": " + ex.InnerException.Message);
                logger.error("  stack: " + ex.StackTrace);
                throw;
            }
            finally {
                comHelper.release(oLastRegion);
                comHelper.release(oUsed);
                foreach (var o in dRegions.Values)
                    comHelper.release(o);
            }
            return lsResult;
        }

        // Seed discovery via an array-walk of UsedRange values.
        // Fetches UsedRange.Value2 as an object[,], then walks the
        // array; for each non-null cell, calls CurrentRegion and
        // records its address. Subsequent cells known to be in an
        // already-seen region are skipped (the hsCovered set).
        //
        // This is the faster and more thorough path for normal
        // sheets. It requires UsedRange to fit in a CLR object
        // array — caller is responsible for the threshold check.
        private static void seedRegionsFromValueWalk(dynamic oSheet, dynamic oUsed,
            int iRowsUsed, int iColsUsed, int iUsedRow0, int iUsedCol0,
            Dictionary<string, dynamic> dRegions)
        {
            object[,] aUsed;
            try {
                logger.info("  Fetching UsedRange.Value2...");
                object oRaw = oUsed.Value2;
                aUsed = normalizeToArray(oRaw, iRowsUsed, iColsUsed);
                logger.info("  Value2 fetched ok (" + aUsed.GetLength(0) + "x" +
                    aUsed.GetLength(1) + " array).");
            } catch (Exception ex) {
                logger.error("  UsedRange.Value2 failed: " +
                    ex.GetType().Name + ": " + ex.Message);
                if (ex.InnerException != null)
                    logger.error("    inner: " + ex.InnerException.GetType().Name +
                        ": " + ex.InnerException.Message);
                // Propagate rather than silently producing an
                // empty region list, which would mislead the user
                // into thinking the sheet was empty.
                throw;
            }

            int iR0 = aUsed.GetLowerBound(0);
            int iC0 = aUsed.GetLowerBound(1);

            var hsCovered = new HashSet<string>();
            for (int r = 0; r < iRowsUsed; r++) {
                for (int c = 0; c < iColsUsed; c++) {
                    object oVal = aUsed[iR0 + r, iC0 + c];
                    if (oVal == null) continue;
                    int iR = iUsedRow0 + r;
                    int iC = iUsedCol0 + c;
                    string sCell = "R" + iR + "C" + iC;
                    if (hsCovered.Contains(sCell)) continue;
                    seedRegionFromCell(oSheet, iR, iC, dRegions, hsCovered);
                }
            }
        }

        // Seed discovery via Excel's SpecialCells. Called when the
        // sheet's reported UsedRange is too large to fetch as a
        // full Value2 array (typically a phantom UsedRange caused
        // by stray formatting). SpecialCells queries Excel directly
        // for cells that actually contain content.
        //
        // We ask in two phases:
        //
        //   1. Constants: the literal data on the sheet. We record
        //      both the Area addresses and the maximum row any
        //      constants area occupies ("the data floor").
        //
        //   2. Formulas: cells containing formulas. For each
        //      formulas area, if its row range extends past the
        //      constants data floor, we clip it to that floor and
        //      seed the clipped range instead. This handles the
        //      common pattern of a formula filled down an entire
        //      column (T2:T1048576) on a sheet where real data
        //      only lives in the first N rows; without clipping we
        //      would render a million trailing fill-down rows.
        //
        // On a sheet with only formulas (no constants), the data
        // floor is 0 and no clipping is applied — formulas are
        // trusted to be the real content.
        //
        // Unlike the value-walk path, we do NOT call CurrentRegion
        // here. On a sheet with scattered phantom cells,
        // CurrentRegion can expand across the whole sheet, throwing
        // OutOfMemoryException and leaving the COM context in an
        // unrecoverable state. The Areas returned by SpecialCells
        // are themselves rectangles of contiguous same-type cells,
        // which makes them safe to treat as seed regions directly.
        //
        // Tradeoff: a data table that mixes constants and formulas
        // in an interleaved pattern, or that has blank cells
        // interior to its body, may be reported as multiple Areas
        // and rendered as several small tables instead of one
        // larger one. The fallback path accepts this tradeoff in
        // exchange for robustness on sheets whose structure is
        // already compromised.
        private static void seedRegionsFromSpecialCells(dynamic oSheet,
            Dictionary<string, dynamic> dRegions)
        {
            const int iXlCellTypeConstants = 2;
            const int iXlCellTypeFormulas = -4123;

            int iMaxConstantsRow = seedRegionsFromSpecialCellsType(
                oSheet, iXlCellTypeConstants, "constants", dRegions, iClipToRow: 0);
            // Pass max-constants-row as clip bound. When it's zero
            // (no constants on the sheet), the clip helper treats
            // it as "no clipping."
            seedRegionsFromSpecialCellsType(
                oSheet, iXlCellTypeFormulas, "formulas", dRegions,
                iClipToRow: iMaxConstantsRow);
        }

        // Walks one SpecialCells type's areas and seeds them into
        // dRegions. Returns the maximum row any processed area
        // actually occupies, so the caller can use it as a clip
        // bound for subsequent passes.
        //
        // iClipToRow: if > 0 and an area extends below that row,
        // clip it to end at iClipToRow. If == 0, no clipping is
        // applied.
        private static int seedRegionsFromSpecialCellsType(dynamic oSheet,
            int iCellType, string sLabel, Dictionary<string, dynamic> dRegions,
            int iClipToRow)
        {
            int iMaxRowSeen = 0;
            dynamic oAllCells = null, oSpec = null, oAreas = null;
            try {
                oAllCells = oSheet.Cells;
                try {
                    oSpec = oAllCells.SpecialCells(iCellType);
                } catch (Exception ex) {
                    // Excel raises "No cells were found" when the
                    // sheet has no cells of the requested type.
                    // Not an error for our purposes.
                    logger.info("  SpecialCells(" + sLabel + "): none found (" +
                        ex.Message.Trim() + ")");
                    return 0;
                }
                oAreas = oSpec.Areas;
                int iAreaCount = 0;
                try { iAreaCount = (int)oAreas.Count; } catch { }
                logger.info("  SpecialCells(" + sLabel + "): " + iAreaCount + " area(s)");

                for (int i = 1; i <= iAreaCount; i++) {
                    dynamic oArea = null, oClipped = null;
                    // Which Range (if either) was handed off to
                    // dRegions? Tracks which one the finally must
                    // NOT release.
                    bool bAreaTransferred = false;
                    bool bClippedTransferred = false;
                    try {
                        oArea = oAreas[i];
                        string sAddr = Convert.ToString(oArea.Address,
                            CultureInfo.InvariantCulture);
                        int iRows = (int)oArea.Rows.Count;
                        int iCols = (int)oArea.Columns.Count;
                        int iRow = (int)oArea.Row;
                        int iCol = (int)oArea.Column;
                        int iLastRow = iRow + iRows - 1;
                        logger.info("    area " + i + ": " + sAddr + " (" +
                            iRows + "x" + iCols + ")");

                        // Decide what range to add. Default is the
                        // Area itself; if a clip is configured and
                        // applicable, derive a clipped Range.
                        bool bUseClipped = false;
                        string sKeyAddr = sAddr;
                        if (iClipToRow > 0 && iLastRow > iClipToRow) {
                            if (iRow > iClipToRow) {
                                logger.info("      (skipped: starts below data floor row " +
                                    iClipToRow + ")");
                                continue;
                            }
                            try {
                                oClipped = oSheet.Range(
                                    oSheet.Cells(iRow, iCol),
                                    oSheet.Cells(iClipToRow, iCol + iCols - 1));
                                sKeyAddr = Convert.ToString(oClipped.Address,
                                    CultureInfo.InvariantCulture);
                                iLastRow = iClipToRow;
                                bUseClipped = true;
                                logger.info("      (clipped to data floor row " +
                                    iClipToRow + ": " + sKeyAddr + ")");
                            } catch (Exception exClip) {
                                logger.error("      clip failed, using original area: " +
                                    exClip.Message);
                                // oClipped may be null or partial;
                                // the finally will release it.
                                bUseClipped = false;
                            }
                        }

                        if (iLastRow > iMaxRowSeen) iMaxRowSeen = iLastRow;

                        if (!dRegions.ContainsKey(sKeyAddr)) {
                            if (bUseClipped) {
                                dRegions[sKeyAddr] = oClipped;
                                bClippedTransferred = true;
                            } else {
                                dRegions[sKeyAddr] = oArea;
                                bAreaTransferred = true;
                            }
                        }
                    } catch (Exception ex) {
                        logger.error("  SpecialCells area " + i + " failed: " +
                            ex.GetType().Name + ": " + ex.Message);
                    } finally {
                        if (!bAreaTransferred) comHelper.release(oArea);
                        if (!bClippedTransferred) comHelper.release(oClipped);
                    }
                }
            } finally {
                comHelper.release(oAreas);
                comHelper.release(oSpec);
                comHelper.release(oAllCells);
            }
            return iMaxRowSeen;
        }

        // Given a single seed cell coordinate, computes its
        // CurrentRegion and adds the result to dRegions (keyed by
        // address). Marks every cell inside the found region as
        // covered so subsequent seeds inside it don't re-query.
        // Shared by both seed-enumeration paths.
        private static void seedRegionFromCell(dynamic oSheet, int iR, int iC,
            Dictionary<string, dynamic> dRegions, HashSet<string> hsCovered)
        {
            dynamic oCell = null, oRegion = null;
            try {
                oCell = oSheet.Cells(iR, iC);
                oRegion = oCell.CurrentRegion;
                string sAddr = Convert.ToString(oRegion.Address,
                    CultureInfo.InvariantCulture);
                if (!dRegions.ContainsKey(sAddr)) {
                    dRegions[sAddr] = oRegion;
                    int iRR = (int)oRegion.Row;
                    int iCC = (int)oRegion.Column;
                    int iRN = (int)oRegion.Rows.Count;
                    int iCN = (int)oRegion.Columns.Count;
                    for (int rr = 0; rr < iRN; rr++)
                        for (int cc = 0; cc < iCN; cc++)
                            hsCovered.Add("R" + (iRR + rr) + "C" + (iCC + cc));
                    oRegion = null; // ownership transferred
                }
            } catch (Exception ex) {
                logger.error("  CurrentRegion at R" + iR + "C" + iC + " failed: " +
                    ex.GetType().Name + ": " + ex.Message);
            } finally {
                comHelper.release(oRegion);
                comHelper.release(oCell);
            }
        }

        // Populates a sheetRegion from an Excel Range object:
        // fetches the values and classifies header structure.
        // Minimum dimensions of DATA content (cells that are neither
        // column headers nor row labels) required for a region to
        // render as a semantic <table>. Below these floors on either
        // axis, the region is rendered as <p> paragraphs instead.
        //
        // A classic spreadsheet table minimum is 2x2 of data: at
        // least two data rows across at least two data columns.
        // Counted in total region size:
        //   - With column headers only:   3 rows x 2 cols
        //   - With row labels only:       2 rows x 3 cols
        //   - With both axes:             3 rows x 3 cols
        //   - With neither axis:          never a <table>
        // Regions that fall below these floors degrade gracefully:
        // a 2x2 region, for example, is just four cells and reads
        // better as paragraphs than as a one-data-row "table".
        public const int iMinDataRows = 2;
        public const int iMinDataCols = 2;

        private static sheetRegion buildRegion(dynamic oRange)
        {
            int iRows = 0, iCols = 0;
            try { iRows = (int)oRange.Rows.Count; } catch { }
            try { iCols = (int)oRange.Columns.Count; } catch { }
            if (iRows == 0 || iCols == 0) return null;

            int iRow = 1, iCol = 1;
            try { iRow = (int)oRange.Row;    } catch { }
            try { iCol = (int)oRange.Column; } catch { }

            object[,] aValues;
            try {
                object oRaw = oRange.Value2;
                aValues = normalizeToArray(oRaw, iRows, iCols);
            } catch {
                aValues = new object[iRows, iCols];
            }

            // Classify each axis independently. The minimum-size
            // check is applied AFTER classification (in writeRegion)
            // because the required dimensions depend on which axes
            // were classified as headers.
            //
            // A region needs at least 2 rows to even contemplate
            // having a header row, and at least 2 cols to have a
            // header column — otherwise the "header" would BE the
            // content with no data behind it.
            bool bCol = (iRows >= 2 && iCols >= 2) &&
                classifyRowAsHeaders(aValues, 0, iCols);
            bool bRow = (iRows >= 2 && iCols >= 2) &&
                classifyColAsHeaders(aValues, 0, iRows);

            return new sheetRegion {
                iRow = iRow, iCol = iCol,
                iRows = iRows, iCols = iCols,
                aValues = aValues,
                bColHeaders = bCol,
                bRowHeaders = bRow
            };
        }

        // A row (typically row 0) qualifies as column headers when
        // every cell is non-empty and all values are unique. This
        // matches xlHeaders.vbs.
        private static bool classifyRowAsHeaders(object[,] aValues, int iRowIdx, int iCols)
        {
            int iR0 = aValues.GetLowerBound(0);
            int iC0 = aValues.GetLowerBound(1);
            var hs = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
            for (int c = 0; c < iCols; c++) {
                object oVal = aValues[iR0 + iRowIdx, iC0 + c];
                if (oVal == null) return false;
                string s = Convert.ToString(oVal, CultureInfo.InvariantCulture).Trim();
                if (s.Length == 0) return false;
                if (!hs.Add(s)) return false;
            }
            return true;
        }

        // A column (typically column 0) qualifies as row headers
        // when every cell is non-empty and all values are unique.
        private static bool classifyColAsHeaders(object[,] aValues, int iColIdx, int iRows)
        {
            int iR0 = aValues.GetLowerBound(0);
            int iC0 = aValues.GetLowerBound(1);
            var hs = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
            for (int r = 0; r < iRows; r++) {
                object oVal = aValues[iR0 + r, iC0 + iColIdx];
                if (oVal == null) return false;
                string s = Convert.ToString(oVal, CultureInfo.InvariantCulture).Trim();
                if (s.Length == 0) return false;
                if (!hs.Add(s)) return false;
            }
            return true;
        }

        // Emits a single region. Decides between <table> and
        // paragraph rendering:
        //   - If neither axis classified as headers → paragraphs.
        //   - If classified but data-cell count falls below the
        //     iMinDataRows / iMinDataCols floor → paragraphs.
        //     (The "table" would be almost all header.)
        //   - Otherwise → <table> with appropriate scope markup.
        private static void writeRegion(TextWriter writer, sheetRegion region,
            string sSheetName, int iSheetNum, int iRegionIdx, int iRegionTotal)
        {
            if (!region.bColHeaders && !region.bRowHeaders) {
                writeRegionAsParagraphs(writer, region);
                return;
            }

            // Compute required total dimensions based on which
            // axes got classified. Every header row or column
            // consumes one unit from the corresponding total.
            int iReqRows = iMinDataRows + (region.bColHeaders ? 1 : 0);
            int iReqCols = iMinDataCols + (region.bRowHeaders ? 1 : 0);
            if (region.iRows < iReqRows || region.iCols < iReqCols) {
                // Not enough data behind the classified headers.
                // Demote: clear the flags and render the whole
                // region as paragraphs.
                region.bColHeaders = false;
                region.bRowHeaders = false;
                writeRegionAsParagraphs(writer, region);
                return;
            }

            string sSummaryId = "sheet-" + iSheetNum + "-region-" + iRegionIdx + "-summary";
            string sSummaryText = buildRegionSummary(region);
            string sCaption = buildRegionCaption(sSheetName, region, iRegionIdx, iRegionTotal);
            htmlWriter.writeTable(writer, region.aValues, sCaption, sSummaryId, sSummaryText,
                region.bColHeaders, region.bRowHeaders);
        }

        private static string buildRegionSummary(sheetRegion region)
        {
            var sb = new StringBuilder();
            sb.Append("Table with ").Append(region.iRows).Append(" rows and ");
            sb.Append(region.iCols).Append(" columns.");
            if (region.bColHeaders) sb.Append(" First row contains column headers.");
            if (region.bRowHeaders) sb.Append(" First column contains row headers.");
            return sb.ToString();
        }

        private static string buildRegionCaption(string sSheetName, sheetRegion region,
            int iRegionIdx, int iRegionTotal)
        {
            // If the sheet has only one region, omit the region
            // index; otherwise label the region for unambiguous
            // reference.
            var sb = new StringBuilder();
            sb.Append(sSheetName);
            if (iRegionTotal > 1)
                sb.Append(" \u2014 region ").Append(iRegionIdx);
            sb.Append(" \u2014 ").Append(region.iRows).Append(" rows \u00d7 ")
              .Append(region.iCols).Append(" columns");
            return sb.ToString();
        }

        // Renders a region as a sequence of <p> paragraphs, one per
        // non-empty row with cells joined by spaces. Used when the
        // region has neither column nor row headers (e.g., a single
        // column of prose or a loose block of text).
        private static void writeRegionAsParagraphs(TextWriter writer, sheetRegion region)
        {
            int iR0 = region.aValues.GetLowerBound(0);
            int iC0 = region.aValues.GetLowerBound(1);
            for (int r = 0; r < region.iRows; r++) {
                var sb = new StringBuilder();
                for (int c = 0; c < region.iCols; c++) {
                    object oVal = region.aValues[iR0 + r, iC0 + c];
                    if (oVal == null) continue;
                    string sCell = Convert.ToString(oVal, CultureInfo.InvariantCulture);
                    if (string.IsNullOrWhiteSpace(sCell)) continue;
                    if (sb.Length > 0) sb.Append(' ');
                    sb.Append(sCell.Trim());
                }
                if (sb.Length > 0)
                    writer.WriteLine("<p>" + htmlWriter.escape(sb.ToString()) + "</p>");
            }
        }

        private static object[,] normalizeToArray(object oRaw, int iRows, int iCols)
        {
            if (oRaw is object[,] a2) return a2;
            var a = new object[iRows, iCols];
            a[0, 0] = oRaw;
            return a;
        }

        // Plain-text conversion. Iterates through the same sheets and
        // cells as the HTML path, but writes each sheet as a
        // tab-separated block with a header line identifying the
        // sheet name and number. Blank rows separate sheets.
        public static void convertToText(string sInPath, string sOutPath)
        {
            int iAttempt = comHelper.retryOfficeOp(
                () => convertToTextAttempt(sInPath, sOutPath), "EXCEL", 3);
            if (iAttempt > 1) logger.info("Converted on attempt " + iAttempt + ": " + sOutPath);
        }

        private static void convertToTextAttempt(string sInPath, string sOutPath)
        {
            dynamic oApp = null, oWb = null;
            bool bWritten = false;
            try {
                oApp = comHelper.op("Excel: createApp", () => comHelper.createApp("Excel.Application"));
                comHelper.op("Excel: set Visible=false", () => { oApp.Visible = false; });
                comHelper.op("Excel: set DisplayAlerts=false", () => { oApp.DisplayAlerts = false; });
                comHelper.op("Excel: set ScreenUpdating=false", () => { oApp.ScreenUpdating = false; });
                try { oApp.AskToUpdateLinks = false; } catch { }
                try { oApp.EnableEvents = false; } catch { }

                oWb = comHelper.op("Excel: Workbooks.Open",
                    () => (object)oApp.Workbooks.Open(
                        Filename: sInPath,
                        ReadOnly: true,
                        UpdateLinks: 0,
                        AddToMru: false));

                using (var writer = new StreamWriter(sOutPath, false, new UTF8Encoding(true))) {
                    int iCount = comHelper.op("Excel: Worksheets.Count",
                        () => (int)oWb.Worksheets.Count);
                    for (int i = 1; i <= iCount; i++) {
                        dynamic oSheet = null;
                        try {
                            int iLocal = i;
                            oSheet = comHelper.op("Excel: Worksheets[" + iLocal + "]",
                                () => (object)oWb.Worksheets[iLocal]);
                            writeSheetAsText(writer, oSheet, i);
                        } finally {
                            comHelper.release(oSheet);
                        }
                    }
                }
                bWritten = true;
            } finally {
                try { if (oWb != null) { try { oWb.Close(false); } catch { } comHelper.release(oWb); } } catch { }
                try { if (oApp != null) { try { oApp.Quit(); } catch { } comHelper.release(oApp); } } catch { }
            }

            if (!bWritten) {
                try { if (File.Exists(sOutPath)) File.Delete(sOutPath); } catch { }
            }
        }

        private static void writeSheetAsText(TextWriter writer, dynamic oSheet, int iNum)
        {
            string sName = "";
            try { sName = Convert.ToString(oSheet.Name, CultureInfo.InvariantCulture); } catch { }
            string sHeader = string.IsNullOrWhiteSpace(sName)
                ? "=== Sheet " + iNum + " ==="
                : "=== Sheet " + iNum + ": " + sName.Trim() + " ===";
            writer.WriteLine(sHeader);

            // Text mode uses the same region-discovery infrastructure
            // as HTML mode. This ensures consistent behaviour between
            // the two output formats, including the SpecialCells
            // fallback for sheets whose UsedRange is too large to
            // load via Value2. Each region is emitted as its own
            // TSV block, separated by blank lines.
            logger.info("Sheet " + iNum + " (" +
                (string.IsNullOrWhiteSpace(sName) ? "unnamed" : sName) +
                "): discovering regions");
            var lsRegions = discoverRegions(oSheet);
            logger.info("Sheet " + iNum + ": found " + lsRegions.Count + " region(s)");
            if (lsRegions.Count == 0) { writer.WriteLine(); return; }

            bool bFirst = true;
            foreach (var oRegion in lsRegions) {
                if (!bFirst) writer.WriteLine();
                bFirst = false;
                writeRegionAsText(writer, oRegion);
            }
            writer.WriteLine();
        }

        // Emits one region as tab-separated text, one row per line.
        // Tabs and newlines inside cell values are replaced with
        // spaces so each spreadsheet row is exactly one line in the
        // output.
        private static void writeRegionAsText(TextWriter writer, sheetRegion region)
        {
            int iR0 = region.aValues.GetLowerBound(0);
            int iC0 = region.aValues.GetLowerBound(1);
            for (int r = 0; r < region.iRows; r++) {
                var sb = new StringBuilder();
                for (int c = 0; c < region.iCols; c++) {
                    if (c > 0) sb.Append('\t');
                    object oCell = region.aValues[iR0 + r, iC0 + c];
                    if (oCell != null) {
                        string sCell = Convert.ToString(oCell, CultureInfo.InvariantCulture);
                        sCell = sCell.Replace('\t', ' ').Replace('\r', ' ').Replace('\n', ' ');
                        sb.Append(sCell);
                    }
                }
                writer.WriteLine(sb.ToString());
            }
        }
    }

    // -----------------------------------------------------------------
    // CSV: native RFC 4180 parser.
    // -----------------------------------------------------------------
    public static class csvConverter
    {
        public static void convert(string sInPath, string sOutPath)
        {
            var dMeta = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase) {
                { "title", Path.GetFileNameWithoutExtension(sInPath) }
            };

            List<string[]> lsRows = parseCsv(File.ReadAllText(sInPath, Encoding.UTF8));
            if (lsRows.Count == 0) {
                using (var writer = new StreamWriter(sOutPath, false, new UTF8Encoding(false))) {
                    htmlWriter.writeHead(writer, dMeta);
                    writer.WriteLine("<main aria-labelledby=\"doc-title\">");
                    writer.WriteLine("<p><em>(empty CSV)</em></p>");
                    writer.WriteLine("</main>");
                    htmlWriter.writeFoot(writer);
                }
                return;
            }

            int iRows = lsRows.Count;
            int iCols = 0;
            foreach (var asRow in lsRows) if (asRow.Length > iCols) iCols = asRow.Length;

            var aValues = new object[iRows, iCols];
            for (int r = 0; r < iRows; r++) {
                string[] asRow = lsRows[r];
                for (int c = 0; c < iCols; c++)
                    aValues[r, c] = c < asRow.Length ? (object)asRow[c] : null;
            }

            using (var writer = new StreamWriter(sOutPath, false, new UTF8Encoding(false))) {
                htmlWriter.writeHead(writer, dMeta);
                writer.WriteLine("<main aria-labelledby=\"doc-title\">");
                string sSummaryId = "csv-summary";
                string sSummaryText = "Table with " + iRows + " rows and " + iCols + " columns. " +
                    "First row contains column headers.";
                string sCaption = Path.GetFileName(sInPath) + " \u2014 " +
                    iRows + " rows \u00d7 " + iCols + " columns";
                htmlWriter.writeTable(writer, aValues, sCaption, sSummaryId, sSummaryText,
                    bColHeaders: true, bRowHeaders: false);
                writer.WriteLine("</main>");
                htmlWriter.writeFoot(writer);
            }
        }

        private static List<string[]> parseCsv(string sIn)
        {
            var lsRows = new List<string[]>();
            var lsCurRow = new List<string>();
            var sb = new StringBuilder();
            bool bInQuotes = false;
            int i = 0;
            int iLen = sIn.Length;

            while (i < iLen) {
                char c = sIn[i];
                if (bInQuotes) {
                    if (c == '"') {
                        if (i + 1 < iLen && sIn[i + 1] == '"') { sb.Append('"'); i += 2; }
                        else { bInQuotes = false; i++; }
                    } else { sb.Append(c); i++; }
                } else {
                    if (c == '"') { bInQuotes = true; i++; }
                    else if (c == ',') { lsCurRow.Add(sb.ToString()); sb.Clear(); i++; }
                    else if (c == '\r' || c == '\n') {
                        lsCurRow.Add(sb.ToString()); sb.Clear();
                        lsRows.Add(lsCurRow.ToArray()); lsCurRow.Clear();
                        if (c == '\r' && i + 1 < iLen && sIn[i + 1] == '\n') i += 2;
                        else i++;
                    }
                    else { sb.Append(c); i++; }
                }
            }
            if (sb.Length > 0 || lsCurRow.Count > 0) {
                lsCurRow.Add(sb.ToString());
                lsRows.Add(lsCurRow.ToArray());
            }
            while (lsRows.Count > 0 && isEmptyRow(lsRows[lsRows.Count - 1]))
                lsRows.RemoveAt(lsRows.Count - 1);
            return lsRows;
        }

        private static bool isEmptyRow(string[] asRow)
        {
            if (asRow == null || asRow.Length == 0) return true;
            foreach (var s in asRow) if (!string.IsNullOrEmpty(s)) return false;
            return true;
        }
    }

    // -----------------------------------------------------------------
    // PowerPoint. Key fix: PowerPoint does not reliably respect
    // Application.Visible = false. We do NOT set it. Individual
    // presentation windows are hidden via WithWindow: msoFalse which
    // is supported on Office 2013+.
    // -----------------------------------------------------------------
    public static class pptConverter
    {
        public const int iPpSaveAsPng = 18;
        public const int iMsoTrue = -1;
        public const int iMsoFalse = 0;
        public const int iMsoPlaceholder = 14;
        public const int iMsoPicture = 13;
        public const int iMsoLinkedPicture = 11;
        public const int iMsoTable = 19;
        public const int iMsoTextEffect = 15;
        public const int iMsoGroup = 6;
        public const int iPpPlaceholderTitle = 1;
        public const int iPpPlaceholderCenterTitle = 3;
        public const int iPpPlaceholderSubtitle = 4;
        public const int iPpBulletNone = 0;
        public const int iPpBulletUnnumbered = 1;
        public const int iPpBulletNumbered = 2;
        public const int iPpAlertsNone = 1;
        public const int iPpShapeFormatPng = 2;

        public static void convert(string sInPath, string sOutPath)
        {
            int iAttempt = comHelper.retryOfficeOp(
                () => convertAttempt(sInPath, sOutPath), "POWERPNT", 3);
            if (iAttempt > 1) logger.info("Converted on attempt " + iAttempt + ": " + sOutPath);
        }

        private static void convertAttempt(string sInPath, string sOutPath)
        {
            dynamic oApp = null, oPres = null;
            string sTempDir = tempManager.newRunDir();
            bool bWritten = false;
            try {
                oApp = comHelper.op("PowerPoint: createApp",
                    () => comHelper.createApp("PowerPoint.Application"));

                // PowerPoint requires a visible Application in many
                // versions. Do NOT set oApp.Visible = false. Hide
                // individual presentation windows instead.
                try { oApp.DisplayAlerts = iPpAlertsNone; } catch { }

                oPres = comHelper.op("PowerPoint: Presentations.Open",
                    () => (object)oApp.Presentations.Open(
                        FileName: sInPath,
                        ReadOnly: iMsoTrue,
                        Untitled: iMsoFalse,
                        WithWindow: iMsoFalse));

                var dMeta = comHelper.readBuiltInDocProps(oPres);
                string sDocTitle = pickTitle(dMeta, sInPath);

                using (var writer = new StreamWriter(sOutPath, false, new UTF8Encoding(false))) {
                    htmlWriter.writeHead(writer, dMeta, sDocTitle);
                    // No synthetic navigation: PowerPoint decks have
                    // no notion of a table-of-contents in the source.
                    // Slides are accessed naturally via the <h2>
                    // section headings below.
                    writer.WriteLine("<main aria-labelledby=\"doc-title\">");

                    int iCount = comHelper.op("PowerPoint: Slides.Count",
                        () => (int)oPres.Slides.Count);
                    for (int i = 1; i <= iCount; i++) {
                        dynamic oSlide = null;
                        try {
                            int iLocal = i;
                            oSlide = comHelper.op("PowerPoint: Slides[" + iLocal + "]",
                                () => (object)oPres.Slides[iLocal]);
                            writeSlide(writer, oSlide, i, sTempDir);
                        } finally { comHelper.release(oSlide); }
                    }

                    writer.WriteLine("</main>");
                    htmlWriter.writeFoot(writer);
                }

                // Reaching here means the output stream closed cleanly
                // (the using block ran the StreamWriter's flush+close
                // without throwing). The file on disk is complete.
                bWritten = true;
            } finally {
                // Teardown is best-effort. If PowerPoint or its COM
                // infrastructure has already disconnected (a common
                // occurrence after a long conversion on some Office
                // builds), the Close/Quit calls may throw. Those
                // errors do not invalidate an already-written output
                // file; we swallow them so the conversion is reported
                // as the success it actually was.
                try { if (oPres != null) { try { oPres.Close(); } catch { } comHelper.release(oPres); } } catch { }
                try { if (oApp  != null) { try { oApp.Quit();  } catch { } comHelper.release(oApp);  } } catch { }
                try { tempManager.tryDelete(sTempDir); } catch { }
            }

            // If the body itself threw before writing completed, the
            // file on disk is partial. Delete it so the next retry
            // (or next run) does not see a truncated "[SKIP]" target.
            if (!bWritten) {
                try { if (File.Exists(sOutPath)) File.Delete(sOutPath); } catch { }
            }
        }

        private static string pickTitle(Dictionary<string, string> dMeta, string sInPath)
        {
            if (dMeta.TryGetValue("title", out var s) && !string.IsNullOrWhiteSpace(s)) return s;
            return Path.GetFileNameWithoutExtension(sInPath);
        }

        private static void writeSlide(TextWriter writer, dynamic oSlide, int iIdx, string sTempDir)
        {
            string sTitle = getSlideTitle(oSlide);
            bool bHasTitle = !string.IsNullOrWhiteSpace(sTitle);
            string sHeading = bHasTitle
                ? "<span class=\"visually-hidden\">Slide </span>" + iIdx +
                  ": " + htmlWriter.escape(sTitle.Trim())
                : "Slide " + iIdx;

            int iShapeCount = 0;
            try { iShapeCount = (int)oSlide.Shapes.Count; } catch { }
            logger.info("Slide " + iIdx + " (" +
                (bHasTitle ? sTitle.Trim() : "untitled") +
                "): " + iShapeCount + " shape(s)");

            writer.WriteLine("<section id=\"slide-" + iIdx + "\" aria-labelledby=\"slide-" + iIdx + "-h\">");
            writer.WriteLine("<h2 id=\"slide-" + iIdx + "-h\">" + sHeading + "</h2>");

            // Emit the shape content: text from text-bearing shapes
            // AND any picture shapes embedded in the slide (unless
            // --strip-images is set). We do NOT export a flattened
            // bitmap of the whole slide — that would hide the text
            // from screen readers. The slide's text is rendered
            // semantically; embedded pictures are embedded as
            // individual images.
            writer.WriteLine("<div class=\"slide-text\">");
            writeShapeContent(writer, oSlide, iIdx, sTempDir);
            writer.WriteLine("</div>");

            string sNotes = getSpeakerNotes(oSlide);
            if (!string.IsNullOrWhiteSpace(sNotes)) {
                writer.WriteLine("<aside class=\"slide-notes\" aria-label=\"Speaker notes for slide " + iIdx + "\">");
                writer.WriteLine("<p><strong>Speaker notes:</strong></p>");
                foreach (var sLine in sNotes.Split(new[] { '\r', '\n' }, StringSplitOptions.RemoveEmptyEntries))
                    writer.WriteLine("<p>" + htmlWriter.escape(sLine.Trim()) + "</p>");
                writer.WriteLine("</aside>");
            }

            writer.WriteLine("</section>");
        }

        // All interior lookups are wrapped in try/catch so that a
        // single bad shape or slide does not kill the whole run. We
        // keep references typed as dynamic until the last possible
        // access.
        private static string getSlideTitle(dynamic oSlide)
        {
            dynamic oShapes = null, oTitle = null, oTf = null;
            try {
                oShapes = oSlide.Shapes;
                // Shapes.HasTitle returns MsoTriState (Int32), not a
                // bool. Late-bound COM via `dynamic` will let the
                // DLR coerce an Int32 to a bool via `Convert.ToBoolean`
                // — sometimes. Do not rely on it. Read as int and
                // compare to msoTrue (-1).
                int iHasTitle = 0;
                try { iHasTitle = (int)oShapes.HasTitle; } catch { return ""; }
                if (iHasTitle != iMsoTrue) return "";
                oTitle = oShapes.Title;
                oTf = oTitle.TextFrame;
                int iHasText = 0;
                try { iHasText = (int)oTf.HasText; } catch { return ""; }
                if (iHasText != iMsoTrue) return "";
                return Convert.ToString(oTf.TextRange.Text, CultureInfo.InvariantCulture);
            } catch { return ""; }
            finally {
                comHelper.release(oTf);
                comHelper.release(oTitle);
                comHelper.release(oShapes);
            }
        }

        // Walk all shapes on a slide and emit appropriate content.
        // We dispatch on shape type and placeholder subtype:
        //   * title placeholders are handled separately by the caller
        //   * picture shapes export as embedded base64 <img>
        //   * table shapes render as <table>
        //   * text-effect shapes (WordArt) render as <p>
        //   * group shapes are walked recursively
        //   * everything else with text renders as <p> or <ul> based
        //     on its ParagraphFormat.Bullet.Type
        private static void writeShapeContent(TextWriter writer, dynamic oSlide, int iSlideIdx,
            string sTempDir)
        {
            dynamic oShapes = null;
            try {
                oShapes = oSlide.Shapes;
                int iCount = 0;
                try { iCount = (int)oShapes.Count; } catch { return; }

                int iPicSeq = 0;
                for (int i = 1; i <= iCount; i++) {
                    dynamic oShape = null;
                    try {
                        oShape = oShapes[i];
                        writeOneShape(writer, oShape, iSlideIdx, sTempDir, ref iPicSeq, true);
                    } catch {
                        // Skip any shape that refuses to cooperate.
                    } finally { comHelper.release(oShape); }
                }
            } catch { }
            finally { comHelper.release(oShapes); }
        }

        // Handles a single shape. bSkipTitle controls whether a title
        // placeholder should be skipped (for the top-level iteration
        // where the title is emitted as the <h2>). For nested group
        // shapes bSkipTitle is still true — a title placeholder
        // inside a group is still redundant with the slide heading.
        private static void writeOneShape(TextWriter writer, dynamic oShape, int iSlideIdx,
            string sTempDir, ref int iPicSeq, bool bSkipTitle)
        {
            if (bSkipTitle && isTitleShape(oShape)) return;

            int iType = 0;
            try { iType = (int)oShape.Type; } catch { return; }

            // Group shapes: recurse into their contained shapes.
            if (iType == iMsoGroup) {
                dynamic oInner = null;
                try {
                    oInner = oShape.GroupItems;
                    int iN = 0;
                    try { iN = (int)oInner.Count; } catch { return; }
                    for (int j = 1; j <= iN; j++) {
                        dynamic oChild = null;
                        try {
                            oChild = oInner[j];
                            writeOneShape(writer, oChild, iSlideIdx, sTempDir, ref iPicSeq, bSkipTitle);
                        } catch { }
                        finally { comHelper.release(oChild); }
                    }
                } catch { }
                finally { comHelper.release(oInner); }
                return;
            }

            // Picture shapes: export and embed.
            if (isPictureShape(oShape)) {
                iPicSeq++;
                writePictureShape(writer, oShape, iSlideIdx, iPicSeq, sTempDir);
                return;
            }

            // Table shapes: render as <table>.
            if (iType == iMsoTable || hasTable(oShape)) {
                writeTableShape(writer, oShape);
                return;
            }

            // Chart shapes: emit the chart title if present. The
            // chart body itself is usually also represented as a
            // picture in the slide; we surface the title so that
            // readers know what the chart is about.
            if (hasChart(oShape)) {
                writeChartTitle(writer, oShape);
                // fall through — a chart shape may also have a
                // text frame with additional text on some layouts
            }

            // Text-effect (WordArt) shapes: emit the effect text.
            if (iType == iMsoTextEffect) {
                writeTextEffectShape(writer, oShape);
                return;
            }

            // Everything else: emit text content if present.
            if (hasNonEmptyText(oShape)) writeTextShape(writer, oShape);
        }

        // True if this shape exposes a Chart child object.
        private static bool hasChart(dynamic oShape)
        {
            try { return (int)oShape.HasChart == iMsoTrue; }
            catch { return false; }
        }

        // True if this shape exposes a Table child object.
        private static bool hasTable(dynamic oShape)
        {
            try { return (int)oShape.HasTable == iMsoTrue; }
            catch { return false; }
        }

        private static void writeChartTitle(TextWriter writer, dynamic oShape)
        {
            string sTitle = "";
            try {
                dynamic oChart = oShape.Chart;
                try {
                    if ((int)oChart.HasTitle == iMsoTrue) {
                        sTitle = Convert.ToString(oChart.ChartTitle.Text,
                            CultureInfo.InvariantCulture);
                    }
                } catch { }
                comHelper.release(oChart);
            } catch { }
            if (!string.IsNullOrWhiteSpace(sTitle))
                writer.WriteLine("<p><strong>Chart: </strong>" +
                    htmlWriter.escape(sTitle.Trim()) + "</p>");
        }

        private static void writeTextEffectShape(TextWriter writer, dynamic oShape)
        {
            string sText = "";
            try {
                dynamic oFx = oShape.TextEffect;
                try { sText = Convert.ToString(oFx.Text, CultureInfo.InvariantCulture); }
                catch { }
                comHelper.release(oFx);
            } catch { }
            if (!string.IsNullOrWhiteSpace(sText))
                writer.WriteLine("<p>" + htmlWriter.escape(sText.Trim()) + "</p>");
        }

        // Emits a table-shape as an HTML table. First row becomes the
        // <thead>, remaining rows are <tbody>. Cells are accessed via
        // oTable.Cell(row, col), 1-indexed.
        private static void writeTableShape(TextWriter writer, dynamic oShape)
        {
            dynamic oTable = null;
            try {
                oTable = oShape.Table;
                int iRows = 0, iCols = 0;
                try { iRows = (int)oTable.Rows.Count; } catch { return; }
                try { iCols = (int)oTable.Columns.Count; } catch { return; }
                if (iRows == 0 || iCols == 0) return;

                writer.WriteLine("<table>");
                for (int r = 1; r <= iRows; r++) {
                    bool bIsHeader = (r == 1);
                    if (bIsHeader) writer.WriteLine("<thead>");
                    if (r == 2)    writer.WriteLine("<tbody>");
                    writer.WriteLine("<tr>");
                    for (int c = 1; c <= iCols; c++) {
                        string sCell = "";
                        dynamic oCell = null;
                        try {
                            oCell = oTable.Cell(r, c);
                            try {
                                sCell = Convert.ToString(
                                    oCell.Shape.TextFrame.TextRange.Text,
                                    CultureInfo.InvariantCulture);
                            } catch { }
                        } catch { }
                        finally { comHelper.release(oCell); }
                        string sTag = bIsHeader ? "th" : "td";
                        writer.WriteLine("<" + sTag + ">" +
                            htmlWriter.escape((sCell ?? "").Trim()) + "</" + sTag + ">");
                    }
                    writer.WriteLine("</tr>");
                    if (bIsHeader) writer.WriteLine("</thead>");
                    if (r == iRows && iRows > 1) writer.WriteLine("</tbody>");
                }
                writer.WriteLine("</table>");
            } catch { }
            finally { comHelper.release(oTable); }
        }

        // Renders a text-bearing shape. Uses ParagraphFormat.Bullet.Type
        // to decide between plain paragraphs, an unordered list, or an
        // ordered list. Also handles subtitle placeholders by rendering
        // their content as a byline-style paragraph rather than a
        // heading.
        private static void writeTextShape(TextWriter writer, dynamic oShape)
        {
            dynamic oTr = null;
            try {
                oTr = oShape.TextFrame.TextRange;
                string sText = Convert.ToString(oTr.Text, CultureInfo.InvariantCulture);
                if (string.IsNullOrWhiteSpace(sText)) return;

                int iBulletType = iPpBulletNone;
                try { iBulletType = (int)oTr.ParagraphFormat.Bullet.Type; } catch { }

                // Split the shape's text into paragraphs. PowerPoint
                // uses \r (CR) as a paragraph separator within a
                // TextRange. Vertical-tab (\v) is a soft line break
                // within a paragraph.
                var aParas = sText.Split('\r');
                var lsParas = new List<string>();
                foreach (var sRaw in aParas) {
                    // Collapse soft line breaks to spaces.
                    string s = sRaw.Replace('\v', ' ').Trim();
                    if (s.Length > 0) lsParas.Add(s);
                }
                if (lsParas.Count == 0) return;

                bool bIsSubtitle = isSubtitleShape(oShape);

                if (iBulletType == iPpBulletUnnumbered) {
                    writer.WriteLine("<ul>");
                    foreach (var s in lsParas)
                        writer.WriteLine("<li>" + htmlWriter.escape(s) + "</li>");
                    writer.WriteLine("</ul>");
                } else if (iBulletType == iPpBulletNumbered) {
                    writer.WriteLine("<ol>");
                    foreach (var s in lsParas)
                        writer.WriteLine("<li>" + htmlWriter.escape(s) + "</li>");
                    writer.WriteLine("</ol>");
                } else {
                    // Plain paragraphs. Subtitle placeholders get a
                    // special class so they can be styled as a byline.
                    string sClass = bIsSubtitle ? " class=\"subtitle\"" : "";
                    foreach (var s in lsParas)
                        writer.WriteLine("<p" + sClass + ">" + htmlWriter.escape(s) + "</p>");
                }
            } finally { comHelper.release(oTr); }
        }

        private static bool isSubtitleShape(dynamic oShape)
        {
            try {
                int iType = (int)oShape.Type;
                if (iType != iMsoPlaceholder) return false;
                int iPh = (int)oShape.PlaceholderFormat.Type;
                return iPh == iPpPlaceholderSubtitle;
            } catch { return false; }
        }

        // True if the shape IS a picture, or is a placeholder whose
        // contained type is a picture (e.g., a picture placeholder on
        // a slide layout).
        private static bool isPictureShape(dynamic oShape)
        {
            try {
                int iType = (int)oShape.Type;
                if (iType == iMsoPicture || iType == iMsoLinkedPicture) return true;
                if (iType == iMsoPlaceholder) {
                    try {
                        int iContained = (int)oShape.PlaceholderFormat.ContainedType;
                        return iContained == iMsoPicture || iContained == iMsoLinkedPicture;
                    } catch { return false; }
                }
                return false;
            } catch { return false; }
        }

        // Exports a picture shape to PNG via Shape.Export, embeds it
        // as a base64 data URL, and uses the shape's AlternativeText
        // (set by the slide author) as the alt attribute. When
        // --strip-images is set, the shape is skipped entirely (no <img>
        // tag, no alt-text-only placeholder).
        private static void writePictureShape(TextWriter writer, dynamic oShape,
            int iSlideIdx, int iPicSeq, string sTempDir)
        {
            if (program.bStripImages) return;

            string sPng = Path.Combine(sTempDir,
                "slide_" + iSlideIdx.ToString("D4") + "_pic_" + iPicSeq.ToString("D2") + ".png");

            string sAlt = "";
            try {
                string sRaw = Convert.ToString(oShape.AlternativeText,
                    CultureInfo.InvariantCulture);
                if (!string.IsNullOrWhiteSpace(sRaw)) sAlt = sRaw.Trim();
            } catch { }

            try {
                oShape.Export(sPng, iPpShapeFormatPng, 0, 0);
                if (!File.Exists(sPng)) return;

                byte[] binPng = File.ReadAllBytes(sPng);
                string sB64 = Convert.ToBase64String(binPng);
                writer.WriteLine("<figure>");
                writer.WriteLine("<img src=\"data:image/png;base64," + sB64 +
                    "\" alt=\"" + htmlWriter.escape(sAlt) + "\">");
                writer.WriteLine("</figure>");
                try { File.Delete(sPng); } catch { }
            } catch {
                // If the export fails, skip the picture silently.
                // We never want to emit a broken image or a caption
                // with no associated image.
            }
        }

        private static bool isTitleShape(dynamic oShape)
        {
            try {
                int iType = (int)oShape.Type;
                if (iType != iMsoPlaceholder) return false;
                int iPh = (int)oShape.PlaceholderFormat.Type;
                return iPh == iPpPlaceholderTitle || iPh == iPpPlaceholderCenterTitle;
            } catch { return false; }
        }

        private static bool hasNonEmptyText(dynamic oShape)
        {
            try {
                // HasTextFrame returns MsoTriState (Int32), not bool.
                // Read as int and compare to msoTrue (-1).
                int iHasFrame = 0;
                try { iHasFrame = (int)oShape.HasTextFrame; } catch { return false; }
                if (iHasFrame != iMsoTrue) return false;
                return (int)oShape.TextFrame.HasText == iMsoTrue;
            } catch { return false; }
        }

        private static string getSpeakerNotes(dynamic oSlide)
        {
            dynamic oNp = null, oShapes = null;
            try {
                oNp = oSlide.NotesPage;
                oShapes = oNp.Shapes;
                int iN = 0;
                try { iN = (int)oShapes.Count; } catch { return ""; }
                for (int i = 1; i <= iN; i++) {
                    dynamic oShape = null;
                    try {
                        oShape = oShapes[i];
                        if (!hasNonEmptyText(oShape)) continue;
                        string s = Convert.ToString(
                            oShape.TextFrame.TextRange.Text, CultureInfo.InvariantCulture);
                        if (s != null && s.Length > 3 && !int.TryParse(s.Trim(), out _))
                            return s;
                    } catch { }
                    finally { comHelper.release(oShape); }
                }
            } catch { }
            finally {
                comHelper.release(oShapes);
                comHelper.release(oNp);
            }
            return "";
        }

        // Plain-text conversion. Iterates slides the same way as the
        // HTML path, but writes titles + shape text + speaker notes
        // as readable prose with slide-number markers.
        public static void convertToText(string sInPath, string sOutPath)
        {
            int iAttempt = comHelper.retryOfficeOp(
                () => convertToTextAttempt(sInPath, sOutPath), "POWERPNT", 3);
            if (iAttempt > 1) logger.info("Converted on attempt " + iAttempt + ": " + sOutPath);
        }

        private static void convertToTextAttempt(string sInPath, string sOutPath)
        {
            dynamic oApp = null, oPres = null;
            bool bWritten = false;
            try {
                oApp = comHelper.op("PowerPoint: createApp",
                    () => comHelper.createApp("PowerPoint.Application"));
                try { oApp.DisplayAlerts = iPpAlertsNone; } catch { }

                oPres = comHelper.op("PowerPoint: Presentations.Open",
                    () => (object)oApp.Presentations.Open(
                        FileName: sInPath,
                        ReadOnly: iMsoTrue,
                        Untitled: iMsoFalse,
                        WithWindow: iMsoFalse));

                using (var writer = new StreamWriter(sOutPath, false, new UTF8Encoding(true))) {
                    int iCount = comHelper.op("PowerPoint: Slides.Count",
                        () => (int)oPres.Slides.Count);
                    for (int i = 1; i <= iCount; i++) {
                        dynamic oSlide = null;
                        try {
                            int iLocal = i;
                            oSlide = comHelper.op("PowerPoint: Slides[" + iLocal + "]",
                                () => (object)oPres.Slides[iLocal]);
                            writeSlideAsText(writer, oSlide, i);
                        } finally { comHelper.release(oSlide); }
                    }
                }
                bWritten = true;
            } finally {
                try { if (oPres != null) { try { oPres.Close(); } catch { } comHelper.release(oPres); } } catch { }
                try { if (oApp  != null) { try { oApp.Quit();  } catch { } comHelper.release(oApp);  } } catch { }
            }

            if (!bWritten) {
                try { if (File.Exists(sOutPath)) File.Delete(sOutPath); } catch { }
            }
        }

        private static void writeSlideAsText(TextWriter writer, dynamic oSlide, int iIdx)
        {
            string sTitle = getSlideTitle(oSlide);
            bool bHasTitle = !string.IsNullOrWhiteSpace(sTitle);
            int iShapeCount = 0;
            try { iShapeCount = (int)oSlide.Shapes.Count; } catch { }
            logger.info("Slide " + iIdx + " (" +
                (bHasTitle ? sTitle.Trim() : "untitled") +
                "): " + iShapeCount + " shape(s)");

            string sHeader = bHasTitle
                ? "=== Slide " + iIdx + ": " + sTitle.Trim() + " ==="
                : "=== Slide " + iIdx + " ===";
            writer.WriteLine(sHeader);

            // Iterate text-bearing shapes (skipping title, which we
            // already wrote). Picture shapes are ignored in plain-
            // text mode — text extraction from images is out of
            // scope.
            dynamic oShapes = null;
            try {
                oShapes = oSlide.Shapes;
                int iN = 0;
                try { iN = (int)oShapes.Count; } catch { }
                for (int i = 1; i <= iN; i++) {
                    dynamic oShape = null;
                    try {
                        oShape = oShapes[i];
                        writeOneShapeAsText(writer, oShape, true);
                    } catch { }
                    finally { comHelper.release(oShape); }
                }
            } catch { }
            finally { comHelper.release(oShapes); }

            string sNotes = getSpeakerNotes(oSlide);
            if (!string.IsNullOrWhiteSpace(sNotes)) {
                writer.WriteLine();
                writer.WriteLine("Speaker notes:");
                foreach (var sLine in sNotes.Split(new[] { '\r', '\n' },
                    StringSplitOptions.RemoveEmptyEntries)) {
                    string s = sLine.Trim();
                    if (s.Length > 0) writer.WriteLine(s);
                }
            }
            writer.WriteLine();
        }

        // Text-mode analogue of writeOneShape. Skips pictures (no
        // text to extract) but handles groups, tables, text effects,
        // charts, and text-bearing shapes.
        private static void writeOneShapeAsText(TextWriter writer, dynamic oShape, bool bSkipTitle)
        {
            if (bSkipTitle && isTitleShape(oShape)) return;

            int iType = 0;
            try { iType = (int)oShape.Type; } catch { return; }

            if (iType == iMsoGroup) {
                dynamic oInner = null;
                try {
                    oInner = oShape.GroupItems;
                    int iN = 0;
                    try { iN = (int)oInner.Count; } catch { return; }
                    for (int j = 1; j <= iN; j++) {
                        dynamic oChild = null;
                        try {
                            oChild = oInner[j];
                            writeOneShapeAsText(writer, oChild, bSkipTitle);
                        } catch { }
                        finally { comHelper.release(oChild); }
                    }
                } catch { }
                finally { comHelper.release(oInner); }
                return;
            }

            if (isPictureShape(oShape)) return;  // nothing to emit in text mode

            if (iType == iMsoTable || hasTable(oShape)) {
                writeTableShapeAsText(writer, oShape);
                return;
            }

            if (hasChart(oShape)) {
                writeChartTitleAsText(writer, oShape);
                // fall through
            }

            if (iType == iMsoTextEffect) {
                writeTextEffectAsText(writer, oShape);
                return;
            }

            if (hasNonEmptyText(oShape)) writeTextShapeAsText(writer, oShape);
        }

        private static void writeChartTitleAsText(TextWriter writer, dynamic oShape)
        {
            string sTitle = "";
            try {
                dynamic oChart = oShape.Chart;
                try {
                    if ((int)oChart.HasTitle == iMsoTrue)
                        sTitle = Convert.ToString(oChart.ChartTitle.Text,
                            CultureInfo.InvariantCulture);
                } catch { }
                comHelper.release(oChart);
            } catch { }
            if (!string.IsNullOrWhiteSpace(sTitle))
                writer.WriteLine("Chart: " + sTitle.Trim());
        }

        private static void writeTextEffectAsText(TextWriter writer, dynamic oShape)
        {
            string sText = "";
            try {
                dynamic oFx = oShape.TextEffect;
                try { sText = Convert.ToString(oFx.Text, CultureInfo.InvariantCulture); }
                catch { }
                comHelper.release(oFx);
            } catch { }
            if (!string.IsNullOrWhiteSpace(sText))
                writer.WriteLine(sText.Trim());
        }

        private static void writeTableShapeAsText(TextWriter writer, dynamic oShape)
        {
            dynamic oTable = null;
            try {
                oTable = oShape.Table;
                int iRows = 0, iCols = 0;
                try { iRows = (int)oTable.Rows.Count; } catch { return; }
                try { iCols = (int)oTable.Columns.Count; } catch { return; }
                if (iRows == 0 || iCols == 0) return;

                for (int r = 1; r <= iRows; r++) {
                    var sb = new StringBuilder();
                    for (int c = 1; c <= iCols; c++) {
                        if (c > 1) sb.Append('\t');
                        string sCell = "";
                        dynamic oCell = null;
                        try {
                            oCell = oTable.Cell(r, c);
                            try {
                                sCell = Convert.ToString(
                                    oCell.Shape.TextFrame.TextRange.Text,
                                    CultureInfo.InvariantCulture);
                            } catch { }
                        } catch { }
                        finally { comHelper.release(oCell); }
                        sCell = (sCell ?? "").Replace('\t', ' ')
                            .Replace('\r', ' ').Replace('\n', ' ').Replace('\v', ' ').Trim();
                        sb.Append(sCell);
                    }
                    writer.WriteLine(sb.ToString());
                }
            } catch { }
            finally { comHelper.release(oTable); }
        }

        private static void writeTextShapeAsText(TextWriter writer, dynamic oShape)
        {
            dynamic oTr = null;
            try {
                oTr = oShape.TextFrame.TextRange;
                string sText = Convert.ToString(oTr.Text, CultureInfo.InvariantCulture);
                if (string.IsNullOrWhiteSpace(sText)) return;
                // PowerPoint uses \r as paragraph separator; \v is a
                // soft line break within a paragraph. Collapse soft
                // breaks to spaces, emit each paragraph as its own
                // line in the output.
                foreach (var sPara in sText.Split('\r')) {
                    string s = sPara.Replace('\v', ' ').Trim();
                    if (s.Length > 0) writer.WriteLine(s);
                }
            } finally { comHelper.release(oTr); }
        }
    }

    // -----------------------------------------------------------------
    // HTML: rewrap existing .html/.htm with our accessibility skeleton.
    // -----------------------------------------------------------------
    public static class htmlConverter
    {
        public static void convert(string sInPath, string sOutPath)
        {
            string sRaw = File.ReadAllText(sInPath, detectEncoding(sInPath));

            var dMeta = extractMeta(sRaw);
            if (!dMeta.ContainsKey("title") || string.IsNullOrWhiteSpace(dMeta["title"]))
                dMeta["title"] = Path.GetFileNameWithoutExtension(sInPath);

            string sBody = extractBody(sRaw);
            sBody = Regex.Replace(sBody, @"<style.*?</style>", "", RegexOptions.Singleline | RegexOptions.IgnoreCase);
            sBody = Regex.Replace(sBody, @"<script.*?</script>", "", RegexOptions.Singleline | RegexOptions.IgnoreCase);
            sBody = Regex.Replace(sBody, @"<!--.*?-->", "", RegexOptions.Singleline);
            sBody = htmlWriter.inlineOrStripImages(sBody, Path.GetDirectoryName(sInPath) ?? "");
            sBody = htmlWriter.shiftAndClampHeadings(sBody, 2);

            using (var writer = new StreamWriter(sOutPath, false, new UTF8Encoding(false))) {
                htmlWriter.writeHead(writer, dMeta);
                writer.WriteLine("<main aria-labelledby=\"doc-title\">");
                writer.WriteLine(sBody);
                writer.WriteLine("</main>");
                htmlWriter.writeFoot(writer);
            }
        }

        private static Encoding detectEncoding(string sPath)
        {
            try {
                byte[] binHead = new byte[Math.Min(2048, (int)new FileInfo(sPath).Length)];
                using (var fileStream = File.OpenRead(sPath)) fileStream.Read(binHead, 0, binHead.Length);
                if (binHead.Length >= 3 && binHead[0] == 0xEF && binHead[1] == 0xBB && binHead[2] == 0xBF)
                    return Encoding.UTF8;
                if (binHead.Length >= 2 && binHead[0] == 0xFF && binHead[1] == 0xFE)
                    return Encoding.Unicode;
                if (binHead.Length >= 2 && binHead[0] == 0xFE && binHead[1] == 0xFF)
                    return Encoding.BigEndianUnicode;
                string sHead = Encoding.ASCII.GetString(binHead);
                Match m = Regex.Match(sHead,
                    @"<meta[^>]+charset\s*=\s*['""]?(?<cs>[A-Za-z0-9_\-]+)",
                    RegexOptions.IgnoreCase);
                if (m.Success) {
                    try { return Encoding.GetEncoding(m.Groups["cs"].Value); } catch { }
                }
            } catch { }
            return Encoding.UTF8;
        }

        private static Dictionary<string, string> extractMeta(string sHtml)
        {
            var d = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);

            Match oTitle = Regex.Match(sHtml, @"<title[^>]*>(?<t>.*?)</title>",
                RegexOptions.Singleline | RegexOptions.IgnoreCase);
            if (oTitle.Success) {
                string sT = htmlToText(oTitle.Groups["t"].Value).Trim();
                if (!string.IsNullOrWhiteSpace(sT)) d["title"] = sT;
            }

            Match matchLang = Regex.Match(sHtml, @"<html[^>]+lang\s*=\s*['""]?(?<l>[A-Za-z\-]+)",
                RegexOptions.IgnoreCase);
            if (matchLang.Success) d["language"] = matchLang.Groups["l"].Value;

            foreach (Match m in Regex.Matches(sHtml, @"<meta\s+(?<attrs>[^>]*)>",
                RegexOptions.IgnoreCase)) {
                string sAttrs = m.Groups["attrs"].Value;
                string sName = attrVal(sAttrs, "name");
                string sContent = attrVal(sAttrs, "content");
                if (string.IsNullOrWhiteSpace(sName) || string.IsNullOrWhiteSpace(sContent))
                    continue;
                switch (sName.ToLowerInvariant()) {
                    case "author":       d["author"]      = sContent; break;
                    case "description":  d["description"] = sContent; break;
                    case "keywords":     d["keywords"]    = sContent; break;
                    case "dc.title":     if (!d.ContainsKey("title"))       d["title"]       = sContent; break;
                    case "dc.creator":   if (!d.ContainsKey("author"))      d["author"]      = sContent; break;
                    case "dc.subject":   if (!d.ContainsKey("subject"))     d["subject"]     = sContent; break;
                    case "dc.description": if (!d.ContainsKey("description")) d["description"] = sContent; break;
                    case "dc.language":  if (!d.ContainsKey("language"))    d["language"]    = sContent; break;
                }
            }
            return d;
        }

        private static string attrVal(string sAttrs, string sAttrName)
        {
            Match m = Regex.Match(sAttrs,
                sAttrName + @"\s*=\s*(?:""(?<v>[^""]*)""|'(?<v>[^']*)'|(?<v>[^\s>]+))",
                RegexOptions.IgnoreCase);
            return m.Success ? m.Groups["v"].Value : "";
        }

        private static string extractBody(string sHtml)
        {
            Match m = Regex.Match(sHtml, @"<body[^>]*>(?<b>.*?)</body>",
                RegexOptions.Singleline | RegexOptions.IgnoreCase);
            return m.Success ? m.Groups["b"].Value : sHtml;
        }

        private static string htmlToText(string sHtml)
        {
            string s = Regex.Replace(sHtml, @"<[^>]+>", "");
            s = s.Replace("&amp;", "&").Replace("&lt;", "<").Replace("&gt;", ">")
                 .Replace("&quot;", "\"").Replace("&#39;", "'").Replace("&nbsp;", " ");
            return s;
        }

        // Plain-text conversion. Strips markup while preserving the
        // line structure implied by block-level tags (paragraphs,
        // headings, list items, line breaks). Images are dropped
        // regardless of --strip-images; plain text has no image
        // concept.
        public static void convertToText(string sInPath, string sOutPath)
        {
            string sRaw = File.ReadAllText(sInPath, detectEncoding(sInPath));
            string sBody = extractBody(sRaw);

            // Remove style/script/comment content before tag stripping
            // so their contents don't leak into the text.
            sBody = Regex.Replace(sBody, @"<style.*?</style>", "",
                RegexOptions.Singleline | RegexOptions.IgnoreCase);
            sBody = Regex.Replace(sBody, @"<script.*?</script>", "",
                RegexOptions.Singleline | RegexOptions.IgnoreCase);
            sBody = Regex.Replace(sBody, @"<!--.*?-->", "", RegexOptions.Singleline);

            string sText = htmlBlocksToText(sBody);
            File.WriteAllText(sOutPath, sText, new UTF8Encoding(true));
        }

        // Replace block-level closing tags and <br> with newlines so
        // the visual structure survives the tag strip. Collapses
        // runs of blank lines to at most two. Decodes common named
        // entities.
        // Converts an HTML body fragment to plain text, preserving
        // paragraph boundaries but NOT wrapping within paragraphs.
        // The resulting text has one line per logical paragraph
        // (headings, list items, table rows, etc.), not one line per
        // source-line-of-HTML. This is the Pandoc "--wrap=none"
        // style and lets editors and screen readers apply their own
        // natural word-wrapping without fighting embedded hard
        // returns at an arbitrary column.
        public static string htmlBlocksToText(string sHtml)
        {
            // Strip entire <figure>...</figure> blocks before
            // processing anything else. A figure contains an image
            // and optionally a <figcaption>; keeping the caption
            // without the image would violate the "no orphan
            // image references or captions" rule.
            string s = Regex.Replace(sHtml, @"<figure\b[^>]*>.*?</figure>", "",
                RegexOptions.Singleline | RegexOptions.IgnoreCase);

            // We need to distinguish "paragraph boundary" (which
            // becomes a real newline in the output) from "any other
            // whitespace, including source-file newlines inside a
            // paragraph" (which should be collapsed to a single
            // space). We use a rare sentinel character — U+E000,
            // the start of the Private Use Area — as the marker
            // for real paragraph breaks, and then normalize all
            // other whitespace before converting the marker back.
            const string sBreakSentinel = "\uE000";

            // Place the sentinel at the boundaries of block-level
            // elements. Double sentinels signal a stronger break
            // (blank line in output).
            s = Regex.Replace(s, @"<br\s*/?>", sBreakSentinel, RegexOptions.IgnoreCase);
            s = Regex.Replace(s,
                @"</(p|div|blockquote|pre|section|article|header|footer|nav|aside|figcaption)\s*>",
                sBreakSentinel + sBreakSentinel, RegexOptions.IgnoreCase);
            s = Regex.Replace(s, @"</(li|dt|dd|tr)\s*>", sBreakSentinel, RegexOptions.IgnoreCase);
            s = Regex.Replace(s, @"</(h[1-6])\s*>", sBreakSentinel + sBreakSentinel, RegexOptions.IgnoreCase);
            s = Regex.Replace(s, @"<(h[1-6])\b[^>]*>", sBreakSentinel, RegexOptions.IgnoreCase);
            s = Regex.Replace(s, @"<li\b[^>]*>", sBreakSentinel + "- ", RegexOptions.IgnoreCase);

            // Strip remaining tags.
            s = Regex.Replace(s, @"<[^>]+>", "");

            // Decode common entities.
            s = s.Replace("&amp;", "&").Replace("&lt;", "<").Replace("&gt;", ">")
                 .Replace("&quot;", "\"").Replace("&apos;", "'").Replace("&#39;", "'")
                 .Replace("&nbsp;", " ").Replace("&mdash;", "\u2014")
                 .Replace("&ndash;", "\u2013").Replace("&hellip;", "\u2026");
            // Numeric character references (decimal only; hex is rarer).
            s = Regex.Replace(s, @"&#(\d+);", m => {
                int iCode;
                if (int.TryParse(m.Groups[1].Value, out iCode) && iCode > 0 && iCode <= 0x10FFFF)
                    return char.ConvertFromUtf32(iCode);
                return m.Value;
            });

            // Collapse ALL runs of whitespace (including literal
            // newlines from the HTML source) into a single space.
            // This is what removes intra-paragraph hard returns.
            s = Regex.Replace(s, @"[ \t\r\n\f\v]+", " ");

            // Replace sentinels with newlines. A double sentinel
            // (which was used for stronger boundaries like
            // paragraph and heading ends) becomes a blank line.
            // Collapse runs of sentinels to at most two so we never
            // emit more than one blank line in a row.
            s = Regex.Replace(s, sBreakSentinel + "+", m => {
                return m.Length >= 2 ? "\n\n" : "\n";
            });

            // Trim leading/trailing whitespace on each line and the
            // whole document.
            var sbOut = new StringBuilder();
            bool bSawAnyText = false;
            int iBlank = 0;
            foreach (var sLine in s.Split('\n')) {
                string t = sLine.Trim();
                if (t.Length == 0) {
                    if (bSawAnyText) iBlank++;
                    continue;
                }
                if (iBlank > 0) sbOut.AppendLine();
                sbOut.AppendLine(t);
                bSawAnyText = true;
                iBlank = 0;
            }
            return sbOut.ToString();
        }
    }

    // -----------------------------------------------------------------
    // Markdown. Uses the embedded Markdig library (CommonMark-compliant
    // with advanced extensions: pipe tables, footnotes, definition
    // lists, task lists, autolinks, etc.). Markdig is loaded from an
    // embedded manifest resource via the AssemblyResolve handler
    // registered in program.Main.
    // -----------------------------------------------------------------
    public static class markdownConverter
    {
        public static void convert(string sInPath, string sOutPath)
        {
            string sMd = File.ReadAllText(sInPath, Encoding.UTF8);

            // Strip the YAML front matter block (Pandoc-style) from
            // the body before rendering, and parse the scalars in it
            // to populate metadata. Markdig's UseYamlFrontMatter()
            // also removes the block from its HTML output, but we
            // still want access to the field values.
            var dMeta = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);
            parseYamlFrontMatter(sMd, dMeta);

            if (!dMeta.ContainsKey("title") || string.IsNullOrWhiteSpace(dMeta["title"]))
                dMeta["title"] = firstHeadingOrBasename(sMd, sInPath);

            string sBody = renderMarkdown(sMd);
            sBody = htmlWriter.inlineOrStripImages(sBody, Path.GetDirectoryName(sInPath) ?? "");
            sBody = htmlWriter.shiftAndClampHeadings(sBody, 2);

            using (var writer = new StreamWriter(sOutPath, false, new UTF8Encoding(false))) {
                htmlWriter.writeHead(writer, dMeta);
                writer.WriteLine("<main aria-labelledby=\"doc-title\">");
                writer.WriteLine(sBody);
                writer.WriteLine("</main>");
                htmlWriter.writeFoot(writer);
            }
        }

        // Parses a leading YAML front matter block of the form
        //     ---
        //     key: value
        //     key: "quoted value"
        //     ---
        // into the supplied metadata dictionary. Only flat scalar
        // fields are supported (this is what Pandoc front matter
        // almost always contains). Recognized keys are mapped to the
        // dictionary keys used by htmlWriter (title, author, subject,
        // description, keywords, language, created, modified).
        // Pandoc also permits "..." as the closing fence; we accept
        // either --- or ... as a closer.
        private static void parseYamlFrontMatter(string sMd, Dictionary<string, string> dMeta)
        {
            if (string.IsNullOrEmpty(sMd)) return;

            // Normalize line endings for scanning. We do not mutate
            // sMd itself; Markdig handles CR/LF fine.
            string[] asLines = sMd.Replace("\r\n", "\n").Split('\n');
            if (asLines.Length < 2) return;
            if (asLines[0].Trim() != "---") return;

            int iEnd = -1;
            for (int i = 1; i < asLines.Length; i++) {
                string sTrim = asLines[i].Trim();
                if (sTrim == "---" || sTrim == "...") { iEnd = i; break; }
            }
            if (iEnd < 0) return;

            for (int i = 1; i < iEnd; i++) {
                string sLine = asLines[i];
                if (sLine.Length == 0) continue;

                // Skip list items and nested structure; we only
                // handle top-level scalar key: value pairs.
                if (sLine.StartsWith(" ") || sLine.StartsWith("\t")) continue;
                if (sLine.StartsWith("#")) continue; // YAML comment

                int iColon = sLine.IndexOf(':');
                if (iColon <= 0) continue;

                string sKey = sLine.Substring(0, iColon).Trim();
                string sVal = sLine.Substring(iColon + 1).Trim();
                if (sVal.Length == 0) continue;

                // Strip surrounding quotes if present.
                if ((sVal.StartsWith("\"") && sVal.EndsWith("\"")) ||
                    (sVal.StartsWith("'")  && sVal.EndsWith("'"))) {
                    if (sVal.Length >= 2)
                        sVal = sVal.Substring(1, sVal.Length - 2);
                }

                // Strip trailing comment (" # comment") if present
                // and value is not quoted.
                int iHash = sVal.IndexOf(" #");
                if (iHash > 0) sVal = sVal.Substring(0, iHash).TrimEnd();

                mapYamlKey(dMeta, sKey, sVal);
            }
        }

        private static void mapYamlKey(Dictionary<string, string> dMeta,
            string sKey, string sVal)
        {
            switch (sKey.ToLowerInvariant()) {
                case "title":       dMeta["title"]       = sVal; break;
                case "subtitle":    dMeta["subtitle"]    = sVal; break;
                case "author":      dMeta["author"]      = sVal; break;
                case "date":        dMeta["created"]     = sVal; break;
                case "modified":    dMeta["modified"]    = sVal; break;
                case "lang":
                case "language":    dMeta["language"]    = sVal; break;
                case "subject":     dMeta["subject"]     = sVal; break;
                case "description": dMeta["description"] = sVal; break;
                case "keywords":    dMeta["keywords"]    = sVal; break;
                // Ignore unknown keys silently.
            }
        }

        // Fallback title: first ATX heading in the document, or the
        // file basename if none is found.
        private static string firstHeadingOrBasename(string sMd, string sInPath)
        {
            foreach (var sLine in sMd.Split('\n')) {
                string s = sLine.TrimEnd('\r').Trim();
                if (s.Length == 0) continue;
                if (s.StartsWith("# ")) return s.Substring(2).Trim();
            }
            return Path.GetFileNameWithoutExtension(sInPath);
        }

        // Renders Markdown to HTML using Markdig with advanced
        // extensions plus YAML front matter support. The YAML
        // extension causes Markdig to recognize and STRIP the
        // leading --- ... --- block so it does not appear in the
        // rendered body. The Markdig assembly itself is embedded in
        // the EXE and loaded via the AssemblyResolve handler
        // registered at the top of Main.
        private static string renderMarkdown(string sMd)
        {
            var pipeline = new MarkdownPipelineBuilder()
                .UseAdvancedExtensions()
                .UseYamlFrontMatter()
                .Build();
            return Markdown.ToHtml(sMd, pipeline);
        }

        // Plain-text conversion. Renders the Markdown through
        // Markdig (so embedded HTML, links, tables, and image syntax
        // are normalized), then runs the resulting HTML through the
        // shared htmlBlocksToText helper. This produces readable
        // text with no image references — image lines in the source
        // markdown (![alt](path)) are rendered as <img> tags in HTML
        // and then stripped by the text pass.
        public static void convertToText(string sInPath, string sOutPath)
        {
            string sMd = File.ReadAllText(sInPath, Encoding.UTF8);
            string sHtml = renderMarkdown(sMd);
            string sText = htmlConverter.htmlBlocksToText(sHtml);
            File.WriteAllText(sOutPath, sText, new UTF8Encoding(true));
        }
    }

    // -----------------------------------------------------------------
    // JSON.
    // -----------------------------------------------------------------
    public static class jsonConverter
    {
        public static void convert(string sInPath, string sOutPath)
        {
            string sJson = File.ReadAllText(sInPath, Encoding.UTF8);
            var dMeta = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase) {
                { "title", Path.GetFileNameWithoutExtension(sInPath) }
            };

            object vRoot;
            try {
                var jsonParserInst = new jsonParser(sJson);
                vRoot = jsonParserInst.parseValue();
                jsonParserInst.skipWhitespace();
                if (!jsonParserInst.isAtEnd())
                    throw new FormatException("Unexpected content after root value at position " + jsonParserInst.pos);
            } catch (Exception ex) {
                using (var writer = new StreamWriter(sOutPath, false, new UTF8Encoding(false))) {
                    htmlWriter.writeHead(writer, dMeta);
                    writer.WriteLine("<main aria-labelledby=\"doc-title\">");
                    writer.WriteLine("<h2>Parse error</h2>");
                    writer.WriteLine("<p>" + htmlWriter.escape(ex.Message) + "</p>");
                    writer.WriteLine("<h2>Raw content</h2>");
                    writer.WriteLine("<pre>" + htmlWriter.escape(sJson) + "</pre>");
                    writer.WriteLine("</main>");
                    htmlWriter.writeFoot(writer);
                }
                return;
            }

            using (var writer = new StreamWriter(sOutPath, false, new UTF8Encoding(false))) {
                htmlWriter.writeHead(writer, dMeta);
                writer.WriteLine("<main aria-labelledby=\"doc-title\">");
                renderValue(writer, vRoot);
                writer.WriteLine("</main>");
                htmlWriter.writeFoot(writer);
            }
        }

        private static void renderValue(TextWriter writer, object oVal)
        {
            if (oVal == null) {
                writer.WriteLine("<span class=\"json-value null\">null</span>");
            } else if (oVal is bool b) {
                writer.WriteLine("<span class=\"json-value bool\">" + (b ? "true" : "false") + "</span>");
            } else if (oVal is double nD) {
                writer.WriteLine("<span class=\"json-value num\">" +
                    htmlWriter.escape(nD.ToString("R", CultureInfo.InvariantCulture)) + "</span>");
            } else if (oVal is long nL) {
                writer.WriteLine("<span class=\"json-value num\">" + nL + "</span>");
            } else if (oVal is string s) {
                writer.WriteLine("<span class=\"json-value str\">" + htmlWriter.escape(s) + "</span>");
            } else if (oVal is List<KeyValuePair<string, object>> lObj) {
                renderObject(writer, lObj);
            } else if (oVal is List<object> lArr) {
                renderArray(writer, lArr);
            } else {
                writer.WriteLine("<span class=\"json-value\">" +
                    htmlWriter.escape(Convert.ToString(oVal, CultureInfo.InvariantCulture)) + "</span>");
            }
        }

        private static void renderObject(TextWriter writer, List<KeyValuePair<string, object>> lObj)
        {
            if (lObj.Count == 0) {
                writer.WriteLine("<span class=\"json-value\">{} <em>(empty object)</em></span>");
                return;
            }
            writer.WriteLine("<dl class=\"json-obj\">");
            foreach (var kv in lObj) {
                writer.WriteLine("<dt>" + htmlWriter.escape(kv.Key) + ":</dt>");
                writer.WriteLine("<dd>");
                renderValue(writer, kv.Value);
                writer.WriteLine("</dd>");
            }
            writer.WriteLine("</dl>");
        }

        private static void renderArray(TextWriter writer, List<object> lArr)
        {
            if (lArr.Count == 0) {
                writer.WriteLine("<span class=\"json-value\">[] <em>(empty array)</em></span>");
                return;
            }
            writer.WriteLine("<ol class=\"json-arr\">");
            foreach (var v in lArr) {
                writer.WriteLine("<li>");
                renderValue(writer, v);
                writer.WriteLine("</li>");
            }
            writer.WriteLine("</ol>");
        }

        public class jsonParser
        {
            private readonly string src;
            public int pos;

            public jsonParser(string sSrc) { src = sSrc; pos = 0; }

            public bool isAtEnd() => pos >= src.Length;

            public void skipWhitespace()
            {
                while (pos < src.Length && char.IsWhiteSpace(src[pos])) pos++;
            }

            public object parseValue()
            {
                skipWhitespace();
                if (isAtEnd()) throw new FormatException("Unexpected end of input");
                char c = src[pos];
                if (c == '{') return parseObject();
                if (c == '[') return parseArray();
                if (c == '"') return parseString();
                if (c == '-' || (c >= '0' && c <= '9')) return parseNumber();
                if (c == 't' || c == 'f') return parseBool();
                if (c == 'n') { parseLiteral("null"); return null; }
                throw new FormatException("Unexpected character '" + c + "' at position " + pos);
            }

            private List<KeyValuePair<string, object>> parseObject()
            {
                var l = new List<KeyValuePair<string, object>>();
                pos++;
                skipWhitespace();
                if (pos < src.Length && src[pos] == '}') { pos++; return l; }
                while (true) {
                    skipWhitespace();
                    if (pos >= src.Length || src[pos] != '"')
                        throw new FormatException("Expected string key at position " + pos);
                    string sKey = parseString();
                    skipWhitespace();
                    if (pos >= src.Length || src[pos] != ':')
                        throw new FormatException("Expected ':' at position " + pos);
                    pos++;
                    object oVal = parseValue();
                    l.Add(new KeyValuePair<string, object>(sKey, oVal));
                    skipWhitespace();
                    if (pos < src.Length && src[pos] == ',') { pos++; continue; }
                    if (pos < src.Length && src[pos] == '}') { pos++; return l; }
                    throw new FormatException("Expected ',' or '}' at position " + pos);
                }
            }

            private List<object> parseArray()
            {
                var l = new List<object>();
                pos++;
                skipWhitespace();
                if (pos < src.Length && src[pos] == ']') { pos++; return l; }
                while (true) {
                    l.Add(parseValue());
                    skipWhitespace();
                    if (pos < src.Length && src[pos] == ',') { pos++; continue; }
                    if (pos < src.Length && src[pos] == ']') { pos++; return l; }
                    throw new FormatException("Expected ',' or ']' at position " + pos);
                }
            }

            private string parseString()
            {
                pos++;
                var sb = new StringBuilder();
                while (pos < src.Length) {
                    char c = src[pos++];
                    if (c == '"') return sb.ToString();
                    if (c == '\\') {
                        if (pos >= src.Length) throw new FormatException("Unterminated escape");
                        char e = src[pos++];
                        switch (e) {
                            case '"':  sb.Append('"'); break;
                            case '\\': sb.Append('\\'); break;
                            case '/':  sb.Append('/'); break;
                            case 'b':  sb.Append('\b'); break;
                            case 'f':  sb.Append('\f'); break;
                            case 'n':  sb.Append('\n'); break;
                            case 'r':  sb.Append('\r'); break;
                            case 't':  sb.Append('\t'); break;
                            case 'u':
                                if (pos + 4 > src.Length)
                                    throw new FormatException("Bad \\u escape at position " + pos);
                                string sHex = src.Substring(pos, 4);
                                pos += 4;
                                sb.Append((char)int.Parse(sHex, NumberStyles.HexNumber, CultureInfo.InvariantCulture));
                                break;
                            default: throw new FormatException("Unknown escape \\" + e + " at position " + (pos - 1));
                        }
                    } else sb.Append(c);
                }
                throw new FormatException("Unterminated string starting near position " + pos);
            }

            private object parseNumber()
            {
                int iStart = pos;
                if (src[pos] == '-') pos++;
                while (pos < src.Length && src[pos] >= '0' && src[pos] <= '9') pos++;
                bool bFrac = false;
                if (pos < src.Length && src[pos] == '.') {
                    bFrac = true;
                    pos++;
                    while (pos < src.Length && src[pos] >= '0' && src[pos] <= '9') pos++;
                }
                if (pos < src.Length && (src[pos] == 'e' || src[pos] == 'E')) {
                    bFrac = true;
                    pos++;
                    if (pos < src.Length && (src[pos] == '+' || src[pos] == '-')) pos++;
                    while (pos < src.Length && src[pos] >= '0' && src[pos] <= '9') pos++;
                }
                string sNum = src.Substring(iStart, pos - iStart);
                if (bFrac)
                    return double.Parse(sNum, CultureInfo.InvariantCulture);
                long nL;
                if (long.TryParse(sNum, NumberStyles.Integer, CultureInfo.InvariantCulture, out nL))
                    return nL;
                return double.Parse(sNum, CultureInfo.InvariantCulture);
            }

            private bool parseBool()
            {
                if (src[pos] == 't') { parseLiteral("true"); return true; }
                parseLiteral("false"); return false;
            }

            private void parseLiteral(string sExpected)
            {
                if (pos + sExpected.Length > src.Length ||
                    src.Substring(pos, sExpected.Length) != sExpected)
                    throw new FormatException("Expected literal '" + sExpected + "' at position " + pos);
                pos += sExpected.Length;
            }
        }

        // Plain-text conversion. JSON is already human-readable in
        // its source form; we simply copy the content as UTF-8
        // without BOM so it matches the output-encoding convention
        // used elsewhere.
        public static void convertToText(string sInPath, string sOutPath)
        {
            textPassthrough.copy(sInPath, sOutPath);
        }
    }

    // -----------------------------------------------------------------
    // Plain text.
    // -----------------------------------------------------------------
    public static class textConverter
    {
        public static void convert(string sInPath, string sOutPath)
        {
            string sContent = File.ReadAllText(sInPath);
            var dMeta = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase) {
                { "title", Path.GetFileNameWithoutExtension(sInPath) }
            };
            using (var writer = new StreamWriter(sOutPath, false, new UTF8Encoding(false))) {
                htmlWriter.writeHead(writer, dMeta);
                writer.WriteLine("<main aria-labelledby=\"doc-title\">");
                foreach (var sPara in Regex.Split(sContent, @"\r?\n\s*\r?\n")) {
                    string s = sPara.Trim();
                    if (s.Length == 0) continue;
                    writer.WriteLine("<p>" + htmlWriter.escape(s).Replace("\n", "<br>") + "</p>");
                }
                writer.WriteLine("</main>");
                htmlWriter.writeFoot(writer);
            }
        }
    }

    // -----------------------------------------------------------------
    // Plain-text pass-through: copy an already-readable text file as
    // UTF-8 (without BOM) to the output path. Used for .txt, .md,
    // and .csv in plain-text mode since those formats are already
    // human-readable as-is.
    // -----------------------------------------------------------------
    public static class textPassthrough
    {
        public static void copy(string sInPath, string sOutPath)
        {
            string sContent = File.ReadAllText(sInPath);
            File.WriteAllText(sOutPath, sContent, new UTF8Encoding(true));
        }
    }

    // -----------------------------------------------------------------
    // Shell integration helper. Opens a folder in File Explorer via
    // the default shell handler, reusing an already-open Explorer
    // window on that same folder when one exists.
    //
    // The detection uses the Shell.Application COM object, whose
    // Windows() collection enumerates open shell windows (Explorer
    // folders and, historically, Internet Explorer). Each item's
    // Document.Folder.Self.Path exposes the file-system path of
    // the folder currently displayed.
    //
    // When a match is found we bring the window to the foreground
    // via user32!SetForegroundWindow. If the window is minimized
    // we also restore it with ShowWindow(SW_RESTORE). If no match
    // exists we launch a new Explorer via Process.Start — the same
    // behaviour as clicking a folder in the shell.
    // -----------------------------------------------------------------
    public static class shellHelper
    {
        [System.Runtime.InteropServices.DllImport("user32.dll")]
        private static extern bool SetForegroundWindow(IntPtr hWnd);

        [System.Runtime.InteropServices.DllImport("user32.dll")]
        private static extern bool ShowWindow(IntPtr hWnd, int nCmdShow);

        [System.Runtime.InteropServices.DllImport("user32.dll")]
        private static extern bool IsIconic(IntPtr hWnd);

        private const int iSwRestore = 9;

        public static void openFolderSmart(string sPath)
        {
            if (string.IsNullOrWhiteSpace(sPath)) return;
            string sTarget;
            try { sTarget = Path.GetFullPath(sPath).TrimEnd('\\', '/'); }
            catch { sTarget = sPath; }

            // Try to find an existing Explorer window showing this
            // folder. If we find one, bring it to the foreground
            // and return. If anything in the COM dance throws,
            // fall through to Process.Start — better to open a new
            // window than to fail silently.
            try {
                if (focusExistingExplorer(sTarget)) {
                    logger.info("View output: focused existing Explorer window on " +
                        sTarget);
                    return;
                }
            } catch (Exception ex) {
                logger.info("View output: could not enumerate Shell.Windows: " +
                    ex.Message);
            }

            try {
                System.Diagnostics.Process.Start("explorer.exe",
                    "\"" + sTarget + "\"");
                logger.info("View output: launched Explorer on " + sTarget);
            } catch (Exception ex) {
                logger.info("View output: could not launch Explorer: " + ex.Message);
            }
        }

        private static bool focusExistingExplorer(string sTarget)
        {
            Type oShellType = Type.GetTypeFromProgID("Shell.Application");
            if (oShellType == null) return false;
            dynamic oShell = null;
            dynamic oWindows = null;
            try {
                oShell = Activator.CreateInstance(oShellType);
                oWindows = oShell.Windows();
                int iCount = 0;
                try { iCount = (int)oWindows.Count; } catch { }
                for (int i = 0; i < iCount; i++) {
                    dynamic oWin = null;
                    try {
                        oWin = oWindows.Item(i);
                        if (oWin == null) continue;

                        // Skip Internet Explorer windows. A real
                        // Explorer folder window has FullName
                        // ending in "explorer.exe"; an IE window
                        // ends in "iexplore.exe". This is the
                        // simplest way to tell them apart.
                        string sFullName = null;
                        try {
                            sFullName = Convert.ToString(oWin.FullName,
                                System.Globalization.CultureInfo.InvariantCulture);
                        } catch { }
                        if (string.IsNullOrEmpty(sFullName) ||
                            !sFullName.EndsWith("explorer.exe",
                                StringComparison.OrdinalIgnoreCase))
                            continue;

                        // Get the folder path this window is
                        // displaying. Older docs suggest
                        // LocationURL; the Path on Document.Folder.Self
                        // is more reliable for file-system folders.
                        string sWinPath = null;
                        dynamic oDoc = null, oFolder = null, oSelf = null;
                        try {
                            oDoc = oWin.Document;
                            oFolder = oDoc.Folder;
                            oSelf = oFolder.Self;
                            sWinPath = Convert.ToString(oSelf.Path,
                                System.Globalization.CultureInfo.InvariantCulture);
                        } catch { }
                        finally {
                            if (oSelf != null)
                                try { System.Runtime.InteropServices.Marshal.ReleaseComObject(oSelf); } catch { }
                            if (oFolder != null)
                                try { System.Runtime.InteropServices.Marshal.ReleaseComObject(oFolder); } catch { }
                            if (oDoc != null)
                                try { System.Runtime.InteropServices.Marshal.ReleaseComObject(oDoc); } catch { }
                        }
                        if (string.IsNullOrEmpty(sWinPath)) continue;

                        string sNormalized;
                        try { sNormalized = Path.GetFullPath(sWinPath).TrimEnd('\\', '/'); }
                        catch { sNormalized = sWinPath; }

                        if (!string.Equals(sNormalized, sTarget,
                            StringComparison.OrdinalIgnoreCase))
                            continue;

                        // Match. Bring it to front.
                        IntPtr hwnd = IntPtr.Zero;
                        try { hwnd = new IntPtr((int)oWin.HWND); } catch { }
                        if (hwnd == IntPtr.Zero) return false;
                        if (IsIconic(hwnd)) ShowWindow(hwnd, iSwRestore);
                        SetForegroundWindow(hwnd);
                        return true;
                    } finally {
                        if (oWin != null)
                            try { System.Runtime.InteropServices.Marshal.ReleaseComObject(oWin); } catch { }
                    }
                }
            } finally {
                if (oWindows != null)
                    try { System.Runtime.InteropServices.Marshal.ReleaseComObject(oWindows); } catch { }
                if (oShell != null)
                    try { System.Runtime.InteropServices.Marshal.ReleaseComObject(oShell); } catch { }
            }
            return false;
        }
    }

    // -----------------------------------------------------------------
    // Per-user configuration file, %LOCALAPPDATA%\2htm\2htm.ini.
    //
    // This file is OPT-IN. 2htm reads it only when -u /
    // --use-configuration is given (or the matching checkbox is on
    // in the GUI), and writes it only on an OK-click of the GUI
    // dialog when that checkbox is checked. A user who never turns
    // on Use configuration will never have 2htm touch the file
    // system outside the conversions they explicitly request.
    //
    // Format: a simple key=value text file, one pair per line,
    // UTF-8 with BOM. No section headers. Lines starting with ';'
    // or empty lines are ignored on read. Keys match the GUI's
    // field names. Values:
    //   - source_files          : string (wildcard/path/folder)
    //   - output_directory      : string (absolute path or empty)
    //   - strip_images          : 0 or 1
    //   - plain_text            : 0 or 1
    //   - force_replacements    : 0 or 1
    //   - view_output           : 0 or 1
    //
    // Read semantics: each key fills the corresponding 2htm global
    // ONLY if that global wasn't set on the command line.
    // Command-line always wins.
    // -----------------------------------------------------------------
    public static class configManager
    {
        public static string getConfigDir()
        {
            string sAppData = Environment.GetFolderPath(
                Environment.SpecialFolder.LocalApplicationData);
            return Path.Combine(sAppData, program.sConfigDirName);
        }

        public static string getConfigPath()
        {
            return Path.Combine(getConfigDir(), program.sConfigFileName);
        }

        public static bool configExists()
        {
            try { return File.Exists(getConfigPath()); }
            catch { return false; }
        }

        // Deletes the saved configuration file and, if the 2htm
        // folder under %LOCALAPPDATA% is left empty afterwards,
        // removes the folder too. Invoked by the GUI's "Default
        // settings" button. Silently no-ops if nothing is present;
        // non-fatal on failure (logs the error but does not
        // throw), because the whole point is to clean up gracefully.
        public static void eraseAll()
        {
            string sDir = getConfigDir();
            string sPath = getConfigPath();
            try {
                if (File.Exists(sPath)) {
                    File.Delete(sPath);
                    logger.info("Deleted configuration file: " + sPath);
                }
            } catch (Exception ex) {
                logger.info("Could not delete configuration file " +
                    sPath + ": " + ex.Message);
            }
            try {
                if (Directory.Exists(sDir)) {
                    // Remove only if empty, to preserve any other
                    // content a future version of 2htm — or a
                    // curious user — might have placed there.
                    bool bEmpty = Directory.EnumerateFileSystemEntries(sDir)
                        .GetEnumerator().MoveNext() == false;
                    if (bEmpty) {
                        Directory.Delete(sDir);
                        logger.info("Removed empty configuration directory: " +
                            sDir);
                    }
                }
            } catch (Exception ex) {
                logger.info("Could not remove configuration directory " +
                    sDir + ": " + ex.Message);
            }
        }

        // Loads saved values into the program's globals. Called
        // before the GUI is shown (so saved values become dialog
        // defaults) and is a no-op if the file doesn't exist.
        // lsFileArgs is populated with a source-files entry only
        // when the user supplied nothing on the command line.
        public static void loadInto(List<string> lsFileArgs)
        {
            string sPath = getConfigPath();
            if (!File.Exists(sPath)) return;

            Dictionary<string, string> dVals;
            try {
                dVals = parseFile(sPath);
            } catch (Exception ex) {
                string sMsg = "Could not read configuration from:\r\n" +
                    sPath + "\r\n\r\n" + ex.Message;
                Console.Error.WriteLine("[WARN] " + sMsg);
                if (program.bGuiMode) {
                    try {
                        System.Windows.Forms.MessageBox.Show(sMsg,
                            "2htm — Configuration not loaded",
                            System.Windows.Forms.MessageBoxButtons.OK,
                            System.Windows.Forms.MessageBoxIcon.Warning);
                    } catch { }
                }
                return;
            }

            // Source files: only if the CLI had no file args at all.
            if (!program.bSourceFromCli) {
                string sSaved;
                if (dVals.TryGetValue("source_files", out sSaved) &&
                    !string.IsNullOrWhiteSpace(sSaved)) {
                    foreach (var sArg in program.splitSourceField(sSaved))
                        lsFileArgs.Add(sArg);
                }
            }

            if (!program.bOutputDirFromCli)
                program.sOutputDir = getOrEmpty(dVals, "output_directory");
            if (!program.bStripImagesFromCli)
                program.bStripImages = getBool(dVals, "strip_images");
            if (!program.bPlainTextFromCli)
                program.bPlainText = getBool(dVals, "plain_text");
            if (!program.bForceFromCli)
                program.bForce = getBool(dVals, "force_replacements");
            if (!program.bViewOutputFromCli)
                program.bViewOutput = getBool(dVals, "view_output");
            if (!program.bLogFromCli)
                program.bLog = getBool(dVals, "log_session");
        }

        public static void save(string sSource, string sOutputDir,
            bool bStrip, bool bPlain, bool bForce, bool bView, bool bLog)
        {
            string sDir = getConfigDir();
            string sPath = getConfigPath();
            try {
                if (!Directory.Exists(sDir)) Directory.CreateDirectory(sDir);
                var sb = new StringBuilder();
                sb.AppendLine("; 2htm configuration");
                sb.AppendLine("; auto-written on OK-click when Use configuration was checked.");
                sb.AppendLine("; Delete this file to reset, or click Default settings in");
                sb.AppendLine("; the GUI, which also deletes the file and the 2htm folder.");
                sb.AppendLine("source_files=" + (sSource ?? ""));
                sb.AppendLine("output_directory=" + (sOutputDir ?? ""));
                sb.AppendLine("strip_images=" + (bStrip ? "1" : "0"));
                sb.AppendLine("plain_text=" + (bPlain ? "1" : "0"));
                sb.AppendLine("force_replacements=" + (bForce ? "1" : "0"));
                sb.AppendLine("view_output=" + (bView ? "1" : "0"));
                sb.AppendLine("log_session=" + (bLog ? "1" : "0"));
                File.WriteAllText(sPath, sb.ToString(), new UTF8Encoding(true));
                logger.info("Saved configuration to " + sPath);
            } catch (Exception ex) {
                // Writing the config is a convenience; don't fail
                // the whole run if it can't be written, but do
                // surface the failure to a GUI user — otherwise
                // the "Use configuration" feature silently fails
                // to persist and the user has no way to tell.
                string sMsg = "Could not save configuration to:\r\n" +
                    sPath + "\r\n\r\n" + ex.Message;
                Console.Error.WriteLine("[WARN] " + sMsg);
                logger.info("Could not save configuration: " + ex.Message);
                if (program.bGuiMode) {
                    try {
                        System.Windows.Forms.MessageBox.Show(sMsg,
                            "2htm — Configuration not saved",
                            System.Windows.Forms.MessageBoxButtons.OK,
                            System.Windows.Forms.MessageBoxIcon.Warning);
                    } catch { }
                }
            }
        }

        private static Dictionary<string, string> parseFile(string sPath)
        {
            var d = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);
            foreach (var sLineRaw in File.ReadAllLines(sPath)) {
                string sLine = sLineRaw.Trim();
                if (sLine.Length == 0) continue;
                if (sLine.StartsWith(";") || sLine.StartsWith("#")) continue;
                if (sLine.StartsWith("[") && sLine.EndsWith("]")) continue;
                int iEq = sLine.IndexOf('=');
                if (iEq <= 0) continue;
                string sKey = sLine.Substring(0, iEq).Trim();
                string sVal = sLine.Substring(iEq + 1).Trim();
                d[sKey] = sVal;
            }
            return d;
        }

        private static bool getBool(Dictionary<string, string> d, string sKey)
        {
            string s;
            if (!d.TryGetValue(sKey, out s)) return false;
            s = (s ?? "").Trim();
            return s == "1" || s.Equals("true", StringComparison.OrdinalIgnoreCase) ||
                s.Equals("yes", StringComparison.OrdinalIgnoreCase);
        }

        private static string getOrEmpty(Dictionary<string, string> d, string sKey)
        {
            string s;
            if (!d.TryGetValue(sKey, out s)) return "";
            return s ?? "";
        }
    }

}
