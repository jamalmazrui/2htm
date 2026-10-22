// =====================================================================
// 2htm.cs - Convert documents to accessible HTML.
// Copyright (c) 2026 Jamal Mazrui — MIT License — https://github.com/JamalMazrui/2htm
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
// Build with build.cmd.
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
using Homer;

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
        // From version.txt, through the Version.cs the build writes on
        // every build, so -v, the log, F11 and the installer report one
        // number.
        public const string sProgramVersion = BuildVersion.Version;
        public const string sConfigDirName = "2htm";
        public const string sConfigFileName = "2htm.inix";
        // The GitHub owner whose 2htm releases F11 asks about, and the
        // settings file of the versions before 1.19.
        public const string sGitHubOwner = "jamalmazrui";
        public const string sLegacyConfigFileName = "2htm.ini";
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
        // %LOCALAPPDATA%\2htm\configs\2htm.inix on startup (before the GUI
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
            // ONE SESSION LOG, ALWAYS: %LOCALAPPDATA%\2htm\logs\
            // 2htm-yyyyMMdd-HHmmss.log, from Homer's Log, with the
            // environment already in its header. -l still writes 2htm.log
            // beside the output as well. Neither Log nor Elevate touches
            // Markdig, so Main still references no embedded assembly.
            Log.start(sProgramName);
            Elevate.configure(sGitHubOwner, sProgramName, sProgramVersion);
            int iExit = iExitFatal;
            try {
                iExit = run(asArgs);
            } catch (Exception ex) {
                // Whatever run() did not catch is recorded with its
                // stack, and the console names the log rather than
                // showing a trace.
                Log.exception(ex);
                Console.Error.WriteLine("2htm stopped: " + ex.Message);
                Console.Error.WriteLine("The session log has the details: " + Log.path);
                iExit = iExitFatal;
            }
            Log.info("Exit code " + iExit);
            Log.close();
            return iExit;
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
            } else if (bOwnConsole && !Console.IsOutputRedirected) {
                // CLI launched by Explorer (context-menu verb).
                // Hide the console flash and mark that we should
                // report any conversion errors via MessageBox.
                bHideConsoleMode = true;
                hideOwnConsoleWindow();
            }
            // NEVER A DIALOG WHEN SOMEBODY IS READING THE OUTPUT.
            //
            // A program that started 2htm and redirected its output is waiting
            // to read that output, not to click OK. HomerView launches it with
            // no window and both streams redirected; the console-owner test
            // above said "Explorer launched me", so an ordinary error went up
            // as a MODAL MESSAGE BOX ON AN INVISIBLE WINDOW -- unseeable,
            // undismissable, and it hung the caller until it was killed.
            //
            // Console.IsOutputRedirected is the honest test: it is true exactly
            // when another program is on the other end of the pipe. If it is,
            // errors go to the stream where that program will actually find
            // them.

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
                // ONE SWITCH CONVENTION (20 October 2026, homer-code): /switch, -switch and --switch
                // alike, in any case, through the kit's Util.switchName; the single letters still
                // work. A known switch is always a switch: on Windows /help is also a path, the Help
                // folder at the drive's root, and 2htm took it for a source where that folder exists.
                // Only an unknown name that is an existing file or folder is a source.
                string sSwitch = Homer.Util.switchName(sArg);
                string[] aKnown = { "f", "force", "g", "gui", "gui-mode", "h", "help", "l", "log", "o", "output", "output-dir",
                    "p", "plain-text", "s", "strip-images", "u", "use-configuration", "v", "version", "view", "view-output" };
                if (sSwitch != "" && Array.IndexOf(aKnown, sSwitch) < 0 && (File.Exists(sArg) || Directory.Exists(sArg))) sSwitch = "";
                if (sSwitch == "help" || sSwitch == "h" || sArg == "/?" || sArg == "-?") {
                    printUsage();
                    return iExitOk;
                }
                if (sSwitch == "version" || sSwitch == "v") {
                    Console.WriteLine(program.sProgramName + " " + sProgramVersion);
                    return iExitOk;
                }
                if (sSwitch == "strip-images" || sSwitch == "s") {
                    bStripImages = true; bStripImagesFromCli = true;
                    continue;
                }
                if (sSwitch == "plain-text" || sSwitch == "p") {
                    bPlainText = true; bPlainTextFromCli = true;
                    continue;
                }
                if (sSwitch == "force" || sSwitch == "f") {
                    bForce = true; bForceFromCli = true;
                    continue;
                }
                if (sSwitch == "log" || sSwitch == "l") {
                    bLog = true; bLogFromCli = true;
                    continue;
                }
                if (sSwitch == "gui-mode" || sSwitch == "gui" || sSwitch == "g") {
                    bGuiMode = true;
                    continue;
                }
                if (sSwitch == "view-output" || sSwitch == "view") {
                    bViewOutput = true; bViewOutputFromCli = true;
                    continue;
                }
                if (sSwitch == "use-configuration" || sSwitch == "u") {
                    bUseConfig = true;
                    continue;
                }
                if (sSwitch == "output-dir" || sSwitch == "output" || sSwitch == "o") {
                    // Expects a following argument giving the
                    // directory. If missing or empty, treat as fatal
                    // so the user learns the correct usage.
                    if (i + 1 >= asArgs.Length ||
                        string.IsNullOrWhiteSpace(asArgs[i + 1])) {
                        Console.Error.WriteLine(sArg +
                            " requires a directory argument.");
                        Console.Error.WriteLine("Run '2htm /help' for usage.");
                        return iExitFatal;
                    }
                    sOutputDir = asArgs[++i];
                    bOutputDirFromCli = true;
                    continue;
                }
                if (sSwitch != "") {
                    Console.Error.WriteLine("Unknown switch: " + sArg);
                    Console.Error.WriteLine("Run '2htm /help' for usage.");
                    return iExitFatal;
                }
                lsFileArgs.Add(sArg);
                bSourceFromCli = true;
            }

            // Configuration file (opt-in, per-user INI at
            // %LOCALAPPDATA%\2htm\configs\2htm.inix):
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
                    Console.Error.WriteLine(sErr);
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
                    Console.WriteLine("No matching files.");
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
                    iExitCode = (iFailed == 0 && iInputErrors == 0) ? iExitOk : iExitPartial;

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
            Console.WriteLine(program.sProgramName + " " + sProgramVersion);
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
            Console.WriteLine("  2htm /plain-text C:\\docs\\*.pdf reports\\*.docx");
            Console.WriteLine("  2htm /force /strip-images slides\\*.pptx");
            Console.WriteLine();
            Console.WriteLine("Switches, written /switch, -switch or --switch, in any case:");
            Console.WriteLine("  /help, /?            Show this help and exit.");
            Console.WriteLine("  /version             Show the version number and exit.");
            Console.WriteLine("  /force               Overwrite the target file if it already");
            Console.WriteLine("                       exists. Default is to skip.");
            Console.WriteLine("  /strip-images        Strip images from the output. No <img>");
            Console.WriteLine("                       tags, no broken references, no alt text.");
            Console.WriteLine("                       Use for large image-heavy documents.");
            Console.WriteLine("  /plain-text          Produce .txt instead of .htm. Output is");
            Console.WriteLine("                       readable plain text (no HTML markup).");
            Console.WriteLine("  /output-dir          <dir>, /output <dir>");
            Console.WriteLine("                       Write output files to <dir> instead of");
            Console.WriteLine("                       the current directory. <dir> is created");
            Console.WriteLine("                       if it does not exist.");
            Console.WriteLine("  /gui-mode            Launch a dialog to enter parameters. Any");
            Console.WriteLine("                       options given on the command line become");
            Console.WriteLine("                       the defaults. After conversion, a message");
            Console.WriteLine("                       box displays the per-file results. GUI mode");
            Console.WriteLine("                       is also entered automatically when 2htm is");
            Console.WriteLine("                       launched with no arguments from a GUI shell");
            Console.WriteLine("                       such as File Explorer (e.g., by double-");
            Console.WriteLine("                       clicking the executable or extracting it");
            Console.WriteLine("                       from a zip archive and running it).");
            Console.WriteLine("  /view-output         After conversion, open the output");
            Console.WriteLine("                       directory in the default file manager");
            Console.WriteLine("                       (typically File Explorer). Fires only if");
            Console.WriteLine("                       at least one file was actually converted;");
            Console.WriteLine("                       reuses an already-open window if one is");
            Console.WriteLine("                       already displaying that directory.");
            Console.WriteLine("  /use-configuration");
            Console.WriteLine("                       Read saved defaults from");
            Console.WriteLine("                       %LOCALAPPDATA%\\2htm\\configs\\2htm.inix. Any other");
            Console.WriteLine("                       options supplied on the command line take");
            Console.WriteLine("                       precedence over saved values. In GUI mode,");
            Console.WriteLine("                       the Use configuration checkbox also");
            Console.WriteLine("                       controls whether the current dialog values");
            Console.WriteLine("                       are written back as next-run defaults.");
            Console.WriteLine("                       Without -u (and no checkbox), no");
            Console.WriteLine("                       settings are saved.");
            Console.WriteLine("  /log                 Also write detailed diagnostics to 2htm.log");
            Console.WriteLine("                       (UTF-8 with BOM) in the output directory if");
            Console.WriteLine("                       one was specified, or the current working");
            Console.WriteLine("                       directory otherwise. Any prior 2htm.log in");
            Console.WriteLine("                       that location is overwritten so the file");
            Console.WriteLine("                       always reflects only the current session.");
            Console.WriteLine("                       Useful for reporting conversion problems.");
            Console.WriteLine("                       A session log is always kept in");
            Console.WriteLine("                       %LOCALAPPDATA%\\2htm\\logs.");
            Console.WriteLine();
            Console.WriteLine("Supported input formats:");
            Console.WriteLine("  Word:       .docx .doc .rtf .odt");
            Console.WriteLine("  PDF:        .pdf   (via Word 2013+ PDF Reflow)");
            Console.WriteLine("  Excel:      .xlsx .xls");
            Console.WriteLine("  CSV:        .csv   (native)");
            Console.WriteLine("  PowerPoint: .pptx .ppt");
            Console.WriteLine("  Web:        .html .htm (native)");
            Console.WriteLine("  Markdown:   .md    (through the Homer kit's conversion)");
            Console.WriteLine("  JSON:       .json  (native)");
            Console.WriteLine("  Text:       .txt   (native)");
            Console.WriteLine();
            Console.WriteLine("Output files are written to the current working directory by");
            Console.WriteLine("default (.htm, or .txt when /plain-text is given), or to the");
            Console.WriteLine("directory given by -o / /output-dir if one is specified. Existing");
            Console.WriteLine("output files are skipped unless /force is given. The tool never");
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

        // INPUTS THAT COULD NOT BE USED (8 October 2026, from an audit by another AI):
        // a missing file, an empty match or a collision was reported and then forgotten,
        // so a batch with one good file and one missing one still ended in success.
        static int iInputErrors = 0;

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
                    Console.Error.WriteLine("Directory not found: " + sDir); iInputErrors = iInputErrors + 1;
                    continue;
                }
                if (sPattern.IndexOfAny(new[] { '*', '?' }) < 0) {
                    string sFull = Path.GetFullPath(sArg);
                    if (File.Exists(sFull) && !isLockFile(sFull)) hsResult.Add(sFull);
                    else if (!File.Exists(sFull))
                        Console.Error.WriteLine("File not found: " + sArg); iInputErrors = iInputErrors + 1;
                } else {
                    string[] aMatches;
                    try {
                        aMatches = Directory.GetFiles(sDir, sPattern);
                    } catch (Exception ex) {
                        Console.Error.WriteLine("Cannot enumerate '" + sArg + "': " + ex.Message); iInputErrors = iInputErrors + 1;
                        continue;
                    }
                    if (aMatches.Length == 0)
                        Console.Error.WriteLine("No files match: " + sArg); iInputErrors = iInputErrors + 1;
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
            // BESIDE ITS SOURCE, AS THE GUIDE SAYS (8 October 2026, from an audit by another
            // AI): with no output folder given, output went to the current folder, so a
            // shortcut started in Documents put it there, away from the file converted.
            string sEffectiveOutputDir = string.IsNullOrEmpty(sOutputDir)
                ? Path.GetDirectoryName(Path.GetFullPath(sInPath))
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
            // ONE OUTPUT, ONE SOURCE (8 October 2026, from an audit by another AI):
            // report.docx and report.md both became report.htm, and the later one
            // silently replaced the first; and a conversion could replace another
            // input. Each destination is now claimed once, and never another input.
            var dClaimed = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);
            var setInputs = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
            foreach (var sOne in lsFiles) { try { setInputs.Add(Path.GetFullPath(sOne)); } catch { } }
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
                string sOutFull = Path.GetFullPath(sOutPath);
                if (dClaimed.ContainsKey(sOutFull)) {
                    logger.info("Skipped (" + Path.GetFileName(sOutPath) + " is already the output of " + dClaimed[sOutFull] + "): " + sFile);
                    Console.Error.WriteLine("Skipped " + sFile + ": its output, " + Path.GetFileName(sOutPath) + ", is already the output of " + dClaimed[sOutFull] + ".");
                    iInputErrors = iInputErrors + 1;
                    continue;
                }
                if (setInputs.Contains(sOutFull)) {
                    logger.info("Skipped (its output " + sOutPath + " is another file being converted): " + sFile);
                    Console.Error.WriteLine("Skipped " + sFile + ": its output would replace " + Path.GetFileName(sOutPath) + ", which is also being converted.");
                    iInputErrors = iInputErrors + 1;
                    continue;
                }
                dClaimed[sOutFull] = Path.GetFileName(sFile);
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
            // THE OLD OUTPUT IS KEPT UNTIL THE NEW ONE EXISTS (8 October 2026, from an
            // audit by another AI): on any failure the destination was deleted, and with
            // --force that was the previous good output. It is moved aside first, put
            // back if this conversion fails, and removed only after success.
            string sAside = "";
            if (File.Exists(sOutPath)) {
                // A RECOVERY NAME OF ITS OWN (14 October 2026, from an audit by another
                // AI): it was always "<output>.previous", and an existing one was
                // deleted first, so a second failed conversion erased the first one's
                // recovery copy. Each conversion now gets a name that does not exist yet.
                sAside = Homer.Util.recoveryPath(sOutPath);
                File.Move(sOutPath, sAside);
            }
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
                    if (sAside != "" && File.Exists(sAside) && !File.Exists(sOutPath)) {
                        try { File.Move(sAside, sOutPath); logger.info("Put back the earlier output: " + sOutPath); }
                        catch (Exception ex) { logger.warn("Could not put back the earlier output; it is kept at " + sAside + ": " + ex.Message); }
                    }
                } else if (sAside != "") {
                    try { File.Delete(sAside); } catch (Exception ex) { logger.info("Could not remove " + sAside + ": " + ex.Message); }
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

        // convertOneToHtml and convertOneToText: THE HOMER KIT'S ONE CONVERSION (20 October 2026).
        // 2htm is a face over Conversion.convert, as every Homer app converts: the kit's route
        // table chooses the way for each type, and every web page ends in the Homer page form,
        // which came from 2htm. --strip-images is passed to the kit's page writer.
        private static void convertOneToHtml(string sExt, string sInPath, string sOutPath)
        {
            Homer.PageWriter.bStripImages = bStripImages;
            string sError;
            if (!Homer.Conversion.convert(sInPath, sOutPath, out sError))
                throw new InvalidOperationException(sError == "" ? "the conversion failed" : sError);
        }

        private static void convertOneToText(string sExt, string sInPath, string sOutPath)
        {
            string sError;
            if (!Homer.Conversion.convert(sInPath, sOutPath, out sError))
                throw new InvalidOperationException(sError == "" ? "the conversion failed" : sError);
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
        // THE DIALOG IS BUILT FROM Homer's Lbc (since 1.19.1), like every
        // Homer dialog: one control per row in tab order, each field's label
        // just before it so the label names it, a band where a field and the
        // button that fills it share a row. Lbc supplies, with no code here,
        // Control+Enter for OK from any control, Shift+F1 for a field's tip,
        // F7 for a list of the dialog's controls, the editing keys of every
        // text field, and the Help button (F1), which lists each field with
        // its tip and ends with the version check. F11 is claimed below. The
        // Guide button opens the full guide and returns to the dialog with
        // everything as it was, as Default settings does.
        public static bool show(ref string sSource, ref string sOutputDir,
            ref bool bStrip, ref bool bPlain, ref bool bForce, ref bool bView,
            ref bool bLog, ref bool bUseCfg)
        {
            while (true) {
                string sButton;
                using (var dlg = new LbcDialog(program.sProgramName, null)) {
                    string sStartSource = string.IsNullOrWhiteSpace(sSource) ? defaultSourceForGui() : sSource;
                    dlg.addBand();
                    System.Windows.Forms.TextBox tbSource = dlg.addInputBox("&Source files:", sStartSource,
                        "One or more files, folders or wildcard patterns such as *.docx, " +
                        "separated by spaces. A folder means every supported file in it.");
                    System.Windows.Forms.Button btnBrowse = dlg.addButton("&Browse source...",
                        "Choose a folder of files to convert.");
                    dlg.addBand();
                    System.Windows.Forms.TextBox tbOut = dlg.addInputBox("&Output directory:",
                        !string.IsNullOrWhiteSpace(sOutputDir) ? sOutputDir : deriveOutputDirFromSource(sStartSource),
                        "Where each converted file is written. Blank means beside the source.");
                    System.Windows.Forms.Button btnChoose = dlg.addButton("&Choose output...",
                        "Choose the folder the converted files go to.");
                    dlg.endBand();

                    btnBrowse.Click += (o, e) => {
                        string sPicked = pickFolder("Choose a folder of files to convert", tbSource.Text);
                        if (sPicked == null) return;
                        // The output follows the source only while it still
                        // mirrors it; a folder the user chose stays.
                        string sOldDerived = deriveOutputDirFromSource(tbSource.Text);
                        bool bOutputMirrorsSource = string.IsNullOrWhiteSpace(tbOut.Text) ||
                            string.Equals(tbOut.Text.Trim(), sOldDerived, StringComparison.OrdinalIgnoreCase);
                        tbSource.Text = sPicked;
                        if (bOutputMirrorsSource) tbOut.Text = deriveOutputDirFromSource(sPicked);
                        tbSource.Focus();
                    };
                    btnChoose.Click += (o, e) => {
                        string sPicked = pickFolder("Choose the output directory", tbOut.Text);
                        if (sPicked == null) return;
                        tbOut.Text = sPicked;
                        tbOut.Focus();
                    };

                    dlg.addSeparator();
                    System.Windows.Forms.CheckBox cbStrip = dlg.addCheckBox("Strip &images", bStrip,
                        "Leave images out of the output.");
                    System.Windows.Forms.CheckBox cbPlain = dlg.addCheckBox("&Plain text", bPlain,
                        "Write plain text, .txt, instead of HTML.");
                    System.Windows.Forms.CheckBox cbForce = dlg.addCheckBox("&Force replacements", bForce,
                        "Overwrite an existing output file instead of skipping the input.");
                    System.Windows.Forms.CheckBox cbView = dlg.addCheckBox("&View output", bView,
                        "Open the output folder in File Explorer when the run is done.");
                    System.Windows.Forms.CheckBox cbLog = dlg.addCheckBox("&Log session", bLog,
                        "Also write " + program.sLogFileName + " in the output folder. A session log is always kept in %LOCALAPPDATA%\\2htm\\logs.");
                    System.Windows.Forms.CheckBox cbUseCfg = dlg.addCheckBox("&Use configuration", bUseCfg,
                        "Load these settings next time, and save them when you press OK, in configs\\" + program.sConfigFileName + ".");

                    // F11: is there a newer 2htm on the web? (Elevate sounds
                    // like eleven.) Claimed before any control sees the key.
                    dlg.commandKey = k => {
                        if (k != System.Windows.Forms.Keys.F11) return false;
                        logger.info("F11: checking the web for a newer version");
                        if (Elevate.offer(dlg.form)) {
                            logger.info("F11: the setup program was started; closing the dialog");
                            dlg.form.DialogResult = System.Windows.Forms.DialogResult.Cancel;
                            dlg.form.Close();
                        }
                        return true;
                    };

                    sButton = dlg.runWithButtons(new string[] {
                        "OK", "Guide", "Default settings", "Cancel" });

                    // Harvest every field before the dialog is disposed, so a
                    // return to the dialog keeps whatever else was typed.
                    sSource = (tbSource.Text ?? "").Trim();
                    sOutputDir = (tbOut.Text ?? "").Trim();
                    bStrip = cbStrip.Checked;
                    bPlain = cbPlain.Checked;
                    bForce = cbForce.Checked;
                    bView = cbView.Checked;
                    bLog = cbLog.Checked;
                    bUseCfg = cbUseCfg.Checked;
                }

                if (string.IsNullOrEmpty(sButton) ||
                        string.Equals(sButton, "Cancel", StringComparison.OrdinalIgnoreCase))
                    return false;
                if (string.Equals(sButton, "Guide", StringComparison.OrdinalIgnoreCase)) {
                    launchReadMe();
                    continue;
                }
                if (string.Equals(sButton, "Default settings", StringComparison.OrdinalIgnoreCase)) {
                    sSource = defaultSourceForGui();
                    sOutputDir = deriveOutputDirFromSource(sSource);
                    bStrip = false; bPlain = false; bForce = false;
                    bView = false; bLog = false; bUseCfg = false;
                    configManager.eraseAll();
                    logger.info("Default settings restored");
                    continue;
                }
                // OK: offer to create an output directory that does not exist
                // yet; a blank one is derived from the source, as before.
                if (!confirmOutputDir(sOutputDir)) continue;
                if (string.IsNullOrEmpty(sOutputDir))
                    sOutputDir = deriveOutputDirFromSource(sSource);
                return true;
            }
        }

        // Offer to create a non-existent output directory, default Yes.
        // Returns false to send the user back to the dialog.
        private static bool confirmOutputDir(string sOutputDir)
        {
            string sOutCandidate = (sOutputDir ?? "").Trim();
            if (sOutCandidate.Length >= 2 && sOutCandidate[0] == '"' && sOutCandidate[sOutCandidate.Length - 1] == '"')
                sOutCandidate = sOutCandidate.Substring(1, sOutCandidate.Length - 2).Trim();
            if (string.IsNullOrEmpty(sOutCandidate)) return true;
            try { if (System.IO.Directory.Exists(sOutCandidate)) return true; } catch { return true; }
            System.Windows.Forms.DialogResult dr = System.Windows.Forms.MessageBox.Show(
                "Create " + sOutCandidate + "?", program.sProgramName,
                System.Windows.Forms.MessageBoxButtons.YesNo,
                System.Windows.Forms.MessageBoxIcon.Question,
                System.Windows.Forms.MessageBoxDefaultButton.Button1);
            if (dr != System.Windows.Forms.DialogResult.Yes) return false;
            try {
                System.IO.Directory.CreateDirectory(sOutCandidate);
                return true;
            } catch (Exception ex) {
                System.Windows.Forms.MessageBox.Show(
                    "Could not create directory:\r\n" + sOutCandidate + "\r\n\r\n" + ex.Message,
                    program.sProgramName,
                    System.Windows.Forms.MessageBoxButtons.OK,
                    System.Windows.Forms.MessageBoxIcon.Warning);
                return false;
            }
        }

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

        // Opens the full guide. The Homer layout puts it in
        // help\2htm.htm, beside exec\ where the program runs, in the
        // installed tree and in the project alike; Paths.installedFolder
        // finds that folder. ReadMe.htm at the top is the fallback, then
        // the Markdown forms.
        private static void launchReadMe()
        {
            string sFolder = Paths.installedFolder;
            string[] aCandidates = {
                Path.Combine(Path.Combine(sFolder, "help"), "2htm.htm"),
                Path.Combine(sFolder, "ReadMe.htm"),
                Path.Combine(Path.Combine(sFolder, "help"), "2htm.md"),
                Path.Combine(sFolder, "ReadMe.md") };
            string sTarget = null;
            foreach (string sCandidate in aCandidates)
                if (sTarget == null && File.Exists(sCandidate)) sTarget = sCandidate;
            if (sTarget == null) {
                logger.warn("No guide found under " + sFolder);
                System.Windows.Forms.MessageBox.Show(
                    "The guide was not found in the 2htm folder:\r\n\r\n" +
                    sFolder + "\r\n\r\nReinstalling 2htm puts it back.",
                    "2htm — Guide not found",
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
                    "Could not open the guide:\r\n\r\n" + ex.Message,
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
                Console.Error.WriteLine("Could not open log file '" + sPath + "': " +
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

        // EVERY LINE ALSO GOES TO THE SESSION LOG that Homer's Log keeps
        // in %LOCALAPPDATA%\2htm\logs, whatever the options, so a
        // failure always leaves a record. 2htm.log (-l) stays what it
        // was: an extra copy beside the output.
        public static void info(string sMsg)  { Log.info(sMsg); write("INFO", sMsg); }
        public static void warn(string sMsg)  { Log.warn(sMsg); write("WARN", sMsg); }
        public static void error(string sMsg) { Log.error(sMsg); write("ERROR", sMsg); }
        public static void debug(string sMsg) { Log.line("DEBUG  " + sMsg); write("DEBUG", sMsg); }

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
            Log.section("Settings");
            if (dParams != null) foreach (var oKv in dParams) Log.keyValue(oKv.Key, oKv.Value);
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

    // 2htm's own converters -- htmlWriter, comHelper and the Word, Excel, PowerPoint, CSV, HTML,
    // Markdown, JSON and text converters -- were retired on 20 October 2026 into
    // notes\retired\2htmConverters.cs. Every conversion is now the Homer kit's Conversion.convert,
    // the one technique every Homer app uses, and the page form is the kit's ConversionPage.


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
    // Per-user configuration file, %LOCALAPPDATA%\2htm\configs\2htm.inix.
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
        // Since 1.19 the settings live in the [Settings] section of
        // %LOCALAPPDATA%\2htm\configs\2htm.inix (Homer's Paths.configs(),
        // written through Homer's InixCodec, which keeps any comment a
        // person adds). The 2htm.ini of earlier versions, directly under
        // %LOCALAPPDATA%\2htm, is read while no .inix exists, its
        // snake_case names read as today's, and it is removed the first
        // time the new file is saved.
        const string sSettingsSection = "Settings";

        public static string getConfigDir()
        {
            return Paths.configs();
        }

        public static string getConfigPath()
        {
            return Path.Combine(getConfigDir(), program.sConfigFileName);
        }

        public static string getLegacyPath()
        {
            string sAppData = Environment.GetFolderPath(
                Environment.SpecialFolder.LocalApplicationData);
            return Path.Combine(Path.Combine(sAppData, program.sConfigDirName), program.sLegacyConfigFileName);
        }

        public static bool configExists()
        {
            try { return File.Exists(getConfigPath()) || File.Exists(getLegacyPath()); }
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
            // Both the settings file and one left by a version before
            // 1.19. The folders stay: %LOCALAPPDATA%\2htm also holds the
            // session logs.
            foreach (string sPath in new string[] { getConfigPath(), getLegacyPath() }) {
                try {
                    if (File.Exists(sPath)) {
                        File.Delete(sPath);
                        logger.info("Deleted configuration file: " + sPath);
                    }
                } catch (Exception ex) {
                    logger.info("Could not delete configuration file " +
                        sPath + ": " + ex.Message);
                }
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
            if (!File.Exists(sPath)) sPath = getLegacyPath();
            if (!File.Exists(sPath)) return;
            logger.info("Reading configuration from " + sPath);

            Dictionary<string, string> dVals;
            try {
                dVals = parseFile(sPath);
            } catch (Exception ex) {
                string sMsg = "Could not read configuration from:\r\n" +
                    sPath + "\r\n\r\n" + ex.Message;
                Console.Error.WriteLine(sMsg);
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
                if (dVals.TryGetValue("SourceFiles", out sSaved) &&
                    !string.IsNullOrWhiteSpace(sSaved)) {
                    foreach (var sArg in program.splitSourceField(sSaved))
                        lsFileArgs.Add(sArg);
                }
            }

            if (!program.bOutputDirFromCli)
                program.sOutputDir = getOrEmpty(dVals, "OutputDirectory");
            if (!program.bStripImagesFromCli)
                program.bStripImages = getBool(dVals, "StripImages");
            if (!program.bPlainTextFromCli)
                program.bPlainText = getBool(dVals, "PlainText");
            if (!program.bForceFromCli)
                program.bForce = getBool(dVals, "ForceReplacements");
            if (!program.bViewOutputFromCli)
                program.bViewOutput = getBool(dVals, "ViewOutput");
            if (!program.bLogFromCli)
                program.bLog = getBool(dVals, "LogSession");
        }

        public static void save(string sSource, string sOutputDir,
            bool bStrip, bool bPlain, bool bForce, bool bView, bool bLog)
        {
            string sDir = getConfigDir();
            string sPath = getConfigPath();
            try {
                if (!Directory.Exists(sDir)) Directory.CreateDirectory(sDir);
                if (!File.Exists(sPath)) {
                    var sb = new StringBuilder();
                    sb.AppendLine("; 2htm settings, written when Use configuration is checked at OK.");
                    sb.AppendLine("; Delete this file to reset, or press Default settings in the dialog.");
                    sb.AppendLine("[" + sSettingsSection + "]");
                    File.WriteAllText(sPath, sb.ToString(), new UTF8Encoding(true));
                }
                // InixCodec.writeValue changes one value in place, leaving
                // every comment and every other key where it was.
                InixCodec.writeValue(sPath, sSettingsSection, "SourceFiles", sSource ?? "");
                InixCodec.writeValue(sPath, sSettingsSection, "OutputDirectory", sOutputDir ?? "");
                InixCodec.writeValue(sPath, sSettingsSection, "StripImages", bStrip ? "yes" : "no");
                InixCodec.writeValue(sPath, sSettingsSection, "PlainText", bPlain ? "yes" : "no");
                InixCodec.writeValue(sPath, sSettingsSection, "ForceReplacements", bForce ? "yes" : "no");
                InixCodec.writeValue(sPath, sSettingsSection, "ViewOutput", bView ? "yes" : "no");
                InixCodec.writeValue(sPath, sSettingsSection, "LogSession", bLog ? "yes" : "no");
                logger.info("Saved configuration to " + sPath);
                if (File.Exists(getLegacyPath())) {
                    File.Delete(getLegacyPath());
                    logger.info("Removed the old " + getLegacyPath() + "; the settings are now in " + sPath);
                }
            } catch (Exception ex) {
                // Writing the config is a convenience; don't fail
                // the whole run if it can't be written, but do
                // surface the failure to a GUI user — otherwise
                // the "Use configuration" feature silently fails
                // to persist and the user has no way to tell.
                string sMsg = "Could not save configuration to:\r\n" +
                    sPath + "\r\n\r\n" + ex.Message;
                Console.Error.WriteLine(sMsg);
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

        static readonly Dictionary<string, string> dLegacyNames = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase) {
            { "force_replacements", "ForceReplacements" }, { "log_session", "LogSession" },
            { "output_directory", "OutputDirectory" }, { "plain_text", "PlainText" },
            { "source_files", "SourceFiles" }, { "strip_images", "StripImages" },
            { "view_output", "ViewOutput" } };

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
                // The names of versions before 1.19, read as today's.
                string sNew;
                if (dLegacyNames.TryGetValue(sKey, out sNew)) sKey = sNew;
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
