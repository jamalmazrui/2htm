; =====================================================================
; 2htm installer script for Inno Setup 6.x
;
; Compile with the Inno Setup IDE (ISCC.exe) to produce 2htm_setup.exe.
; The resulting installer:
;   - Targets 64-bit Windows 10 (and later) only.
;   - Requires administrator privileges.
;   - Prompts the user for the installation directory; default is
;     C:\Program Files\2htm.
;   - Shows a brief MIT license summary on the welcome page (no extra
;     wizard screen). The full license text is installed alongside
;     the program as License.htm.
;   - Registers the product for "Apps & Features" uninstall.
;   - Creates a desktop shortcut with hotkey Alt+Ctrl+2 that
;     launches 2htm in GUI mode with saved-configuration loading
;     enabled (equivalent to 2htm -g -u).
;   - Adds "Convert via 2htm" to the File Explorer right-click
;     menu for all file types.
;   - On the final wizard page, offers two PostInstall checkboxes
;     (both checked by default): launch 2htm (with a hotkey reminder)
;     and read the HTML documentation.
;
; This installer ships only the runtime distribution (the .exe, the
; documentation in HTML form, and the license). The Markdown sources,
; the C# source, the build script, and this .iss script live in the
; GitHub repository.
; =====================================================================

#define sAppName       "2htm"
#define sAppVersion    "1.18.3"
#define sAppPublisher  "Jamal Mazrui"
#define sAppUrl        "https://github.com/jamalmazrui/2htm"
#define sAppExeName    "2htm.exe"
#define sAppCopyright  "Copyright (c) 2026 Jamal Mazrui. MIT License."
#define sHotKey        "Alt+Ctrl+2"

[Setup]
AppId={{AD1DA195-AEA7-406B-9B92-AB52D0F3E48A}

AppName={#sAppName}
AppVersion={#sAppVersion}
AppVerName={#sAppName} {#sAppVersion}
AppPublisher={#sAppPublisher}
AppPublisherURL={#sAppUrl}
AppSupportURL={#sAppUrl}
AppUpdatesURL={#sAppUrl}/releases
AppCopyright={#sAppCopyright}
VersionInfoVersion={#sAppVersion}

; Install under Program Files. {autopf} resolves to "Program Files"
; on 64-bit Windows when the installer runs in 64-bit mode (see
; ArchitecturesInstallIn64BitMode below). The user can override this
; default on the wizard's directory page.
DefaultDirName={autopf}\{#sAppName}
DefaultGroupName={#sAppName}
DisableProgramGroupPage=yes
UsePreviousAppDir=yes

; Force the "Select Destination Location" page to always be shown,
; even on reinstall. Without this, DisableDirPage defaults to "auto",
; which means: hide the directory page if a prior install of the same
; AppId is detected. We want the page shown every time so the user
; can review the install location, and so it is obviously editable.
; UsePreviousAppDir=yes pre-fills the field with the previous
; directory, so the user just presses Next on a reinstall to keep the
; same path -- but they can also change it.
DisableDirPage=no
UsePreviousGroup=yes

OutputDir=.
OutputBaseFilename={#sAppName}_setup
Compression=lzma2
SolidCompression=yes
SetupIconFile={#sAppName}.ico
WizardStyle=modern

; Installer requires admin to write to Program Files and HKLM.
PrivilegesRequired=admin
PrivilegesRequiredOverridesAllowed=

; 64-bit Windows only.
ArchitecturesAllowed=x64compatible
ArchitecturesInstallIn64BitMode=x64compatible

Uninstallable=yes
UninstallDisplayIcon={app}\{#sAppExeName}
UninstallDisplayName={#sAppName} {#sAppVersion}

MinVersion=10.0

[Languages]
Name: "english"; MessagesFile: "compiler:Default.isl"

[Messages]
; Replace the default welcome-page body text with one that includes a
; brief MIT license notice. This satisfies the requirement that the
; license summary appear on an existing wizard screen rather than on
; an additional dedicated page (which is what LicenseFile= would
; produce). The full license text is installed alongside the program.
WelcomeLabel2=This will install [name/ver] on your computer.%n%n[name] is licensed under the MIT License: free to use, copy, modify, and distribute; provided "as is" with no warranty. The full license text will be installed as License.htm in the program folder.%n%nIt is recommended that you close all other applications before continuing.

[Files]
; The runtime distribution: the executable, the HTML docs, and the
; license. The icon is embedded in 2htm.exe at build time (csc
; /win32icon flag), so the .ico does not need to ship in the install
; directory.
Source: "{#sAppName}.exe";    DestDir: "{app}"; Flags: ignoreversion
Source: "ReadMe.htm";         DestDir: "{app}"; Flags: ignoreversion
Source: "Announce.htm";       DestDir: "{app}"; Flags: ignoreversion
Source: "License.htm";        DestDir: "{app}"; Flags: ignoreversion

[Icons]
; Start Menu group.
Name: "{group}\{#sAppName}"; \
  Filename: "{app}\{#sAppExeName}"; \
  WorkingDir: "{app}"; \
  Comment: "Convert Office, PDF, and Markdown files to accessible HTML"

Name: "{group}\{#sAppName} ReadMe"; \
  Filename: "{app}\ReadMe.htm"; \
  WorkingDir: "{app}"; \
  Comment: "Documentation for {#sAppName}"

Name: "{group}\Uninstall {#sAppName}"; \
  Filename: "{uninstallexe}"; \
  Comment: "Remove {#sAppName} from this computer"

; Desktop shortcut with the Alt+Ctrl+2 hotkey. Launches 2htm in
; GUI mode (-g) with saved-configuration loading (-u). The
; hotkey is free on Windows but is intercepted by Microsoft
; Word (where it maps to Heading 2 style) when Word has focus;
; it works from Explorer, the desktop, and all non-Word apps.
Name: "{userdesktop}\{#sAppName}"; \
  Filename: "{app}\{#sAppExeName}"; \
  WorkingDir: "{app}"; \
  Parameters: "-g -u"; \
  HotKey: {#sHotKey}; \
  Comment: "Convert files to accessible HTML ({#sHotKey})"

[Registry]
; File Explorer "Convert via 2htm" right-click menu entry.
; Registered unconditionally under HKLM so it is available for
; every user on the machine. The verb passes "%1" (the full
; absolute path of the right-clicked file) to 2htm.exe. The
; process's current working directory is set by Windows to the
; folder containing that file, so the default "-o is CWD"
; behavior lands the output next to the source without any
; shell-placeholder tricks. -f forces overwriting an existing
; output so repeated right-clicks refresh rather than prompt.
;
; The uninsdeletekey flag on the parent verb key causes Inno
; Setup to remove the entire subtree (including the command
; subkey) automatically on uninstall.
Root: HKLM; Subkey: "SOFTWARE\Classes\*\shell\2htm"; \
  ValueType: string; ValueName: ""; ValueData: "Convert via &2htm"; \
  Flags: uninsdeletekey
Root: HKLM; Subkey: "SOFTWARE\Classes\*\shell\2htm"; \
  ValueType: string; ValueName: "Icon"; ValueData: """{app}\{#sAppExeName}"",0"
Root: HKLM; Subkey: "SOFTWARE\Classes\*\shell\2htm\command"; \
  ValueType: string; ValueName: ""; \
  ValueData: """{app}\{#sAppExeName}"" -f ""%1"""

[Run]
; Post-install checkboxes shown on the final wizard page. Both
; default to checked; the user can uncheck either to skip. The launch
; checkbox label includes a reminder of the desktop hotkey so the
; user notices and remembers it.

FileName: "{app}\{#sAppExeName}"; \
  Parameters: "-g"; \
  WorkingDir: "{app}"; \
  Description: "Launch {#sAppName} now (desktop hotkey: {#sHotKey})"; \
  Flags: nowait postinstall skipifsilent

FileName: "{app}\ReadMe.htm"; \
  Description: "Read documentation for {#sAppName}"; \
  Flags: postinstall shellexec skipifsilent
