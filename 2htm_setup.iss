; =====================================================================
; 2htm installer script for Inno Setup 6.x
;
; Compile with the Inno Setup IDE (ISCC.exe) to produce 2htm_setup.exe.
; The resulting installer:
;   - Requires administrator privileges.
;   - Installs 2htm.exe and all supporting documentation files to
;     C:\Program Files\2htm (standard GUI-program install path).
;   - Registers the product for "Apps & Features" uninstall.
;   - Creates a desktop shortcut with hotkey Alt+Ctrl+2 that
;     launches 2htm in GUI mode with saved-configuration loading
;     enabled (equivalent to 2htm -g -u).
;   - Adds "Convert with 2htm" to the File Explorer right-click
;     menu for all file types.
;   - Generates readMe.htm from readMe.md during install by
;     running the just-installed 2htm.exe on its own Markdown
;     documentation.
;   - On the final wizard page, offers two PostInstall checkboxes
;     (both checked by default): launch 2htm, and read the HTML
;     documentation.
; =====================================================================

#define cAppName       "2htm"
#define cAppVersion    "1.18.3"
#define cAppPublisher  "Jamal Mazrui"
#define cAppUrl        "https://github.com/jamalmazrui/2htm"
#define cAppExeName    "2htm.exe"
#define cAppCopyright  "Copyright (c) 2026 Jamal Mazrui. MIT License."

[Setup]
AppId={{AD1DA195-AEA7-406B-9B92-AB52D0F3E48A}

AppName={#cAppName}
AppVersion={#cAppVersion}
AppVerName={#cAppName} {#cAppVersion}
AppPublisher={#cAppPublisher}
AppPublisherURL={#cAppUrl}
AppSupportURL={#cAppUrl}
AppUpdatesURL={#cAppUrl}/releases
AppCopyright={#cAppCopyright}
VersionInfoVersion={#cAppVersion}

; Install under Program Files (standard GUI-program location).
DefaultDirName={pf}\{#cAppName}
DefaultGroupName={#cAppName}
DisableProgramGroupPage=yes
UsePreviousAppDir=yes
UsePreviousGroup=yes

OutputDir=.
OutputBaseFilename=2htm_setup
Compression=lzma2
SolidCompression=yes
SetupIconFile=

; Installer requires admin to write to Program Files and HKLM.
PrivilegesRequired=admin
PrivilegesRequiredOverridesAllowed=

; 64-bit Windows only.
ArchitecturesAllowed=x64compatible
ArchitecturesInstallIn64BitMode=x64compatible

Uninstallable=yes
UninstallDisplayIcon={app}\{#cAppExeName}
UninstallDisplayName={#cAppName} {#cAppVersion}

MinVersion=10.0

[Languages]
Name: "english"; MessagesFile: "compiler:Default.isl"

[Files]
Source: "2htm.exe";         DestDir: "{app}"; Flags: ignoreversion
Source: "readMe.md";        DestDir: "{app}"; Flags: ignoreversion
Source: "license.htm";      DestDir: "{app}"; Flags: ignoreversion
Source: "announce.md";      DestDir: "{app}"; Flags: ignoreversion
Source: "Camel_Type_C#.md"; DestDir: "{app}"; Flags: ignoreversion
Source: "2htm.cs";          DestDir: "{app}"; Flags: ignoreversion
Source: "build2htm.cmd";    DestDir: "{app}"; Flags: ignoreversion

[Icons]
; Start Menu group.
Name: "{group}\{#cAppName}"; \
  Filename: "{app}\{#cAppExeName}"; \
  WorkingDir: "{app}"; \
  Comment: "Convert Office, PDF, and Markdown files to accessible HTML"

Name: "{group}\{#cAppName} readMe"; \
  Filename: "{app}\readMe.htm"; \
  WorkingDir: "{app}"; \
  Comment: "Documentation for {#cAppName}"

Name: "{group}\Uninstall {#cAppName}"; \
  Filename: "{uninstallexe}"; \
  Comment: "Remove {#cAppName} from this computer"

; Desktop shortcut with the Alt+Ctrl+2 hotkey. Launches 2htm in
; GUI mode (-g) with saved-configuration loading (-u). The
; hotkey is free on Windows but is intercepted by Microsoft
; Word (where it maps to Heading 2 style) when Word has focus;
; it works from Explorer, the desktop, and all non-Word apps.
Name: "{userdesktop}\{#cAppName}"; \
  Filename: "{app}\{#cAppExeName}"; \
  WorkingDir: "{app}"; \
  Parameters: "-g -u"; \
  HotKey: Alt+Ctrl+2; \
  Comment: "Convert files to accessible HTML (Alt+Ctrl+2)"

[Registry]
; File Explorer "Convert with 2htm" right-click menu entry.
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
  ValueType: string; ValueName: ""; ValueData: "Convert with &2htm"; \
  Flags: uninsdeletekey
Root: HKLM; Subkey: "SOFTWARE\Classes\*\shell\2htm"; \
  ValueType: string; ValueName: "Icon"; ValueData: """{app}\{#cAppExeName}"",0"
Root: HKLM; Subkey: "SOFTWARE\Classes\*\shell\2htm\command"; \
  ValueType: string; ValueName: ""; \
  ValueData: """{app}\{#cAppExeName}"" -f ""%1"""

[Run]
; Install-phase step (runs before the final wizard page):
; generate readMe.htm from readMe.md using the just-installed
; 2htm.exe. The -f flag overwrites any stale readMe.htm from a
; previous install. runhidden keeps any console window from
; flashing; waituntilterminated blocks until done so postinstall
; can launch readMe.htm without a race.
FileName: "{app}\{#cAppExeName}"; \
  Parameters: "-f readMe.md"; \
  WorkingDir: "{app}"; \
  StatusMsg: "Generating HTML documentation..."; \
  Flags: runhidden waituntilterminated

; Post-install checkboxes shown on the final wizard page. Both
; default to checked; the user can uncheck either to skip.

; Launch 2htm (GUI mode).
FileName: "{app}\{#cAppExeName}"; \
  Parameters: "-g"; \
  WorkingDir: "{app}"; \
  Description: "Launch {#cAppName} now"; \
  Flags: nowait postinstall skipifsilent

; Open the HTML documentation.
FileName: "{app}\readMe.htm"; \
  Description: "Read documentation for {#cAppName}"; \
  Flags: postinstall shellexec skipifsilent

[UninstallDelete]
; readMe.htm is a derived artifact (generated at install time
; from readMe.md), not a shipped file, so it does not get
; removed by Inno Setup's normal "delete what I installed"
; logic. Remove it explicitly so uninstall leaves no trace.
Type: files; Name: "{app}\readMe.htm"
