[Setup]
AppName=Exchange Mirror
; Keep the historical identity (previously defaulted to AppName) for updates.
AppId=Exchange Mirror
AppVersion=1.0
AppPublisher=Exchange Mirror
DefaultDirName={autopf}\Exchange Mirror
DefaultGroupName=Exchange Mirror
UninstallDisplayIcon={app}\mirror.exe
OutputDir=dist
OutputBaseFilename=Exchange_Mirror_Setup
Compression=lzma2
SolidCompression=yes
ArchitecturesAllowed=x64compatible
ArchitecturesInstallIn64BitMode=x64compatible
PrivilegesRequired=lowest
PrivilegesRequiredOverridesAllowed=dialog
UsePreviousPrivileges=yes

[Files]
; Main executable and runtime
Source: "dist\mirror\mirror.exe"; DestDir: "{app}"; Flags: ignoreversion
Source: "dist\mirror\_internal\*"; DestDir: "{app}\_internal"; Flags: ignoreversion recursesubdirs createallsubdirs

; Config file - never overwrite user edits on upgrade
Source: "dist\mirror\config.py"; DestDir: "{app}"; Flags: onlyifdoesntexist

[Icons]
Name: "{group}\Exchange Mirror"; Filename: "{app}\mirror.exe"
Name: "{group}\Edit Configuration"; Filename: "notepad.exe"; Parameters: """{app}\config.py"""
Name: "{group}\Uninstall Exchange Mirror"; Filename: "{uninstallexe}"
Name: "{autodesktop}\Exchange Mirror"; Filename: "{app}\mirror.exe"; Tasks: desktopicon

[Tasks]
Name: "desktopicon"; Description: "Create a &desktop shortcut"; GroupDescription: "Additional shortcuts:"
; Global installs leave per-user startup preferences to the application.
Name: "startup"; Description: "Start Exchange Mirror on Windows &login"; GroupDescription: "Windows startup:"; Check: not IsAdminInstallMode

[Registry]
; Never write to the elevated installer's user profile in global install mode.
Root: HKCU; Subkey: "Software\Microsoft\Windows\CurrentVersion\Run"; ValueType: string; ValueName: "ExchangeMirror"; ValueData: """{app}\mirror.exe"""; Flags: uninsdeletevalue; Tasks: startup; Check: not IsAdminInstallMode

[Run]
Filename: "{app}\mirror.exe"; Description: "Launch Exchange Mirror now"; Flags: nowait postinstall skipifsilent
