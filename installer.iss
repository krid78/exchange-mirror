[Setup]
AppName=Exchange Mirror
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
PrivilegesRequired=admin

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
Name: "startup"; Description: "Start Exchange Mirror on Windows &login"; GroupDescription: "Windows startup:"

[Registry]
; Auto-start on login (only if user checked the startup task)
Root: HKCU; Subkey: "Software\Microsoft\Windows\CurrentVersion\Run"; ValueType: string; ValueName: "ExchangeMirror"; ValueData: """{app}\mirror.exe"""; Flags: uninsdeletevalue; Tasks: startup

[Run]
Filename: "{app}\mirror.exe"; Description: "Launch Exchange Mirror now"; Flags: nowait postinstall skipifsilent
