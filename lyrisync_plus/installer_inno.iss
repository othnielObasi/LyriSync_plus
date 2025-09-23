; =========================================
; installer_inno.iss — Inno Setup for LyriSync+
; Safe for both local and CI usage
; =========================================

#define DefaultSourceDir "..\lyrisync_plus\dist\LyriSyncPlus"

#ifndef SourceDir
  #define SourceDir DefaultSourceDir
#endif

#ifndef MyAppName
  #define MyAppName "LyriSync+"
#endif

#ifndef MyAppVersion
  #define MyAppVersion "0.2.0"
#endif

#ifndef MyAppPublisher
  #define MyAppPublisher "Your Team"
#endif

#ifndef MyAppExeName
  #define MyAppExeName "LyriSyncPlus.exe"
#endif

; Optional icon (will be used if present)
#ifndef MyIconFile
  #define MyIconFile SourceDir + "\iconLyriSync.ico"
#endif

[Setup]
AppId={{C61E1D3A-9A0E-4E46-9CF2-0E0B70B9A1AC}
AppName={#MyAppName}
AppVersion={#MyAppVersion}
AppPublisher={#MyAppPublisher}
DefaultDirName={pf}\LyriSyncPlus
DefaultGroupName=LyriSync+
OutputBaseFilename=LyriSyncPlus-Setup-{#MyAppVersion}
Compression=lzma
SolidCompression=yes
ArchitecturesInstallIn64BitMode=x64
WizardStyle=modern
SetupIconFile={#MyIconFile}
DisableReadyMemo=no
DisableDirPage=no

[Files]
; Copy everything from PyInstaller directory
Source: "{#SourceDir}\*"; DestDir: "{app}"; Flags: recursesubdirs createallsubdirs ignoreversion

; Keep user config on updates: only create if missing
Source: "{#SourceDir}\lyrisync_config.yaml"; DestDir: "{app}"; Flags: onlyifdoesntexist ignoreversion

; Optional extras if present
Source: "{#SourceDir}\README.md"; DestDir: "{app}"; Flags: ignoreversion
Source: "{#SourceDir}\requirements.txt"; DestDir: "{app}"; Flags: ignoreversion
Source: "{#SourceDir}\iconLyriSync.ico"; DestDir: "{app}"; Flags: ignoreversion

[Icons]
Name: "{group}\{#MyAppName}"; Filename: "{app}\{#MyAppExeName}"; IconFilename: "{app}\iconLyriSync.ico"
Name: "{commondesktop}\{#MyAppName}"; Filename: "{app}\{#MyAppExeName}"; IconFilename: "{app}\iconLyriSync.ico"; Tasks: desktopicon

[Tasks]
Name: "desktopicon"; Description: "Create a &desktop shortcut"; GroupDescription: "Additional icons:"; Flags: unchecked

[Run]
Filename: "{app}\{#MyAppExeName}"; Description: "Launch {#MyAppName}"; Flags: nowait postinstall skipifsilent

[UninstallDelete]
; Keep config by default (comment this line if you want to delete it on uninstall)
; Type: files; Name: "{app}\lyrisync_config.yaml"
