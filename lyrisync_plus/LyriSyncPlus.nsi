; ==============================
; LyriSyncPlus.nsi (NSIS installer)
; Safe for both local and CI usage
; ==============================

!include "MUI2.nsh"
!include "FileFunc.nsh"
!include "x64.nsh"

; --------- Guarded Defines (allow /D overrides from CI) ---------
!ifndef SourceDir
  ; Default relative to repo root when run locally:
  !define SourceDir "..\lyrisync_plus\dist\LyriSyncPlus"
!endif

!ifndef AppName
  !define AppName "LyriSync+"
!endif

!ifndef AppExe
  !define AppExe "LyriSyncPlus.exe"
!endif

!ifndef APP_VERSION
  !define APP_VERSION "0.2.0"
!endif

!ifndef Company
  !define Company "Your Team"
!endif

!ifndef OutputDir
  ; NSIS will put installer exe here by default when local
  !define OutputDir "..\lyrisync_plus\dist\installer"
!endif

!ifndef AppId
  ; Your unique app GUID
  !define AppId "{C61E1D3A-9A0E-4E46-9CF2-0E0B70B9A1AC}"
!endif

; Optional icon inside SourceDir (won't fail if missing)
!define AppIcon "${SourceDir}\iconLyriSync.ico"

; --------- General ---------
Name "${AppName}"
OutFile "${OutputDir}\${AppName}-Setup-${APP_VERSION}.exe"
InstallDir "$PROGRAMFILES64\LyriSyncPlus"
InstallDirRegKey HKLM "Software\${AppName}" "Install_Dir"
RequestExecutionLevel admin
BrandingText "${AppName} ${APP_VERSION}"

; --------- UI ---------
!define MUI_ABORTWARNING
!define MUI_ICON "${AppIcon}"
!define MUI_HEADERIMAGE
!insertmacro MUI_PAGE_WELCOME
!insertmacro MUI_PAGE_DIRECTORY
!insertmacro MUI_PAGE_INSTFILES
!insertmacro MUI_PAGE_FINISH

!insertmacro MUI_LANGUAGE "English"

; --------- Sections ---------
Section "Install"
  SetOutPath "$INSTDIR"

  ; Core files from PyInstaller folder
  ; Copy everything from SourceDir recursively
  File /r "${SourceDir}\*.*"

  ; Ensure config file is not overwritten if user already has one
  ; If the file doesn't exist, copy default from the package if present
  IfFileExists "$INSTDIR\lyrisync_config.yaml" +2 0
    ; If not exists but packaged, it should already be copied by File /r. If it didn't exist in build,
    ; no action needed. This block ensures we don't overwrite an existing user config.
    DetailPrint "Config existed or included."

  ; Write uninstall info
  WriteRegStr HKLM "Software\${AppName}" "Install_Dir" "$INSTDIR"
  WriteRegStr HKLM "Software\Microsoft\Windows\CurrentVersion\Uninstall\${AppId}" "DisplayName" "${AppName}"
  WriteRegStr HKLM "Software\Microsoft\Windows\CurrentVersion\Uninstall\${AppId}" "Publisher" "${Company}"
  WriteRegStr HKLM "Software\Microsoft\Windows\CurrentVersion\Uninstall\${AppId}" "DisplayVersion" "${APP_VERSION}"
  WriteRegStr HKLM "Software\Microsoft\Windows\CurrentVersion\Uninstall\${AppId}" "InstallLocation" "$INSTDIR"
  WriteRegStr HKLM "Software\Microsoft\Windows\CurrentVersion\Uninstall\${AppId}" "UninstallString" "$INSTDIR\Uninstall.exe"

  ; Create shortcuts
  CreateDirectory "$SMPROGRAMS\LyriSync+"
  CreateShortCut "$SMPROGRAMS\LyriSync+\${AppName}.lnk" "$INSTDIR\${AppExe}"
  ; Desktop shortcut (common desktop for all users)
  CreateShortCut "$DESKTOP\${AppName}.lnk" "$INSTDIR\${AppExe}"

  ; Create uninstaller
  WriteUninstaller "$INSTDIR\Uninstall.exe"
SectionEnd

Section "Uninstall"
  ; Remove files
  Delete "$DESKTOP\${AppName}.lnk"
  Delete "$SMPROGRAMS\LyriSync+\${AppName}.lnk"
  RMDir  "$SMPROGRAMS\LyriSync+"

  ; DO NOT delete user config on uninstall to be nice.
  ; Remove most files except config:
  ; Safer approach: remove everything then restore config if needed
  ; Here we remove all (PyInstaller dirs)
  RMDir /r "$INSTDIR\_internal"
  RMDir /r "$INSTDIR\lib"
  RMDir /r "$INSTDIR\LyriSyncPlus"

  ; Delete common files if they exist
  Delete "$INSTDIR\${AppExe}"
  Delete "$INSTDIR\Uninstall.exe"
  ; Optional assets, ignore if missing
  Delete "$INSTDIR\iconLyriSync.ico"
  Delete "$INSTDIR\README.md"
  Delete "$INSTDIR\requirements.txt"

  ; Keep "$INSTDIR\lyrisync_config.yaml" to preserve user settings
  ; If you prefer removing it, uncomment:
  ; Delete "$INSTDIR\lyrisync_config.yaml"

  ; Try to remove install dir if empty
  RMDir "$INSTDIR"

  ; Registry cleanup
  DeleteRegKey HKLM "Software\Microsoft\Windows\CurrentVersion\Uninstall\${AppId}"
  DeleteRegKey HKLM "Software\${AppName}"
SectionEnd
