; LyriSyncPlus.nsi — NSIS installer
!include "MUI2.nsh"

!define APPNAME   "LyriSyncPlus"
!define COMPANY   "LyriSync"
!define VERSION   "${Version}"
!define SRCDIR    "${SourceDir}"      ; passed in by /DSourceDir
!define OUTFILE   "LyriSyncPlus-${VERSION}-Setup.exe"

; ---------------------------------
; General
; ---------------------------------
Name        "${APPNAME}"
OutFile     "${OUTFILE}"
InstallDir  "$PROGRAMFILES64\${APPNAME}"
RequestExecutionLevel admin
SetCompress auto
SetCompressor /SOLID lzma

; ---------------------------------
; UI
; ---------------------------------
!define MUI_ABORTWARNING
!define MUI_ICON "${SRCDIR}\app.ico"
!define MUI_UNICON "${SRCDIR}\app.ico"

!insertmacro MUI_PAGE_WELCOME
!insertmacro MUI_PAGE_DIRECTORY
!insertmacro MUI_PAGE_INSTFILES
!insertmacro MUI_PAGE_FINISH

!insertmacro MUI_UNPAGE_CONFIRM
!insertmacro MUI_UNPAGE_INSTFILES
!insertmacro MUI_UNPAGE_FINISH

!insertmacro MUI_LANGUAGE "English"

; ---------------------------------
; Sections
; ---------------------------------
Section "Install"
  SetOutPath "$INSTDIR"

  ; Copy all built files from the PyInstaller dist folder
  File /r "${SRCDIR}\*.*"

  ; Start Menu shortcut
  CreateDirectory "$SMPROGRAMS\${APPNAME}"
  CreateShortCut "$SMPROGRAMS\${APPNAME}\${APPNAME}.lnk" "$INSTDIR\LyriSyncPlus.exe"

  ; Desktop shortcut (optional)
  CreateShortCut "$DESKTOP\${APPNAME}.lnk" "$INSTDIR\LyriSyncPlus.exe"

  ; Uninstaller
  WriteUninstaller "$INSTDIR\Uninstall.exe"

  ; Registry entries (Add/Remove Programs)
  WriteRegStr HKLM "Software\Microsoft\Windows\CurrentVersion\Uninstall\${APPNAME}" "DisplayName" "${APPNAME}"
  WriteRegStr HKLM "Software\Microsoft\Windows\CurrentVersion\Uninstall\${APPNAME}" "Publisher"   "${COMPANY}"
  WriteRegStr HKLM "Software\Microsoft\Windows\CurrentVersion\Uninstall\${APPNAME}" "DisplayVersion" "${VERSION}"
  WriteRegStr HKLM "Software\Microsoft\Windows\CurrentVersion\Uninstall\${APPNAME}" "InstallLocation" "$INSTDIR"
  WriteRegStr HKLM "Software\Microsoft\Windows\CurrentVersion\Uninstall\${APPNAME}" "UninstallString" "$INSTDIR\Uninstall.exe"
  WriteRegDWORD HKLM "Software\Microsoft\Windows\CurrentVersion\Uninstall\${APPNAME}" "NoModify" 1
  WriteRegDWORD HKLM "Software\Microsoft\Windows\CurrentVersion\Uninstall\${APPNAME}" "NoRepair" 1
SectionEnd

Section "Uninstall"
  Delete "$SMPROGRAMS\${APPNAME}\${APPNAME}.lnk"
  RMDir  "$SMPROGRAMS\${APPNAME}"
  Delete "$DESKTOP\${APPNAME}.lnk"

  RMDir /r "$INSTDIR"

  DeleteRegKey HKLM "Software\Microsoft\Windows\CurrentVersion\Uninstall\${APPNAME}"
SectionEnd
