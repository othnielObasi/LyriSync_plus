@echo off
setlocal

REM Always run from this script's directory.
cd /d "%~dp0"

REM Build with the maintained spec file for consistent packaging.
pyinstaller --clean --noconfirm lyrisync_plus.spec

if errorlevel 1 (
  echo Build failed.
  exit /b 1
)

echo Build completed successfully.
endlocal
