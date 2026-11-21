@echo off
cd /d "%~dp0"
if exist "%~dp0KTG-Outlook-Mail-Manager-Desktop-0.5.19.exe" (
  start "" "%~dp0KTG-Outlook-Mail-Manager-Desktop-0.5.19.exe"
) else (
  echo Desktop EXE is missing. Run build.ps1 or extract the complete release package.
  pause
  exit /b 1
)
