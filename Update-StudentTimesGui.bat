@echo off
REM Copies this whole student-times-gui folder to other student PCs (replacing what is there).
REM Usage:   Update-StudentTimesGui.bat 192.168.1.21 192.168.1.22
REM          Update-StudentTimesGui.bat            (uses the PCs listed in PushTargets.txt)
REM          Update-StudentTimesGui.bat -DryRun    (shows what would happen, changes nothing)
REM Run from an ADMINISTRATOR Command Prompt / PowerShell. The newest Update-StudentTimesGui.v*.ps1
REM next to this file is picked automatically, so this file never needs editing.
setlocal
set "TARGET="
for /f "delims=" %%F in ('dir /b /o-n "%~dp0Update-StudentTimesGui.v*.ps1" 2^>nul') do (
    if not defined TARGET set "TARGET=%~dp0%%F"
)
if not defined TARGET (
    echo Could not find Update-StudentTimesGui.v*.ps1 next to this file.
    exit /b 2
)
powershell.exe -NoProfile -ExecutionPolicy Bypass -File "%TARGET%" %*
exit /b %ERRORLEVEL%
