@echo off
REM Double-click launcher for Set-StudentTimes.v<version>.ps1.
REM The .ps1's filename carries its own version (e.g. Set-StudentTimes.v2026-09-17.1.ps1)
REM so `dir` alone shows which one you have - no need to open it. This launcher
REM finds whichever version is actually present next to it, so it never needs
REM editing when the version bumps; if more than one somehow coexists, it picks
REM the newest by name (date-stamped versions sort correctly as plain text).
REM
REM Self-elevates (needed so the GUI's copied "net user" command can be
REM pasted straight into an already-elevated Command Prompt afterwards, and
REM so ADSI can read LoginHours for protected accounts), then runs the .ps1
REM with the execution policy bypassed for just this one run - no permanent
REM change to the machine's PowerShell settings, no setup needed first.
REM Only fails if execution policy is locked by Group Policy, which no
REM per-machine trick can get around.

set "TARGET="
for /f "delims=" %%F in ('dir /b /o-n "%~dp0Set-StudentTimes.v*.ps1" 2^>nul') do (
    if not defined TARGET set "TARGET=%~dp0%%F"
)

if not defined TARGET (
    echo Could not find a Set-StudentTimes.v*.ps1 file in this folder:
    echo   %~dp0
    echo Make sure Run-StudentTimes.bat stays in the same folder as the .ps1 file.
    pause
    exit /b 1
)

net session >nul 2>&1
if %errorlevel% neq 0 (
    echo Requesting administrator privileges...
    powershell -NoProfile -Command "Start-Process -FilePath '%~f0' -Verb RunAs"
    exit /b
)

powershell -NoProfile -ExecutionPolicy Bypass -File "%TARGET%"
if %errorlevel% neq 0 (
    echo.
    echo PowerShell exited with an error before its own window could open.
    pause
)
