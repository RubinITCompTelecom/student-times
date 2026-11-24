@echo off
REM Makes THIS PC accept Network Push / Update-StudentTimesGui (no GUI needed).
REM Run from an administrator prompt (Splashtop Command Prompt normally is). Safe to run twice.
setlocal
set "TARGET="
for /f "delims=" %%F in ('dir /b /o-n "%~dp0Enable-PushAcceptance.v*.ps1" 2^>nul') do (
    if not defined TARGET set "TARGET=%~dp0%%F"
)
if not defined TARGET (
    echo Could not find Enable-PushAcceptance.v*.ps1 next to this file.
    exit /b 2
)
powershell.exe -NoProfile -ExecutionPolicy Bypass -File "%TARGET%" %*
exit /b %ERRORLEVEL%
