<#
.SYNOPSIS
  Copies this whole student-times-gui folder to other student PCs over the network, replacing
  what is there, so every PC has the newest version without anyone logging on.

.DESCRIPTION
  Uses PowerShell Remoting (WinRM) - the same channel as the tool's Network Push - so each target
  needs "Enable This PC To Receive Pushes" (Step 1) done once. Run from an ELEVATED prompt (it may
  need to add the target to this PC's TrustedHosts).

  For every target it: creates the destination folder, deletes older Set-StudentTimes.v*.ps1 / Enable-PushAcceptance.v*.ps1 /
  Update-StudentTimesGui.v*.ps1 copies (so the launcher can never pick a stale one), copies every
  file in this folder over the top (-Force), then checks the newest version file really arrived.
  Per-machine files (StudentTimesLog.csv, StudentTimesProfiles.json, StudentTimesDebug.log,
  PushTargets.txt) and .git are NEVER copied or deleted - each PC keeps its own.

.EXAMPLE
  .\Update-StudentTimesGui.bat 192.168.1.21 192.168.1.22
.EXAMPLE
  powershell -ExecutionPolicy Bypass -File Update-StudentTimesGui.v2026-10-25.1.ps1 -Targets 192.168.1.21,192.168.1.22
.EXAMPLE
  .\Update-StudentTimesGui.bat -DryRun          # shows what would be copied, touches nothing
#>
[CmdletBinding()]
param(
    [Parameter(Position = 0, ValueFromRemainingArguments = $true)]
    [string[]]$Targets,                                   # names or IPs; blank = every entry in PushTargets.txt
    [string]$Destination = 'C:\Users\Admin\Desktop\student-times-gui',
    [string]$Source = '',
    [System.Management.Automation.PSCredential]$Credential,
    [switch]$DryRun
)

$ErrorActionPreference = 'Stop'
if (-not $Source) { $Source = Split-Path -Parent $MyInvocation.MyCommand.Path }   # ($PSScriptRoot is empty inside param defaults on 5.1)
$ScriptVersion = '2026-10-29.2'
$KeepPerMachine = @('StudentTimesLog.csv', 'StudentTimesProfiles.json', 'StudentTimesDebug.log', 'PushTargets.txt')
$VersionedPatterns = @('Set-StudentTimes.v*.ps1', 'Update-StudentTimesGui.v*.ps1', 'Enable-PushAcceptance.v*.ps1')

function Out-Status([string]$tag, [string]$msg) {
    $color = switch ($tag) { 'OK' { 'Green' } 'FAIL' { 'Red' } 'WARN' { 'Yellow' } default { 'Gray' } }
    Write-Host "[$tag] $msg" -ForegroundColor $color
}

# --- targets: command line, else PushTargets.txt (notes after # are ignored) ---
$Targets = @($Targets | ForEach-Object { $_ -split ',' } | ForEach-Object { ($_ -replace '\s*#.*$', '').Trim() } | Where-Object { $_ })
if ($Targets.Count -eq 0) {
    $listFile = Join-Path $Source 'PushTargets.txt'
    if (Test-Path $listFile) {
        $Targets = @(Get-Content $listFile | ForEach-Object { ($_ -replace '\s*#.*$', '').Trim() } | Where-Object { $_ })
    }
}
if ($Targets.Count -eq 0) {
    Write-Host 'No target PCs. Give names/IPs:  Update-StudentTimesGui.bat 192.168.1.21 192.168.1.22   (or list them in PushTargets.txt)' -ForegroundColor Red
    exit 2
}

# --- what to copy ---
$Source = (Resolve-Path $Source).Path.TrimEnd('\')
$files = @(Get-ChildItem -Path $Source -Recurse -File -Force |
    Where-Object { $KeepPerMachine -notcontains $_.Name -and $_.FullName -notmatch '\\\.git(\\|$)' })
$newestVersion = Get-ChildItem -Path $Source -Filter 'Set-StudentTimes.v*.ps1' -File | Sort-Object Name -Descending | Select-Object -First 1
Write-Host "Update-StudentTimesGui v$ScriptVersion" -ForegroundColor Cyan
Write-Host "Source     : $Source ($($files.Count) files; newest tool = $(if ($newestVersion) { $newestVersion.Name } else { 'NONE FOUND' }))"
Write-Host "Destination: $Destination"
Write-Host "Targets    : $($Targets -join ', ')"
if (-not $newestVersion) { Out-Status 'FAIL' 'No Set-StudentTimes.v*.ps1 in the source folder - nothing to deploy.'; exit 2 }

if ($DryRun) {
    Write-Host "`nDRY RUN - nothing is contacted or changed. Would copy:" -ForegroundColor Yellow
    $files | ForEach-Object { '  ' + $_.FullName.Substring($Source.Length + 1) }
    Write-Host "Would first delete older versions on each target matching: $($VersionedPatterns -join ', ') (except the ones being copied)."
    Write-Host "Never touched on targets: $($KeepPerMachine -join ', ')"
    exit 0
}

# --- credential: given, else the one the GUI remembered on this PC, else ask ---
if (-not $Credential) {
    $saved = Join-Path $env:LOCALAPPDATA 'StudentTimes\PushCredential.xml'
    if (Test-Path $saved) {
        try { $Credential = Import-Clixml $saved; Write-Host "Using the saved credential: $($Credential.UserName)" } catch { }
    }
}
if (-not $Credential) { $Credential = Get-Credential -Message 'Admin account that exists on the student PCs (e.g. .\Admin)' }
if (-not $Credential) { Out-Status 'FAIL' 'No credential.'; exit 2 }

function Ensure-TrustedHost([string]$name) {
    $current = (Get-Item WSMan:\localhost\Client\TrustedHosts -ErrorAction SilentlyContinue).Value
    $list = @()
    if ($current) { $list = @($current -split ',' | ForEach-Object { $_.Trim() } | Where-Object { $_ }) }
    if ($list -icontains $name -or $list -contains '*') { return }
    $list += $name
    Set-Item WSMan:\localhost\Client\TrustedHosts -Value ($list -join ',') -Force
}

$failed = 0
foreach ($target in $Targets) {
    Write-Host "`n=== $target ===" -ForegroundColor Cyan
    $session = $null
    try {
        try { Ensure-TrustedHost $target } catch { throw "Could not add $target to TrustedHosts - run this from an ELEVATED (Run as administrator) prompt. $($_.Exception.Message)" }
        $session = New-PSSession -ComputerName $target -Credential $Credential -ErrorAction Stop
        $who = Invoke-Command -Session $session -ScriptBlock { "$env:COMPUTERNAME" }
        Out-Status 'OK' "Connected to $who"

        # destination folder + clear out older versioned scripts so only the new ones remain
        $removed = Invoke-Command -Session $session -ScriptBlock {
            param($dest, $patterns, $keepNames)
            if (-not (Test-Path $dest)) { New-Item -ItemType Directory -Path $dest -Force | Out-Null }
            $gone = @()
            foreach ($pat in $patterns) {
                foreach ($f in @(Get-ChildItem -Path $dest -Filter $pat -File -ErrorAction SilentlyContinue)) {
                    if ($keepNames -notcontains $f.Name) { Remove-Item $f.FullName -Force; $gone += $f.Name }
                }
            }
            return $gone
        } -ArgumentList $Destination, $VersionedPatterns, @($files | ForEach-Object { $_.Name })
        foreach ($r in @($removed)) { Out-Status 'OK' "Removed old version: $r" }

        $copied = 0
        foreach ($f in $files) {
            $rel = $f.FullName.Substring($Source.Length + 1)
            $destFile = Join-Path $Destination $rel
            $destDir = Split-Path $destFile
            Invoke-Command -Session $session -ScriptBlock { param($d) if (-not (Test-Path $d)) { New-Item -ItemType Directory -Path $d -Force | Out-Null } } -ArgumentList $destDir
            Copy-Item -Path $f.FullName -Destination $destFile -ToSession $session -Force
            $copied++
        }
        Out-Status 'OK' "Copied $copied file(s) to $Destination (existing files replaced)"

        # prove it: the newest tool file must be there and byte-identical
        $srcHash = (Get-FileHash $newestVersion.FullName -Algorithm SHA256).Hash
        $check = Invoke-Command -Session $session -ScriptBlock {
            param($dest, $name)
            $p = Join-Path $dest $name
            [PSCustomObject]@{ Exists = (Test-Path $p); Hash = $(if (Test-Path $p) { (Get-FileHash $p -Algorithm SHA256).Hash } else { '' }); Versions = @(Get-ChildItem -Path $dest -Filter 'Set-StudentTimes.v*.ps1' -File | ForEach-Object { $_.Name }) }
        } -ArgumentList $Destination, $newestVersion.Name
        if ($check.Exists -and $check.Hash -eq $srcHash) { Out-Status 'OK' "Verified: $($newestVersion.Name) is on $who and matches this PC's copy. Versions there now: $($check.Versions -join ', ')" }
        else { throw "Verification failed - $($newestVersion.Name) is missing or different on $who." }
    }
    catch {
        $failed++
        Out-Status 'FAIL' "$target : $($_.Exception.Message -replace '\s+', ' ')"
    }
    finally {
        if ($session) { Remove-PSSession $session -ErrorAction SilentlyContinue }
    }
}

Write-Host ''
if ($failed -gt 0) { Out-Status 'FAIL' "$failed of $($Targets.Count) PC(s) FAILED."; exit 1 }
Out-Status 'OK' "All $($Targets.Count) PC(s) updated and verified."
exit 0
