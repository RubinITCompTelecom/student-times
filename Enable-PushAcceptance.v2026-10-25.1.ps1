<#
.SYNOPSIS
  Makes THIS PC accept Network Push / Update-StudentTimesGui from the technician's PC.
  Same as the GUI's "Enable This PC To Receive Pushes" button, but runnable with no screen
  (Splashtop Command Prompt / PowerShell, or any admin prompt). Safe to run again.

  1. Turns on the WinRM listener + firewall rule (Enable-PSRemoting, works on a Public network too).
  2. Lets a local admin account log in remotely with a FULL token (LocalAccountTokenFilterPolicy=1);
     without this the push is silently downgraded and `net user` fails with "access denied".
  3. Checks that the PC now answers on port 5985 and prints its name and IP addresses.
#>
$ErrorActionPreference = 'Stop'
$fail = 0
function Say([string]$tag, [string]$msg) { Write-Host "[$tag] $msg" }

$isAdmin = ([Security.Principal.WindowsPrincipal][Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)
if (-not $isAdmin) { Say 'FAIL' 'Not running as administrator - nothing was changed. Use an elevated prompt (Splashtop Command Prompt normally is).'; exit 1 }

try { Enable-PSRemoting -Force -SkipNetworkProfileCheck | Out-Null; Say 'OK' 'WinRM listener + firewall rule enabled.' }
catch { Say 'FAIL' "Enable-PSRemoting: $($_.Exception.Message)"; $fail++ }

try {
    Set-ItemProperty -Path 'HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Policies\System' -Name LocalAccountTokenFilterPolicy -Value 1 -Type DWord -Force
    Say 'OK' 'Local admin accounts can authenticate remotely with a full token.'
} catch { Say 'FAIL' "LocalAccountTokenFilterPolicy: $($_.Exception.Message)"; $fail++ }

try { $svc = Get-Service WinRM; if ($svc.Status -eq 'Running') { Say 'OK' 'WinRM service is running.' } else { Say 'FAIL' "WinRM service is $($svc.Status)."; $fail++ } }
catch { Say 'FAIL' "WinRM service: $($_.Exception.Message)"; $fail++ }

try {
    $null = Test-WSMan -ComputerName localhost -ErrorAction Stop
    $ips = @(Get-NetIPAddress -AddressFamily IPv4 -ErrorAction SilentlyContinue | Where-Object { $_.IPAddress -notmatch '^(127\.|169\.254\.)' } | ForEach-Object { $_.IPAddress })
    Say 'OK' "READY to receive pushes: $env:COMPUTERNAME  ($($ips -join ', '))"
} catch { Say 'FAIL' "This PC does not answer WinRM yet: $($_.Exception.Message -replace '\s+', ' ')"; $fail++ }

if ($fail -gt 0) { Say 'FAIL' "$fail step(s) failed."; exit 1 }
exit 0
