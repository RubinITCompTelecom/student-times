# Set-StudentTimes.ps1
# Simple GUI to build/inspect Windows "net user /times:" logon-hour schedules,
# replacing the manual command-prompt process described in
# "Big Customers Info\Student Times Checklist.xlsx".
#
# Run via Run-StudentTimes.bat (self-elevating double-click launcher) on the
# target machine itself (the student/customer PC), while remoted in via
# Splashtop/RDP. Running the .ps1 directly needs an elevated PowerShell
# window:
#
#   powershell -ExecutionPolicy Bypass -File Set-StudentTimes.ps1
#
# This tool never runs "net user" for you. It only builds the exact command
# text and copies it to the clipboard -- you paste it into an elevated
# Command Prompt yourself. It can also READ the account's current schedule
# (via ADSI LoginHours) so you don't have to guess what's already set.
#
# It keeps a local log (StudentTimesLog.csv, next to this script) of every
# schedule you generate on this machine, and gives you a ready
# tab-separated line to paste straight into the "Student Times Checklist"
# spreadsheet.
#
# It also keeps per-account SAVED SCHEDULES (StudentTimesProfiles.json, next
# to this script) for one-click temporary overrides:
#   - "Unlock All For Today" auto-saves the account's real current schedule
#     before switching everything open, so tomorrow's "put it back" is a
#     single click.
#   - "Restore Previous" re-applies whatever was auto-saved right before the
#     last override.
#   - 3 general-purpose Quick Slots let you save a couple of named schedules
#     (e.g. "School Hours" vs "Summer Hours") per account and reload either
#     in one click.

Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing

# ---------------------------------------------------------------------------
# Constants / state
# ---------------------------------------------------------------------------

# Bump this on every change that ships. Shown in the window title and logged
# on startup so it's obvious at a glance whether a given PC is running an old
# copy - useful once this script is on more than one machine.
$ScriptVersion = '2026-11-23.1'

$DayNames     = @('Su','M','T','W','Th','F','Sa')
$ScriptDir    = Split-Path -Parent $MyInvocation.MyCommand.Path
$LogPath      = Join-Path $ScriptDir 'StudentTimesLog.csv'
$ProfilesPath = Join-Path $ScriptDir 'StudentTimesProfiles.json'
$DebugLogPath = Join-Path $ScriptDir 'StudentTimesDebug.log'

function Write-DebugLog([string]$message) {
    try {
        $ts = Get-Date -Format 'yyyy-MM-dd HH:mm:ss'
        Add-Content -Path $DebugLogPath -Value "$ts  $message"
    } catch { }
}

# $Grid[$day,$hour] = $true (allowed) / $false (blocked), day 0=Sunday, hour 0..23 local time
$Grid = New-Object 'bool[,]' 7,24

function Reset-Grid([bool]$value) {
    for ($d = 0; $d -lt 7; $d++) {
        for ($h = 0; $h -lt 24; $h++) { $Grid[$d,$h] = $value }
    }
}
Reset-Grid $false

# ---------------------------------------------------------------------------
# ADSI read/write of LoginHours (168-bit week, stored as UTC hour-of-week)
# ---------------------------------------------------------------------------

function Get-LocalUtcOffsetHours {
    # Windows stores LoginHours relative to STANDARD time (no daylight-saving
    # adjustment), even while DST is in effect -- a well-known quirk. Using
    # the DST-aware current offset here shifts every hour by 1 whenever DST
    # is active, which is exactly what caused schedules to read back wrong.
    $tz = [System.TimeZoneInfo]::Local
    return $tz.BaseUtcOffset.TotalHours
}

function Get-CurrentLoginHours([string]$userName) {
    # Returns the 7x24 bool grid (local time) currently set on this machine
    # for $userName, or $null if the account can't be read.
    Write-DebugLog "Get-CurrentLoginHours: PC=$env:COMPUTERNAME  querying WinNT://./$userName,user"
    try {
        $adsiUser = [ADSI]"WinNT://./$userName,user"
        # Force a real bind now (ADSI is lazy) so a bad account name throws
        # here, not silently later.
        $null = $adsiUser.Name
        $bytes = $adsiUser.Properties['LoginHours'].Value
    } catch {
        Write-DebugLog "Get-CurrentLoginHours: ADSI ERROR for '$userName': $($_.Exception.Message)"
        return $null
    }

    if (-not $bytes -or $bytes.Count -eq 0) {
        Write-DebugLog "Get-CurrentLoginHours: LoginHours property is empty/null for '$userName' -> treating as UNRESTRICTED (all hours allowed). If this account actually has restricted hours, this is the bug."
        $g = New-Object 'bool[,]' 7,24
        for ($d = 0; $d -lt 7; $d++) { for ($h = 0; $h -lt 24; $h++) { $g[$d,$h] = $true } }
        return ,$g
    }

    Write-DebugLog "Get-CurrentLoginHours: raw LoginHours bytes ($($bytes.Count)) = $($bytes -join ',')"

    $offset = Get-LocalUtcOffsetHours
    Write-DebugLog "Get-CurrentLoginHours: local timezone = $([System.TimeZoneInfo]::Local.Id), using standard-time offset = $offset hours"

    $g = New-Object 'bool[,]' 7,24
    for ($utcIndex = 0; $utcIndex -lt 168; $utcIndex++) {
        $byteIndex = [math]::Floor($utcIndex / 8)
        $bitIndex  = $utcIndex % 8
        $bitSet = ($bytes[$byteIndex] -band (1 -shl $bitIndex)) -ne 0
        if ($bitSet) {
            $localIndex = (($utcIndex + $offset) % 168 + 168) % 168
            $localIndex = [int][math]::Round($localIndex)
            $day  = [math]::Floor($localIndex / 24)
            $hour = $localIndex % 24
            $g[$day, $hour] = $true
        }
    }

    $decodedSummary = @()
    for ($d = 0; $d -lt 7; $d++) {
        $row = @(); for ($h = 0; $h -lt 24; $h++) { $row += $g[$d, $h] }
        $ranges = Get-DayRanges $row
        if ($ranges.Count -gt 0) { $decodedSummary += "$($DayNames[$d]):$($ranges -join '/')" }
    }
    Write-DebugLog "Get-CurrentLoginHours: decoded -> $($decodedSummary -join '; ')"

    # NOTE: the leading comma is required. Without it, PowerShell unravels a
    # 2D array on return into a flat System.Object[] -- $existing[$d,$h] then
    # silently returns a 2-element slice instead of one cell, which is what
    # was making every loaded schedule come out as "all hours allowed".
    return ,$g
}

# ---------------------------------------------------------------------------
# Command-string generation from the grid, and the reverse (parse back)
# ---------------------------------------------------------------------------

function Get-DayRanges([bool[]]$dayHours) {
    # Collapses a 24-element bool array into a list of "HH:MM-HH:MM" ranges.
    $ranges = @()
    $start = -1
    for ($h = 0; $h -lt 24; $h++) {
        if ($dayHours[$h] -and $start -lt 0) { $start = $h }
        if (-not $dayHours[$h] -and $start -ge 0) {
            $ranges += ('{0:D2}:00-{1:D2}:00' -f $start, $h)
            $start = -1
        }
    }
    if ($start -ge 0) { $ranges += ('{0:D2}:00-24:00' -f $start) }
    return $ranges
}

function Build-TimesString {
    # Returns the value that goes after "/times:" -- empty string means
    # "deny all hours" (still a valid, intentional /times: with nothing after it).
    $allAllowed = $true
    $allBlocked = $true
    for ($d = 0; $d -lt 7; $d++) {
        for ($h = 0; $h -lt 24; $h++) {
            if ($Grid[$d,$h]) { $allBlocked = $false } else { $allAllowed = $false }
        }
    }
    if ($allAllowed) { return 'all' }
    if ($allBlocked) { return '' }

    # Per-day range signature, so identical consecutive days can be merged
    # (e.g. "M-F,09:00-17:00" instead of five separate entries).
    $daySignatures = @()
    for ($d = 0; $d -lt 7; $d++) {
        $hours = @()
        for ($h = 0; $h -lt 24; $h++) { $hours += $Grid[$d,$h] }
        $ranges = Get-DayRanges $hours
        $daySignatures += ,($ranges -join ',')
    }

    $entries = @()
    $d = 0
    while ($d -lt 7) {
        if ([string]::IsNullOrEmpty($daySignatures[$d])) { $d++; continue }
        $runEnd = $d
        while (($runEnd + 1) -lt 7 -and $daySignatures[$runEnd + 1] -eq $daySignatures[$d]) { $runEnd++ }

        $dayLabel = if ($runEnd -eq $d) { $DayNames[$d] } else { "$($DayNames[$d])-$($DayNames[$runEnd])" }
        foreach ($range in ($daySignatures[$d] -split ',')) {
            $entries += "$dayLabel,$range"
        }
        $d = $runEnd + 1
    }
    return ($entries -join ';')
}

function Normalize-DayToken([string]$token) {
    # Accepts our own casing (Su/M/T/W/Th/F/Sa) plus common variants seen in
    # the checklist sheet and setup docs: bare "S" for Sunday, any casing
    # ("TH", "th", "sa"). Returns the index into $DayNames, or -1.
    $t = $token.Trim()
    if ($t.Length -eq 0) { return -1 }
    if ($t.ToUpper() -eq 'S') { return 0 }  # bare "S" = Sunday, seen in older docs
    $normalized = $t.Substring(0,1).ToUpper() + $t.Substring(1).ToLower()
    return [Array]::IndexOf($DayNames, $normalized)
}

function Convert-TimeTokenTo24H([string]$token) {
    # Accepts "09:00", "9:00", "9AM", "9", "12AM", "12PM", "24:00", "24" etc.
    # (case-insensitive AM/PM, colon optional) and returns "HH:MM", or $null.
    $t = $token.Trim()
    if ($t -match '^(\d{1,2})(?::(\d{2}))?\s*([AaPp][Mm])?$') {
        $hour = [int]$matches[1]
        $minute = if ($matches[2]) { [int]$matches[2] } else { 0 }
        $ampm = $matches[3]
        if ($ampm) {
            if ($hour -lt 1 -or $hour -gt 12) { return $null }
            $ampmUpper = $ampm.ToUpper()
            if ($ampmUpper -eq 'AM') { if ($hour -eq 12) { $hour = 0 } }
            else { if ($hour -ne 12) { $hour += 12 } }
        } else {
            # No AM/PM suffix -- treat as already-24h, "24" allowed only as an
            # end-of-day marker (this tool's own output uses "24:00").
            if ($hour -lt 0 -or $hour -gt 24) { return $null }
        }
        if ($minute -lt 0 -or $minute -gt 59) { return $null }
        return ('{0:D2}:{1:D2}' -f $hour, $minute)
    }
    return $null
}

function ConvertTo-GridFromTimesString([string]$timesValue) {
    # Builds a fresh 7x24 bool[,] from a times value WITHOUT touching the
    # global $Grid -- used both to actually load a schedule and to validate
    # pasted text before committing it to a slot.
    $g = New-Object 'bool[,]' 7,24
    if ($timesValue -eq 'all') {
        for ($d = 0; $d -lt 7; $d++) { for ($h = 0; $h -lt 24; $h++) { $g[$d,$h] = $true } }
        return ,$g
    }
    if ([string]::IsNullOrWhiteSpace($timesValue)) { return ,$g }

    foreach ($entry in ($timesValue -split ';')) {
        if (-not $entry.Trim()) { continue }
        $parts = $entry -split ',', 2
        if ($parts.Count -lt 2) { continue }
        $dayPart = $parts[0].Trim()
        $timePart = $parts[1].Trim()

        if ($dayPart -match '^([A-Za-z]+)\s*-\s*([A-Za-z]+)$') {
            $d1 = Normalize-DayToken $matches[1]
            $d2 = Normalize-DayToken $matches[2]
        } else {
            $d1 = Normalize-DayToken $dayPart
            $d2 = $d1
        }
        if ($d1 -lt 0 -or $d2 -lt 0) { continue }

        $timeTokens = $timePart -split '-', 2
        if ($timeTokens.Count -lt 2) { continue }
        $startStr = Convert-TimeTokenTo24H $timeTokens[0]
        $endStr   = Convert-TimeTokenTo24H $timeTokens[1]
        if (-not $startStr -or -not $endStr) { continue }
        $startH = [int]($startStr -split ':')[0]
        $endH   = [int]($endStr -split ':')[0]

        if ($endH -eq 0) {
            # "...-12AM" / "...-0:00" means through midnight, same day.
            $endH = 24
        }

        if ($endH -gt $startH) {
            # Normal same-day range.
            for ($d = $d1; $d -le $d2; $d++) {
                for ($h = $startH; $h -lt $endH; $h++) { $g[$d, $h] = $true }
            }
        } else {
            # Wraps past midnight into the next day, e.g. "22:00-01:00" -- a
            # real pattern used in the checklist sheet (single token spanning
            # midnight, not split into two day entries). Set the rest of the
            # start day plus the early hours of the following day.
            for ($d = $d1; $d -le $d2; $d++) {
                for ($h = $startH; $h -lt 24; $h++) { $g[$d, $h] = $true }
                $nextDay = ($d + 1) % 7
                for ($h = 0; $h -lt $endH; $h++) { $g[$nextDay, $h] = $true }
            }
        }
    }
    return ,$g
}

function Parse-TimesStringToGrid([string]$timesValue) {
    # Fills the global $Grid (used when actually loading a schedule into the
    # editor). See ConvertTo-GridFromTimesString for the actual parsing.
    $g = ConvertTo-GridFromTimesString $timesValue
    for ($d = 0; $d -lt 7; $d++) { for ($h = 0; $h -lt 24; $h++) { $Grid[$d, $h] = $g[$d, $h] } }
}

function Extract-TimesValueFromText([string]$text) {
    # Pulls the times value out of a pasted full command like
    # net user students /times:M-F,09:00-17:00
    # or accepts a bare times value pasted on its own. Strips wrapping quotes.
    $t = $text.Trim()
    if ($t -match '(?i)/times:\s*"?([^"\r\n]*?)"?\s*$') { return $matches[1].Trim() }
    return $t.Trim('"', "'", ' ')
}

function Extract-AccountFromText([string]$text) {
    # The account name in a pasted full command, e.g. "Student" from
    # net user Student /times:M-F,09:00-17:00 -- or $null when the paste is
    # just a bare times value, or the sheet's template placeholder
    # ("net user *name of user* /times:..."), which isn't a real account.
    if ($text -match '(?i)\bnet\s+user\s+("?)([^"\s/]+)\1\s+(?:\S.*?\s)?/times:') {
        $name = $matches[2]
        if ($name -notmatch '[\*<>]') { return $name }
    }
    return $null
}

function Get-DefaultStudentAccount {
    # Pre-fill the account box with whichever common student account actually
    # exists on THIS PC (the sheet shows clients using Student, Students,
    # Bochurim...), instead of assuming one. Falls back to 'students'.
    foreach ($candidate in @('Students', 'Student', 'Stu', 'Bochurim')) {
        try { if ([ADSI]::Exists("WinNT://./$candidate,user")) { return $candidate } } catch { }
    }
    return 'students'
}

# ---------------------------------------------------------------------------
# Logging
# ---------------------------------------------------------------------------

function Write-LogEntry([string]$userName, [string]$oldTimes, [string]$newTimes, [string]$note) {
    if (-not (Test-Path $LogPath)) {
        'Timestamp,PCName,UserName,PreviousTimes,NewTimes,Note' | Out-File -FilePath $LogPath -Encoding utf8
    }
    $ts = Get-Date -Format 'yyyy-MM-dd HH:mm:ss'
    $pc = $env:COMPUTERNAME
    $escNote = $note -replace '"', '""'
    $line = '"{0}","{1}","{2}","{3}","{4}","{5}"' -f $ts, $pc, $userName, $oldTimes, $newTimes, $escNote
    Add-Content -Path $LogPath -Value $line
}

# ---------------------------------------------------------------------------
# Saved-schedule slots (per Windows account, stored locally as JSON)
# ---------------------------------------------------------------------------

function Load-ProfilesData {
    if (Test-Path $ProfilesPath) {
        try {
            $raw = Get-Content -Path $ProfilesPath -Raw
            if ($raw) { return ($raw | ConvertFrom-Json) }
        } catch { }
    }
    return (New-Object PSObject)
}

function Save-ProfilesData($data) {
    $data | ConvertTo-Json -Depth 6 | Out-File -FilePath $ProfilesPath -Encoding utf8
}

$SharedSlotOwner = '_ALL_ACCOUNTS'

function Get-UserSlot([string]$userName, [string]$slotKey) {
    # Quick slots (Slot1..) are shared by every account on the PC; only AutoBackup (the saved schedule
    # for an unlock) is per account. Older per-account slots are still read when no shared one exists.
    $data = Load-ProfilesData
    if ($slotKey -like 'Slot*' -and $data.PSObject.Properties[$SharedSlotOwner] -and $data.$SharedSlotOwner.PSObject.Properties[$slotKey]) {
        return $data.$SharedSlotOwner.$slotKey
    }
    if (-not $data.PSObject.Properties[$userName]) { return $null }
    $userObj = $data.$userName
    if (-not $userObj.PSObject.Properties[$slotKey]) { return $null }
    return $userObj.$slotKey
}

function Set-UserSlot([string]$userName, [string]$slotKey, [string]$timesValue, [string]$name = '') {
    if ($slotKey -like 'Slot*') { $userName = $SharedSlotOwner }   # slots are for ALL accounts
    $data = Load-ProfilesData
    if (-not $data.PSObject.Properties[$userName]) {
        $data | Add-Member -NotePropertyName $userName -NotePropertyValue (New-Object PSObject)
    }
    $userObj = $data.$userName
    $slotObj = [PSCustomObject]@{ TimesString = $timesValue; Timestamp = (Get-Date -Format 'yyyy-MM-dd HH:mm:ss'); Name = $name }
    if ($userObj.PSObject.Properties[$slotKey]) {
        $userObj.$slotKey = $slotObj
    } else {
        $userObj | Add-Member -NotePropertyName $slotKey -NotePropertyValue $slotObj
    }
    Save-ProfilesData $data
}

# ---------------------------------------------------------------------------
# GUI
# ---------------------------------------------------------------------------

$form = New-Object System.Windows.Forms.Form
$form.Text = "Student Times - Logon Hours Editor (v$ScriptVersion)"
# Default window fits as much of the full content (~1340x1010 for the Daily
# Use tab) as the actual screen allows, so a normal monitor shows everything
# with no scrolling. The form is resizable and each tab's content lives in
# its own AutoScroll panel, so on a genuinely small/odd screen, whatever
# doesn't fit is still reachable via scrollbars instead of being cut off.
$desiredContentSize = New-Object System.Drawing.Size(1340, 1010)
$workArea = [System.Windows.Forms.Screen]::PrimaryScreen.WorkingArea
$formWidth = [Math]::Min($desiredContentSize.Width, $workArea.Width - 40)
$formHeight = [Math]::Min($desiredContentSize.Height, $workArea.Height - 60)
$form.Size = New-Object System.Drawing.Size($formWidth, $formHeight)
$form.MinimumSize = New-Object System.Drawing.Size(500, 400)
$form.StartPosition = 'CenterScreen'
$form.Font = New-Object System.Drawing.Font('Segoe UI', 9)

$tabControl = New-Object System.Windows.Forms.TabControl
$tabControl.Dock = 'Fill'
$form.Controls.Add($tabControl)

$tabDaily = New-Object System.Windows.Forms.TabPage
$tabDaily.Text = 'Daily Use'
$tabControl.TabPages.Add($tabDaily)

$tabSetup = New-Object System.Windows.Forms.TabPage
$tabSetup.Text = 'Setup / Verify'
$tabControl.TabPages.Add($tabSetup)

$tabNetwork = New-Object System.Windows.Forms.TabPage
$tabNetwork.Text = 'Network Push'
$tabControl.TabPages.Add($tabNetwork)

$tabLock = New-Object System.Windows.Forms.TabPage
$tabLock.Text = 'Desktop Lock'
$tabControl.TabPages.Add($tabLock)

$lockPanel = New-Object System.Windows.Forms.Panel
$lockPanel.Dock = 'Fill'
$lockPanel.AutoScroll = $true
$tabLock.Controls.Add($lockPanel)

$mainPanel = New-Object System.Windows.Forms.Panel
$mainPanel.Dock = 'Fill'
$mainPanel.AutoScroll = $true
$tabDaily.Controls.Add($mainPanel)

$networkPanel = New-Object System.Windows.Forms.Panel
$networkPanel.Dock = 'Fill'
$networkPanel.AutoScroll = $true
$tabNetwork.Controls.Add($networkPanel)

$setupPanel = New-Object System.Windows.Forms.Panel
$setupPanel.Dock = 'Fill'
$setupPanel.AutoScroll = $true
$tabSetup.Controls.Add($setupPanel)

# --- Top row: username + load/save/clear buttons ---
$lblUser = New-Object System.Windows.Forms.Label
$lblUser.Text = 'Windows account name:'
$lblUser.Location = New-Object System.Drawing.Point(15, 18)
$lblUser.AutoSize = $true
$mainPanel.Controls.Add($lblUser)

$txtUser = New-Object System.Windows.Forms.TextBox
$txtUser.Text = Get-DefaultStudentAccount
$txtUser.Location = New-Object System.Drawing.Point(160, 15)
$txtUser.Size = New-Object System.Drawing.Size(150, 24)
$mainPanel.Controls.Add($txtUser)

$btnLoad = New-Object System.Windows.Forms.Button
$btnLoad.Text = 'Load Current Schedule From This PC'
$btnLoad.Location = New-Object System.Drawing.Point(325, 13)
$btnLoad.Size = New-Object System.Drawing.Size(230, 28)
$mainPanel.Controls.Add($btnLoad)

$btnAllowAll = New-Object System.Windows.Forms.Button
$btnAllowAll.Text = 'Allow All (24/7)'
$btnAllowAll.Location = New-Object System.Drawing.Point(565, 13)
$btnAllowAll.Size = New-Object System.Drawing.Size(120, 28)
$mainPanel.Controls.Add($btnAllowAll)

$btnClear = New-Object System.Windows.Forms.Button
$btnClear.Text = 'Block All'
$btnClear.Location = New-Object System.Drawing.Point(695, 13)
$btnClear.Size = New-Object System.Drawing.Size(90, 28)
$mainPanel.Controls.Add($btnClear)

$btnAccounts = New-Object System.Windows.Forms.Button
$btnAccounts.Text = 'Accounts on this PC...'
$btnAccounts.Location = New-Object System.Drawing.Point(795, 13)
$btnAccounts.Size = New-Object System.Drawing.Size(170, 28)
$mainPanel.Controls.Add($btnAccounts)

$lblApplyTo = New-Object System.Windows.Forms.Label
$lblApplyTo.Text = ''
$lblApplyTo.Location = New-Object System.Drawing.Point(975, 19)
$lblApplyTo.AutoSize = $true
$mainPanel.Controls.Add($lblApplyTo)

$lblHint = New-Object System.Windows.Forms.Label
$lblHint.Text = 'Click/drag cells to toggle allowed (green) vs blocked (gray) hours. Rows = days, columns = hours (local time). Use Copy/Paste on the right to duplicate a whole day onto another.'
$lblHint.Location = New-Object System.Drawing.Point(15, 48)
$lblHint.AutoSize = $true
$lblHint.ForeColor = [System.Drawing.Color]::DimGray
$mainPanel.Controls.Add($lblHint)

$lblDayClipboard = New-Object System.Windows.Forms.Label
$lblDayClipboard.Text = ''
$lblDayClipboard.Location = New-Object System.Drawing.Point(15, 63)
$lblDayClipboard.AutoSize = $true
$lblDayClipboard.ForeColor = [System.Drawing.Color]::DarkSlateBlue
$mainPanel.Controls.Add($lblDayClipboard)

# --- Grid of toggle buttons ---
$gridTop = 95
$cellW = 40
$cellH = 40
$labelColW = 40

$hourLabelsPanel = New-Object System.Windows.Forms.Panel
$hourLabelsPanel.Location = New-Object System.Drawing.Point($labelColW, ($gridTop - 20))
$hourLabelsPanel.Size = New-Object System.Drawing.Size((24 * $cellW), 20)
$mainPanel.Controls.Add($hourLabelsPanel)
for ($h = 0; $h -lt 24; $h++) {
    $lbl = New-Object System.Windows.Forms.Label
    $lbl.Text = "$h"
    $lbl.TextAlign = 'MiddleCenter'
    $lbl.Location = New-Object System.Drawing.Point(($h * $cellW), 0)
    $lbl.Size = New-Object System.Drawing.Size($cellW, 20)
    $lbl.Font = New-Object System.Drawing.Font('Segoe UI', 7)
    $hourLabelsPanel.Controls.Add($lbl)
}

$cellButtons = New-Object 'System.Windows.Forms.Button[,]' 7,24
$script:isDragging = $false
$script:dragAllow = $true
$script:dayClipboard = $null
$script:dayClipboardSource = ''

function Set-CellVisual($btn, [bool]$allowed) {
    if ($allowed) {
        $btn.BackColor = [System.Drawing.Color]::MediumSeaGreen
    } else {
        $btn.BackColor = [System.Drawing.Color]::Gainsboro
    }
}

for ($d = 0; $d -lt 7; $d++) {
    $dayLbl = New-Object System.Windows.Forms.Label
    $dayLbl.Text = $DayNames[$d]
    $dayLbl.TextAlign = 'MiddleCenter'
    $dayLbl.Location = New-Object System.Drawing.Point(0, ($gridTop + $d * $cellH))
    $dayLbl.Size = New-Object System.Drawing.Size($labelColW, $cellH)
    $mainPanel.Controls.Add($dayLbl)

    $rowRightEdge = $labelColW + 24 * $cellW
    $btnCopyDay = New-Object System.Windows.Forms.Button
    $btnCopyDay.Text = 'Copy'
    $btnCopyDay.Location = New-Object System.Drawing.Point(($rowRightEdge + 10), ($gridTop + $d * $cellH + 3))
    $btnCopyDay.Size = New-Object System.Drawing.Size(60, 34)
    $btnCopyDay.Tag = $d
    $mainPanel.Controls.Add($btnCopyDay)

    $btnPasteDay = New-Object System.Windows.Forms.Button
    $btnPasteDay.Text = 'Paste'
    $btnPasteDay.Location = New-Object System.Drawing.Point(($rowRightEdge + 75), ($gridTop + $d * $cellH + 3))
    $btnPasteDay.Size = New-Object System.Drawing.Size(60, 34)
    $btnPasteDay.Tag = $d
    $mainPanel.Controls.Add($btnPasteDay)

    $btnCopyDay.Add_Click({
        param($s, $e)
        $srcDay = $s.Tag
        $pattern = New-Object 'bool[]' 24
        for ($h = 0; $h -lt 24; $h++) { $pattern[$h] = $Grid[$srcDay, $h] }
        $script:dayClipboard = $pattern
        $script:dayClipboardSource = $DayNames[$srcDay]
        $lblDayClipboard.Text = "Day clipboard: $($script:dayClipboardSource)'s hours copied - click Paste on another day to apply."
    })
    $btnPasteDay.Add_Click({
        param($s, $e)
        $destDay = $s.Tag
        if ($null -eq $script:dayClipboard) {
            [System.Windows.Forms.MessageBox]::Show('Click Copy on a day first.', 'Nothing copied yet') | Out-Null
            return
        }
        for ($h = 0; $h -lt 24; $h++) {
            $Grid[$destDay, $h] = $script:dayClipboard[$h]
            Set-CellVisual $cellButtons[$destDay, $h] $Grid[$destDay, $h]
        }
        $lblDayClipboard.Text = "Pasted $($script:dayClipboardSource)'s hours onto $($DayNames[$destDay])."
    })

    for ($h = 0; $h -lt 24; $h++) {
        $btn = New-Object System.Windows.Forms.Button
        $btn.Location = New-Object System.Drawing.Point(($labelColW + $h * $cellW), ($gridTop + $d * $cellH))
        $btn.Size = New-Object System.Drawing.Size($cellW, $cellH)
        $btn.FlatStyle = 'Flat'
        $btn.FlatAppearance.BorderSize = 1
        $btn.Tag = @{ Day = $d; Hour = $h }
        Set-CellVisual $btn $false

        $btn.Add_MouseDown({
            param($s, $e)
            $tag = $s.Tag
            $newVal = -not $Grid[$tag.Day, $tag.Hour]
            $Grid[$tag.Day, $tag.Hour] = $newVal
            Set-CellVisual $s $newVal
            $script:isDragging = $true
            $script:dragAllow = $newVal
        })
        $btn.Add_MouseEnter({
            param($s, $e)
            if ($script:isDragging) {
                $tag = $s.Tag
                $Grid[$tag.Day, $tag.Hour] = $script:dragAllow
                Set-CellVisual $s $script:dragAllow
            }
        })
        $mainPanel.Controls.Add($btn)
        $cellButtons[$d, $h] = $btn
    }
}
$form.Add_MouseUp({ $script:isDragging = $false })
# Buttons swallow MouseUp when the cursor is over them; catch it there too.
for ($d = 0; $d -lt 7; $d++) {
    for ($h = 0; $h -lt 24; $h++) {
        $cellButtons[$d, $h].Add_MouseUp({ $script:isDragging = $false })
    }
}

function Refresh-GridVisuals {
    for ($d = 0; $d -lt 7; $d++) {
        for ($h = 0; $h -lt 24; $h++) {
            Set-CellVisual $cellButtons[$d, $h] $Grid[$d, $h]
        }
    }
}

# --- Output area ---
$outTop = $gridTop + 7 * $cellH + 20

$lblOut = New-Object System.Windows.Forms.Label
$lblOut.Text = 'Generated command (elevated Command Prompt):'
$lblOut.Location = New-Object System.Drawing.Point(15, $outTop)
$lblOut.AutoSize = $true
$mainPanel.Controls.Add($lblOut)

$txtOut = New-Object System.Windows.Forms.TextBox
$txtOut.Location = New-Object System.Drawing.Point(15, ($outTop + 20))
$txtOut.Size = New-Object System.Drawing.Size(1130, 50)
$txtOut.Multiline = $true
$txtOut.ReadOnly = $true
$txtOut.BackColor = [System.Drawing.Color]::White
$txtOut.Font = New-Object System.Drawing.Font('Consolas', 10)
$mainPanel.Controls.Add($txtOut)

$btnGenerate = New-Object System.Windows.Forms.Button
$btnGenerate.Text = 'Generate + Copy Command'
$btnGenerate.Location = New-Object System.Drawing.Point(15, ($outTop + 78))
$btnGenerate.Size = New-Object System.Drawing.Size(200, 30)
$mainPanel.Controls.Add($btnGenerate)

$chkRunNow = New-Object System.Windows.Forms.CheckBox
$chkRunNow.Text = 'Run it immediately here instead of copying (this PC only)'
$chkRunNow.Location = New-Object System.Drawing.Point(15, ($outTop + 110))
$chkRunNow.AutoSize = $true
$chkRunNow.ForeColor = [System.Drawing.Color]::DarkRed
$mainPanel.Controls.Add($chkRunNow)

$btnCopyRow = New-Object System.Windows.Forms.Button
$btnCopyRow.Text = 'Copy Row For Checklist Sheet'
$btnCopyRow.Location = New-Object System.Drawing.Point(225, ($outTop + 78))
$btnCopyRow.Size = New-Object System.Drawing.Size(220, 30)
$btnCopyRow.Enabled = $false
$mainPanel.Controls.Add($btnCopyRow)

$lblLog = New-Object System.Windows.Forms.Label
$lblLog.Text = "Every generate is logged locally to: $LogPath"
$lblLog.Location = New-Object System.Drawing.Point(460, ($outTop + 85))
$lblLog.AutoSize = $true
$lblLog.ForeColor = [System.Drawing.Color]::DimGray
$mainPanel.Controls.Add($lblLog)

$btnOpenDebugLog = New-Object System.Windows.Forms.Button
$btnOpenDebugLog.Text = 'Open Debug Log'
$btnOpenDebugLog.Location = New-Object System.Drawing.Point(950, ($outTop + 78))
$btnOpenDebugLog.Size = New-Object System.Drawing.Size(160, 30)
$mainPanel.Controls.Add($btnOpenDebugLog)
$btnOpenDebugLog.Add_Click({
    if (-not (Test-Path $DebugLogPath)) {
        'Debug log created.' | Out-File -FilePath $DebugLogPath -Encoding utf8
    }
    Start-Process notepad.exe $DebugLogPath
})

$script:lastOldTimes = ''
$script:lastNewTimes = ''

# --- One-click temporary override section ---
$overrideTop = $outTop + 145

$sep1 = New-Object System.Windows.Forms.Label
$sep1.BorderStyle = 'Fixed3D'
$sep1.Location = New-Object System.Drawing.Point(15, ($overrideTop - 12))
$sep1.Size = New-Object System.Drawing.Size(1130, 2)
$mainPanel.Controls.Add($sep1)

$lblOverride = New-Object System.Windows.Forms.Label
$lblOverride.Text = 'Temporary unlock - applies IMMEDIATELY, then re-locks BY ITSELF at midnight (puts the saved schedule and the logout tasks back). Nothing more to do:'
$lblOverride.Location = New-Object System.Drawing.Point(15, $overrideTop)
$lblOverride.AutoSize = $true
$mainPanel.Controls.Add($lblOverride)

$btnUnlockToday = New-Object System.Windows.Forms.Button
$btnUnlockToday.Text = 'Unlock Today (re-locks at midnight)'
$btnUnlockToday.Location = New-Object System.Drawing.Point(15, ($overrideTop + 22))
$btnUnlockToday.Size = New-Object System.Drawing.Size(235, 34)
$btnUnlockToday.BackColor = [System.Drawing.Color]::Khaki
$mainPanel.Controls.Add($btnUnlockToday)

$btnUnlock2Days = New-Object System.Windows.Forms.Button
$btnUnlock2Days.Text = 'Unlock 2 Days'
$btnUnlock2Days.Location = New-Object System.Drawing.Point(255, ($overrideTop + 22))
$btnUnlock2Days.Size = New-Object System.Drawing.Size(120, 34)
$btnUnlock2Days.BackColor = [System.Drawing.Color]::Khaki
$mainPanel.Controls.Add($btnUnlock2Days)

$btnUnlock3Days = New-Object System.Windows.Forms.Button
$btnUnlock3Days.Text = 'Unlock 3 Days'
$btnUnlock3Days.Location = New-Object System.Drawing.Point(380, ($overrideTop + 22))
$btnUnlock3Days.Size = New-Object System.Drawing.Size(120, 34)
$btnUnlock3Days.BackColor = [System.Drawing.Color]::Khaki
$mainPanel.Controls.Add($btnUnlock3Days)

$btnRestorePrevious = New-Object System.Windows.Forms.Button
$btnRestorePrevious.Text = 'Re-lock Now (restore previous)'
$btnRestorePrevious.Location = New-Object System.Drawing.Point(510, ($overrideTop + 22))
$btnRestorePrevious.Size = New-Object System.Drawing.Size(210, 34)
$btnRestorePrevious.Enabled = $false
$mainPanel.Controls.Add($btnRestorePrevious)

$lblAutoBackupInfo = New-Object System.Windows.Forms.Label
$lblAutoBackupInfo.Text = 'No auto-saved previous schedule yet for this account.'
$lblAutoBackupInfo.Location = New-Object System.Drawing.Point(730, ($overrideTop + 24))
$lblAutoBackupInfo.AutoSize = $true
$lblAutoBackupInfo.ForeColor = [System.Drawing.Color]::DimGray
$mainPanel.Controls.Add($lblAutoBackupInfo)

# --- Quick schedule slots (2-3 named saved versions per account) ---
$slotsTop = $overrideTop + 75

$sep2 = New-Object System.Windows.Forms.Label
$sep2.BorderStyle = 'Fixed3D'
$sep2.Location = New-Object System.Drawing.Point(15, ($slotsTop - 12))
$sep2.Size = New-Object System.Drawing.Size(1130, 2)
$mainPanel.Controls.Add($sep2)

$lblSlots = New-Object System.Windows.Forms.Label
$lblSlots.Text = 'Quick schedule slots (save a couple of versions per account, reload either in one click). Paste a net user command or bare times value in the box to fill a slot straight from the checklist sheet, then click Save Text:'
$lblSlots.Location = New-Object System.Drawing.Point(15, $slotsTop)
$lblSlots.Size = New-Object System.Drawing.Size(1290, 32)
$mainPanel.Controls.Add($lblSlots)

$slotNameBoxes = @()
$slotTimestampLabels = @()
$slotTextBoxes = @()
$slotCount = 3
$slotsRowStart = $slotsTop + 36
for ($i = 1; $i -le $slotCount; $i++) {
    $rowY = $slotsRowStart + (($i - 1) * 38)

    $txtSlotName = New-Object System.Windows.Forms.TextBox
    $txtSlotName.Text = "Slot $i"
    $txtSlotName.Location = New-Object System.Drawing.Point(15, ($rowY + 3))
    $txtSlotName.Size = New-Object System.Drawing.Size(140, 24)
    $mainPanel.Controls.Add($txtSlotName)
    $slotNameBoxes += $txtSlotName

    $lblSlotTimestamp = New-Object System.Windows.Forms.Label
    $lblSlotTimestamp.Text = '(empty)'
    $lblSlotTimestamp.Location = New-Object System.Drawing.Point(160, ($rowY + 7))
    $lblSlotTimestamp.Size = New-Object System.Drawing.Size(160, 22)
    $lblSlotTimestamp.ForeColor = [System.Drawing.Color]::DimGray
    $mainPanel.Controls.Add($lblSlotTimestamp)
    $slotTimestampLabels += $lblSlotTimestamp

    $txtSlot = New-Object System.Windows.Forms.TextBox
    $txtSlot.Location = New-Object System.Drawing.Point(330, ($rowY + 3))
    $txtSlot.Size = New-Object System.Drawing.Size(390, 24)
    $txtSlot.Font = New-Object System.Drawing.Font('Consolas', 9)
    $mainPanel.Controls.Add($txtSlot)
    $slotTextBoxes += $txtSlot

    $btnSlotSaveText = New-Object System.Windows.Forms.Button
    $btnSlotSaveText.Text = "Save Text -> Slot $i"
    $btnSlotSaveText.Location = New-Object System.Drawing.Point(730, $rowY)
    $btnSlotSaveText.Size = New-Object System.Drawing.Size(140, 30)
    $btnSlotSaveText.Tag = $i
    $mainPanel.Controls.Add($btnSlotSaveText)

    $btnSlotSave = New-Object System.Windows.Forms.Button
    $btnSlotSave.Text = "Save Grid -> Slot $i"
    $btnSlotSave.Location = New-Object System.Drawing.Point(875, $rowY)
    $btnSlotSave.Size = New-Object System.Drawing.Size(150, 30)
    $btnSlotSave.Tag = $i
    $mainPanel.Controls.Add($btnSlotSave)

    $btnSlotLoad = New-Object System.Windows.Forms.Button
    $btnSlotLoad.Text = "Load Slot $i"
    $btnSlotLoad.Location = New-Object System.Drawing.Point(1030, $rowY)
    $btnSlotLoad.Size = New-Object System.Drawing.Size(120, 30)
    $btnSlotLoad.Tag = $i
    $mainPanel.Controls.Add($btnSlotLoad)

    $btnSlotSaveText.Add_Click({
        param($s, $e)
        $slotIndex = $s.Tag
        $userName = $txtUser.Text.Trim()
        if (-not $userName) {
            [System.Windows.Forms.MessageBox]::Show('Enter a Windows account name first.', 'Missing account') | Out-Null
            return
        }
        $pasted = $slotTextBoxes[$slotIndex - 1].Text
        if (-not $pasted.Trim()) {
            [System.Windows.Forms.MessageBox]::Show("Paste a net user command (e.g. `"net user $userName /times:M-F,09:00-17:00`") or just the times value (e.g. `"M-F,09:00-17:00`") into the box for Slot $slotIndex first.", 'Nothing to save') | Out-Null
            return
        }
        $timesValue = Extract-TimesValueFromText $pasted
        $testGrid = ConvertTo-GridFromTimesString $timesValue
        $anyAllowed = $false
        for ($d = 0; $d -lt 7; $d++) { for ($h = 0; $h -lt 24; $h++) { if ($testGrid[$d,$h]) { $anyAllowed = $true } } }
        $looksValid = ($timesValue -eq 'all') -or ($timesValue -eq '') -or $anyAllowed
        if (-not $looksValid) {
            $proceed = [System.Windows.Forms.MessageBox]::Show(
                "Couldn't recognize any day/time pattern in this:`n`n$timesValue`n`nCheck the day abbreviations (Su/M/T/W/Th/F/Sa) and time format (e.g. 9AM, 09:00). Save it to Slot $slotIndex anyway?",
                'Does not look valid', 'YesNo', 'Warning')
            if ($proceed -ne 'Yes') { return }
        }
        $slotName = $slotNameBoxes[$slotIndex - 1].Text.Trim()

        # Slots are shared by ALL accounts, so the account named in a pasted command is irrelevant to the slot.
        # It must NOT change the account box either: that used to switch the account and drop the ticked
        # accounts, which is why pasting another user's command "did not save".
        $pastedAccount = Extract-AccountFromText $pasted
        $switchNote = ''
        if ($pastedAccount -and $pastedAccount -cne $userName) {
            $switchNote = "`n`n(The pasted command named '$pastedAccount' - only its times were used. Slots work for every account.)"
            Write-DebugLog "Slot $slotIndex paste: command named '$pastedAccount', account box stays '$userName'"
        }

        Set-UserSlot -userName $userName -slotKey "Slot$slotIndex" -timesValue $timesValue -name $slotName
        Write-DebugLog "Slot $slotIndex ('$slotName') saved from pasted text for '$userName': $timesValue"
        Update-SlotLabels
        [System.Windows.Forms.MessageBox]::Show("Saved to '$slotName' (for all accounts):`n`n/times:$timesValue$switchNote", 'Saved') | Out-Null
    })

    $btnSlotSave.Add_Click({
        param($s, $e)
        $slotIndex = $s.Tag
        $userName = $txtUser.Text.Trim()
        if (-not $userName) {
            [System.Windows.Forms.MessageBox]::Show('Enter a Windows account name first.', 'Missing account') | Out-Null
            return
        }
        $timesValue = Build-TimesString
        $slotName = $slotNameBoxes[$slotIndex - 1].Text.Trim()
        Set-UserSlot -userName $userName -slotKey "Slot$slotIndex" -timesValue $timesValue -name $slotName
        Update-SlotLabels
        [System.Windows.Forms.MessageBox]::Show("Current grid saved to '$slotName' for '$userName'.", 'Saved') | Out-Null
    })

    $btnSlotLoad.Add_Click({
        param($s, $e)
        $slotIndex = $s.Tag
        $userName = $txtUser.Text.Trim()
        if (-not $userName) {
            [System.Windows.Forms.MessageBox]::Show('Enter a Windows account name first.', 'Missing account') | Out-Null
            return
        }
        $slot = Get-UserSlot -userName $userName -slotKey "Slot$slotIndex"
        $slotName = $slotNameBoxes[$slotIndex - 1].Text.Trim()
        if (-not $slotName) { $slotName = "Slot $slotIndex" }
        if ($null -eq $slot) {
            [System.Windows.Forms.MessageBox]::Show("'$slotName' is empty for '$userName'.", 'Empty slot') | Out-Null
            return
        }
        $oldSnapshot = Build-TimesString
        Parse-TimesStringToGrid $slot.TimesString
        Refresh-GridVisuals
        $script:lastOldTimes = $oldSnapshot
        Apply-Grid-AndEmit -userName $userName -actionLabel "Load '$slotName'"
    })
}

# --- Schedule a future change (e.g. "switch to Bein Hazmanim hours on Sept 22,
# switch back on Oct 18") via a one-time Windows Scheduled Task, so nobody has
# to be there or remember to do it on the day. ---
$scheduleTop = $slotsRowStart + ($slotCount * 38) + 20

$sep3 = New-Object System.Windows.Forms.Label
$sep3.BorderStyle = 'Fixed3D'
$sep3.Location = New-Object System.Drawing.Point(15, ($scheduleTop - 12))
$sep3.Size = New-Object System.Drawing.Size(1290, 2)
$mainPanel.Controls.Add($sep3)

$lblSchedule = New-Object System.Windows.Forms.Label
$lblSchedule.Text = 'Schedule this exact grid for a future date/time (runs unattended on this PC via Task Scheduler, even if nobody is logged in):'
$lblSchedule.Location = New-Object System.Drawing.Point(15, $scheduleTop)
$lblSchedule.Size = New-Object System.Drawing.Size(1290, 20)
$mainPanel.Controls.Add($lblSchedule)

$dtpSchedule = New-Object System.Windows.Forms.DateTimePicker
$dtpSchedule.Format = 'Custom'
$dtpSchedule.CustomFormat = 'ddd MM/dd/yyyy   HH:mm'
$dtpSchedule.ShowUpDown = $true
$dtpSchedule.Value = (Get-Date).AddDays(1).Date.AddHours(8)
$dtpSchedule.Location = New-Object System.Drawing.Point(15, ($scheduleTop + 24))
$dtpSchedule.Size = New-Object System.Drawing.Size(220, 26)
$mainPanel.Controls.Add($dtpSchedule)

$txtScheduleLabel = New-Object System.Windows.Forms.TextBox
$txtScheduleLabel.Text = ''
$txtScheduleLabel.Location = New-Object System.Drawing.Point(245, ($scheduleTop + 25))
$txtScheduleLabel.Size = New-Object System.Drawing.Size(220, 24)
$mainPanel.Controls.Add($txtScheduleLabel)

$lblScheduleLabelHint = New-Object System.Windows.Forms.Label
$lblScheduleLabelHint.Text = '(optional label, e.g. "Bein Hazmanim")'
$lblScheduleLabelHint.Location = New-Object System.Drawing.Point(245, ($scheduleTop + 50))
$lblScheduleLabelHint.AutoSize = $true
$lblScheduleLabelHint.ForeColor = [System.Drawing.Color]::DimGray
$mainPanel.Controls.Add($lblScheduleLabelHint)

$btnScheduleThis = New-Object System.Windows.Forms.Button
$btnScheduleThis.Text = 'Schedule This Change'
$btnScheduleThis.Location = New-Object System.Drawing.Point(475, ($scheduleTop + 22))
$btnScheduleThis.Size = New-Object System.Drawing.Size(180, 30)
$mainPanel.Controls.Add($btnScheduleThis)

$lstScheduledTasks = New-Object System.Windows.Forms.ListBox
$lstScheduledTasks.Location = New-Object System.Drawing.Point(15, ($scheduleTop + 60))
$lstScheduledTasks.Size = New-Object System.Drawing.Size(1290, 90)
$lstScheduledTasks.Font = New-Object System.Drawing.Font('Consolas', 9)
$lstScheduledTasks.HorizontalScrollbar = $true
$mainPanel.Controls.Add($lstScheduledTasks)

$btnRefreshTasks = New-Object System.Windows.Forms.Button
$btnRefreshTasks.Text = 'Refresh List'
$btnRefreshTasks.Location = New-Object System.Drawing.Point(15, ($scheduleTop + 155))
$btnRefreshTasks.Size = New-Object System.Drawing.Size(140, 30)
$mainPanel.Controls.Add($btnRefreshTasks)

$btnDeleteTask = New-Object System.Windows.Forms.Button
$btnDeleteTask.Text = 'Delete Selected Scheduled Task'
$btnDeleteTask.Location = New-Object System.Drawing.Point(165, ($scheduleTop + 155))
$btnDeleteTask.Size = New-Object System.Drawing.Size(220, 30)
$mainPanel.Controls.Add($btnDeleteTask)

function Register-FutureNetUserChange([string]$userName, [string]$timesValue, [datetime]$when, [string]$label, $taskGrid = $null) {
    # Creates a one-time Scheduled Task, running as SYSTEM (so it fires with
    # nobody logged in and needs no stored password), that runs the exact
    # net.exe command at $when. Returns {Success, Output, TaskName}.
    $safeLabel = ($label -replace '[^A-Za-z0-9]', '')
    $taskName = "StudentTimes-$userName-$($when.ToString('yyyyMMdd-HHmm'))"
    if ($safeLabel) { $taskName += "-$safeLabel" }

    $netCommand = 'net.exe user "' + $userName + '" "/times:' + $timesValue + '"'
    # The task runs a small script (a command line is limited to ~260 chars) that changes the
    # hours AND re-syncs the weekly logout tasks at that moment.
    $scriptDir = Join-Path $env:ProgramData 'StudentTimes'
    if (-not (Test-Path $scriptDir)) { New-Item -ItemType Directory -Path $scriptDir -Force | Out-Null }
    [void](Protect-StudentTimesDir)   # these scripts run as SYSTEM: ordinary users must only be able to read them
    $scriptPath = Join-Path $scriptDir ($taskName + '.ps1')
    $gridForTasks = $Grid
    if ($null -ne $taskGrid) { $gridForTasks = $taskGrid }   # a plain if-expression would flatten the 2-D array
    $groupLiterals = @(Get-LogoutKickoffGroups $gridForTasks | ForEach-Object { "[pscustomobject]@{ TimeHHmm = '$($_.TimeHHmm)'; Days = @($($_.Days -join ',')) }" })
    $actionName = if ($cmbLogoutAction.SelectedItem -eq 'Full log off') { 'LogOff' } else { 'Lock' }
    $scriptLines = @(
        "`$log = Join-Path `$PSScriptRoot '$taskName.log'",
        "`"`$(Get-Date -Format s) net user: `$((& net.exe user '$userName' '/times:$timesValue' 2>&1) -join ' ')`" | Out-File -Append -FilePath `$log -Encoding UTF8",
        '$sync = {', $SyncLogoutTasksBlock.ToString(), '}',
        "`$groups = @($($groupLiterals -join ', '))",
        "& `$sync '$userName' `$groups '$actionName' '' | Out-File -Append -FilePath `$log -Encoding UTF8"
    )
    Set-Content -Path $scriptPath -Value $scriptLines -Encoding UTF8
    $trCommand = 'powershell.exe -NoProfile -ExecutionPolicy Bypass -File "' + $scriptPath + '"'
    $trEscaped = $trCommand -replace '"', '\"'
    $dateStr = $when.ToString('MM/dd/yyyy')
    $timeStr = $when.ToString('HH:mm')

    $psi = New-Object System.Diagnostics.ProcessStartInfo
    $psi.FileName = 'schtasks.exe'
    $psi.Arguments = '/create /tn "' + $taskName + '" /tr "' + $trEscaped + '" /sc once /sd ' + $dateStr + ' /st ' + $timeStr + ' /ru SYSTEM /rl HIGHEST /f'
    $psi.RedirectStandardOutput = $true
    $psi.RedirectStandardError = $true
    $psi.UseShellExecute = $false
    $psi.CreateNoWindow = $true
    $proc = New-Object System.Diagnostics.Process
    $proc.StartInfo = $psi
    [void]$proc.Start()
    $stdout = $proc.StandardOutput.ReadToEnd()
    $stderr = $proc.StandardError.ReadToEnd()
    $proc.WaitForExit()
    return [PSCustomObject]@{
        Success  = ($proc.ExitCode -eq 0)
        Output   = ($stdout + $stderr).Trim()
        TaskName = $taskName
        Command  = $netCommand
    }
}

function Get-StudentTimesScheduledTasks {
    # Lists this tool's own one-time tasks (name prefix "StudentTimes-").
    $psi = New-Object System.Diagnostics.ProcessStartInfo
    $psi.FileName = 'schtasks.exe'
    $psi.Arguments = '/query /fo CSV /v'
    $psi.RedirectStandardOutput = $true
    $psi.RedirectStandardError = $true
    $psi.UseShellExecute = $false
    $psi.CreateNoWindow = $true
    $proc = New-Object System.Diagnostics.Process
    $proc.StartInfo = $psi
    [void]$proc.Start()
    $stdout = $proc.StandardOutput.ReadToEnd()
    $proc.WaitForExit()
    if (-not $stdout.Trim()) { return @() }
    $rows = $stdout | ConvertFrom-Csv
    return $rows | Where-Object { $_.TaskName -match 'StudentTimes-' } |
        ForEach-Object {
            [PSCustomObject]@{
                TaskName     = ($_.TaskName -replace '^\\', '')
                NextRunTime  = $_.'Next Run Time'
                Status       = $_.Status
                TaskToRun    = $_.'Task To Run'
            }
        }
}

function Refresh-ScheduledTasksList {
    $lstScheduledTasks.Items.Clear()
    $tasks = Get-StudentTimesScheduledTasks
    if ($tasks.Count -eq 0) {
        $lstScheduledTasks.Items.Add('(no scheduled changes on this PC yet)') | Out-Null
        return
    }
    foreach ($t in $tasks) {
        $lstScheduledTasks.Items.Add("$($t.TaskName)  |  next run: $($t.NextRunTime)  |  status: $($t.Status)  |  $($t.TaskToRun)") | Out-Null
    }
}

$btnScheduleThis.Add_Click({
    $targets = @(Get-TargetAccounts)
    if ($targets.Count -eq 0) { [System.Windows.Forms.MessageBox]::Show('Enter a Windows account name (or pick accounts) first.', 'Missing account') | Out-Null; return }
    $when = $dtpSchedule.Value
    if ($when -le (Get-Date)) { [System.Windows.Forms.MessageBox]::Show('Pick a date/time in the future.', 'Invalid time') | Out-Null; return }
    $timesValue = Build-TimesString
    $label = $txtScheduleLabel.Text.Trim()
    $commands = @($targets | ForEach-Object { "net user $_ /times:$timesValue" })
    $confirm = [System.Windows.Forms.MessageBox]::Show(
        "This will register a one-time Scheduled Task per account on $($env:COMPUTERNAME) that runs:`n`n$($commands -join "`n")`n`nat $($when.ToString('dddd, MMMM d, yyyy  h:mm tt'))`n`nContinue?",
        'Confirm scheduled change', 'YesNo', 'Question')
    if ($confirm -ne 'Yes') { return }
    $ok = @(); $failed = @()
    foreach ($acct in $targets) {
        $result = Register-FutureNetUserChange -userName $acct -timesValue $timesValue -when $when -label $label
        Write-DebugLog "Register-FutureNetUserChange $acct : task='$($result.TaskName)' success=$($result.Success) output=$($result.Output)"
        if ($result.Success) {
            Write-LogEntry -userName $acct -oldTimes $script:lastOldTimes -newTimes "net user $acct /times:$timesValue" -note "SCHEDULED for $($when.ToString('yyyy-MM-dd HH:mm')) (task: $($result.TaskName))"
            $ok += $acct
        } else { $failed += "$acct : $($result.Output)" }
    }
    Refresh-ScheduledTasksList
    if ($failed.Count -eq 0) { [System.Windows.Forms.MessageBox]::Show("Scheduled for $($ok -join ', ') at $($when.ToString('dddd, MMMM d, yyyy  h:mm tt')) on $($env:COMPUTERNAME), even if nobody is logged in at the time.", 'Scheduled') | Out-Null }
    else { [System.Windows.Forms.MessageBox]::Show("Scheduled: $($ok -join ', ')`n`nFAILED:`n$($failed -join "`n")", 'Some failed', 'OK', 'Warning') | Out-Null }
})

$btnRefreshTasks.Add_Click({ Refresh-ScheduledTasksList })

$btnDeleteTask.Add_Click({
    if ($lstScheduledTasks.SelectedIndex -lt 0) {
        [System.Windows.Forms.MessageBox]::Show('Select a scheduled task in the list first.', 'Nothing selected') | Out-Null
        return
    }
    $line = $lstScheduledTasks.SelectedItem.ToString()
    if ($line -notmatch '^(StudentTimes-\S+)\s+\|') {
        [System.Windows.Forms.MessageBox]::Show('Nothing to delete in this row.', 'Nothing selected') | Out-Null
        return
    }
    $taskName = $matches[1]
    $confirm = [System.Windows.Forms.MessageBox]::Show("Delete scheduled task '$taskName'? This cancels that future change.", 'Confirm delete', 'YesNo', 'Warning')
    if ($confirm -ne 'Yes') { return }

    $psi = New-Object System.Diagnostics.ProcessStartInfo
    $psi.FileName = 'schtasks.exe'
    $psi.Arguments = '/delete /tn "' + $taskName + '" /f'
    $psi.RedirectStandardOutput = $true
    $psi.RedirectStandardError = $true
    $psi.UseShellExecute = $false
    $psi.CreateNoWindow = $true
    $proc = New-Object System.Diagnostics.Process
    $proc.StartInfo = $psi
    [void]$proc.Start()
    $out = ($proc.StandardOutput.ReadToEnd() + $proc.StandardError.ReadToEnd())
    $proc.WaitForExit()
    Write-DebugLog "Deleted scheduled task '$taskName': exit=$($proc.ExitCode) output=$($out.Trim())"
    Refresh-ScheduledTasksList
})

# ---------------------------------------------------------------------------
# Initial setup steps 2-4 from our internal "Student Times Checklist":
#   2. Set net user (the grid + Generate/Run above already does this)
#   3. Set Task Scheduler to log the user out 1 minute after each lockout
#   4. Remove the account's time/time-zone-change rights (secpol.cpl)
#   5/6. Test that it worked -- the "Verify Everything" report below
# ---------------------------------------------------------------------------

function Get-LogoutKickoffGroups($grid) {
    # Finds every allowed->blocked transition in the grid and returns it as
    # "HH:01" (one minute after lockout, per the checklist's own convention),
    # grouped by identical time across days so one weekly task can cover
    # several days at once.
    $points = @()
    for ($d = 0; $d -lt 7; $d++) {
        for ($h = 0; $h -lt 24; $h++) {
            $prevDay  = if ($h -eq 0) { ($d + 6) % 7 } else { $d }
            $prevHour = if ($h -eq 0) { 23 } else { $h - 1 }
            if ($grid[$prevDay, $prevHour] -and -not $grid[$d, $h]) {
                $points += [PSCustomObject]@{ Day = $d; TimeHHmm = ('{0:D2}:01' -f $h) }
            }
        }
    }
    $groups = $points | Group-Object TimeHHmm
    return $groups | ForEach-Object {
        [PSCustomObject]@{ TimeHHmm = $_.Name; Days = @($_.Group.Day | Sort-Object -Unique) }
    }
}

$SchtasksDayAbbrev = @('SUN','MON','TUE','WED','THU','FRI','SAT')
$DaysOfWeekBits = @{ 1 = 0; 2 = 1; 4 = 2; 8 = 3; 16 = 4; 32 = 5; 64 = 6 }  # bit value -> $DayNames index (0=Su)

function Get-AccountLogoutLikeTasks([string]$userName) {
    # Finds ANY scheduled task that locks/logs off this account -- matched by
    # the task's actual run-as user and action content, NOT by name. Real
    # client machines have these set up manually years ago (often just
    # called "Lock PC"), so name-based matching would miss every one of them.
    try { Import-Module ScheduledTasks -ErrorAction Stop } catch { return @() }
    $results = @()
    $broad = '^(S-1-5-32-545|S-1-1-0|S-1-5-11|S-1-5-4)$|(^|\\)(Users|Everyone|Authenticated Users|INTERACTIVE)$'
    foreach ($t in (Get-ScheduledTask)) {
        # A task can run as a single account OR as a GROUP (typically Users): a group task locks whichever
        # standard account is signed in, so it covers every non-admin account.
        $uid = $t.Principal.UserId; $gid = $t.Principal.GroupId
        $scope = $null
        if ($uid -and (($uid -split '\\')[-1] -ieq $userName)) { $scope = 'user' }
        elseif (($uid -and $uid -match $broad) -or ($gid -and $gid -match $broad)) { $scope = 'group' }
        if (-not $scope) { continue }
        $actionText = ($t.Actions | ForEach-Object { "$($_.Execute) $($_.Arguments)" }) -join '; '
        if ($actionText -notmatch 'LockWorkStation' -and $actionText -notmatch '(?i)shutdown.*\s/l\b') { continue }
        foreach ($trig in $t.Triggers) {
            if ($trig.CimClass.CimClassName -ne 'MSFT_TaskWeeklyTrigger') { continue }
            try { $timeHHmm = ([datetime]$trig.StartBoundary).ToString('HH:mm') } catch { continue }
            $days = @()
            foreach ($bit in $DaysOfWeekBits.Keys) { if (([int]$trig.DaysOfWeek -band $bit) -ne 0) { $days += $DaysOfWeekBits[$bit] } }
            $results += [PSCustomObject]@{ TaskName = $t.TaskName; State = $t.State; TimeHHmm = $timeHHmm; Days = $days; Action = $actionText; Scope = $scope }
        }
    }
    return $results
}

function Get-LogoutTaskCoverageLines($grid, $expected, $actualTasks, [string]$userName) {
    # Compares the weekly lock/log-off tasks with what the grid needs.
    # A task does NOT have to match the days exactly: a task that also fires on days when the account is
    # already blocked at that hour is harmless (reported with a note). It is only a PROBLEM when it fires
    # on a day when the grid ALLOWS use at that hour (it would cut the session short), or when a needed
    # day is not covered at all. Tasks within 10 minutes after the top of the lockout hour count as the
    # lockout (the usual one is HH:01).
    $lines = New-Object System.Collections.Generic.List[string]
    $expected = @($expected); $actualTasks = @($actualTasks)
    $hourOf = { param($t) [int]($t.Substring(0, 2)) }
    $minOf  = { param($t) [int]($t.Substring(3, 2)) }
    $dayName = { param($list) (@($list | Sort-Object) | ForEach-Object { $DayNames[$_] }) -join ',' }
    $accounted = @{}
    if ($expected.Count -eq 0) { $lines.Add('(Grid has no lockout points, so no logout tasks are needed.)') }
    foreach ($g in $expected) {
        $h = & $hourOf $g.TimeHHmm
        $same = @($actualTasks | Where-Object { (& $hourOf $_.TimeHHmm) -eq $h -and (& $minOf $_.TimeHHmm) -le 10 })
        foreach ($t in $same) { $accounted[$t.TaskName] = $true }
        $cov = @(); foreach ($t in $same) { $cov += @($t.Days) }; $cov = @($cov | Sort-Object -Unique)
        $missing = @($g.Days | Where-Object { $cov -notcontains $_ })
        $extras  = @($cov | Where-Object { $g.Days -notcontains $_ })
        $harm    = @($extras | Where-Object { $grid[$_, $h] })
        $names = (@($same | ForEach-Object { $_.TaskName } | Sort-Object -Unique) | ForEach-Object { "'$_'" }) -join ', '
        $dayStr = & $dayName $g.Days
        if ($missing.Count -gt 0) {
            if ($same.Count -eq 0) { $lines.Add("[FAIL] $($g.TimeHHmm) ($dayStr) -> no logout task found for '$userName'") }
            else { $lines.Add("[FAIL] $($g.TimeHHmm) ($dayStr) -> $names only cover part of it; missing day(s): $(& $dayName $missing)") }
            continue
        }
        $notes = @()
        if (@($same | Where-Object { $_.Scope -eq 'group' }).Count -gt 0) { $notes += 'runs as a shared GROUP such as Users - covers every standard account' }
        if (@($same | Where-Object { $_.State -eq 'Disabled' }).Count -gt 0) { $notes += 'WARNING: a task is DISABLED' }
        if (@($same | Where-Object { $_.TimeHHmm -ne $g.TimeHHmm }).Count -gt 0) { $notes += "fires at $((@($same | ForEach-Object { $_.TimeHHmm } | Sort-Object -Unique)) -join '/') instead of the usual $($g.TimeHHmm)" }
        if ($harm.Count -gt 0) {
            $lines.Add("[FAIL] $($g.TimeHHmm) ($dayStr) -> covered by $names BUT it also fires on $(& $dayName $harm), when the hours are ALLOWED at that time - it would cut those sessions short. Click 'Sync Logout Tasks To Grid' (makes exact tasks) or fix that task.")
        } else {
            if ($extras.Count -gt 0) { $notes += "also fires on $(& $dayName $extras) - harmless, those days are already blocked at that time" }
            $lines.Add("[PASS] $($g.TimeHHmm) ($dayStr) -> found: $names$(if ($notes.Count) { ' (' + ($notes -join '; ') + ')' })")
        }
    }
    foreach ($t in $actualTasks) {
        if ($accounted.ContainsKey($t.TaskName)) { continue }
        $h = & $hourOf $t.TimeHHmm
        $harm = @(@($t.Days) | Where-Object { $grid[$_, $h] })
        $dayStr = & $dayName $t.Days
        if ($harm.Count -gt 0) {
            $lines.Add("[WARN] Extra logout task '$($t.TaskName)' at $($t.TimeHHmm) ($dayStr) fires on $(& $dayName $harm) when the hours are ALLOWED - it would cut those sessions short (stale from an older schedule?).")
        } else {
            $lines.Add("       note: extra logout task '$($t.TaskName)' at $($t.TimeHHmm) ($dayStr) only fires while the account is already blocked - harmless (can be deleted).")
        }
    }
    return @($lines)
}

function Run-ProcessCaptured([string]$fileName, [string]$arguments) {
    $psi = New-Object System.Diagnostics.ProcessStartInfo
    $psi.FileName = $fileName
    $psi.Arguments = $arguments
    $psi.RedirectStandardOutput = $true
    $psi.RedirectStandardError = $true
    $psi.UseShellExecute = $false
    $psi.CreateNoWindow = $true
    $proc = New-Object System.Diagnostics.Process
    $proc.StartInfo = $psi
    [void]$proc.Start()
    $stdout = $proc.StandardOutput.ReadToEnd()
    $stderr = $proc.StandardError.ReadToEnd()
    $proc.WaitForExit()
    return [PSCustomObject]@{ Success = ($proc.ExitCode -eq 0); ExitCode = $proc.ExitCode; Output = ($stdout + $stderr).Trim() }
}

function Register-WeeklyLogoutTask([string]$userName, [string]$password, [string]$timeHHmm, [int[]]$days, [string]$action) {
    # Runs as the STUDENT ACCOUNT ITSELF -- never SYSTEM: a SYSTEM task cannot
    # lock/log off an interactive user's session at all.
    #
    # Blank $password (the normal case here: these accounts have no password)
    # => interactive-only (/it), which needs no password and fires only while
    # that account is logged on -- which is the only time a lock matters.
    # "Run whether logged on or not" CANNOT work for a no-password account:
    # Windows' default LimitBlankPasswordUse policy bars blank-password
    # accounts from batch logon, and schtasks hangs on a blank /rp.
    #
    # A non-blank $password => "run whether logged on or not" (the mode in
    # our "Setting Student Time Restrictions" setup notes), for accounts
    # that do have a password.
    $dayAbbrevs = ($days | ForEach-Object { $SchtasksDayAbbrev[$_] }) -join ','
    $dayTag = ($days | Sort-Object) -join ''
    $taskName = "StudentTimes-Logout-$userName-$($timeHHmm -replace ':','')-$dayTag"
    $command = if ($action -eq 'LogOff') { 'shutdown.exe /l' } else { 'rundll32.exe user32.dll,LockWorkStation' }
    $credPart = if ($password) { ' /rp "' + $password + '"' } else { ' /it' }
    $args = '/create /tn "' + $taskName + '" /tr "' + $command + '" /sc weekly /d ' + $dayAbbrevs + ' /st ' + $timeHHmm + ' /ru "' + $userName + '"' + $credPart + ' /rl LIMITED /f'
    $result = Run-ProcessCaptured 'schtasks.exe' $args
    return [PSCustomObject]@{ Success = $result.Success; Output = $result.Output; TaskName = $taskName }
}

# ---------------------------------------------------------------------------
# Logout-task sync: THE rule is "whenever hours are set, the weekly logout tasks are
# made to match". This one self-contained scriptblock does it (no helper functions
# needed) so the very same code can run locally, on a remote PC during a Network Push,
# and inside a future-dated scheduled change. $groups = @{TimeHHmm='HH:mm'; Days=@(0..6)}.
# It keeps tasks that already match, removes this account's lock/logoff tasks that don't
# (stale ones, e.g. an old "Lock"), and creates whatever is missing. Every result line
# starts with [OK] or [FAIL].
# ---------------------------------------------------------------------------
$SyncLogoutTasksBlock = {
    param([string]$userName, $groups, [string]$action, [string]$password)
    $out = New-Object System.Collections.ArrayList
    $abbr = @('SUN','MON','TUE','WED','THU','FRI','SAT'); $dn = @('Su','M','T','W','Th','F','Sa')
    $bits = @{ 1 = 0; 2 = 1; 4 = 2; 8 = 3; 16 = 4; 32 = 5; 64 = 6 }
    $groups = @($groups | Where-Object { $_ })
    try { Import-Module ScheduledTasks -ErrorAction Stop } catch { return @("[FAIL] Logout tasks: Task Scheduler module not available: $($_.Exception.Message)") }
    $satisfied = @{}
    $broad = '^(S-1-5-32-545|S-1-1-0|S-1-5-11|S-1-5-4)$|(^|\\)(Users|Everyone|Authenticated Users|INTERACTIVE)$'
    foreach ($t in @(Get-ScheduledTask -ErrorAction SilentlyContinue)) {
        $uid = $t.Principal.UserId; $gid = $t.Principal.GroupId
        $scope = $null
        if ($uid -and ((($uid -split '\\')[-1]) -ieq $userName)) { $scope = 'user' }
        elseif (($uid -and $uid -match $broad) -or ($gid -and $gid -match $broad)) { $scope = 'group' }
        if (-not $scope) { continue }
        $actionText = (@($t.Actions) | ForEach-Object { "$($_.Execute) $($_.Arguments)" }) -join '; '
        if ($actionText -notmatch 'LockWorkStation' -and $actionText -notmatch '(?i)shutdown.*\s/l\b') { continue }
        $trigs = @($t.Triggers); $hhmm = ''; $days = @(); $keep = $false
        if ($trigs.Count -eq 1 -and $trigs[0].CimClass.CimClassName -eq 'MSFT_TaskWeeklyTrigger') {
            try { $hhmm = ([datetime]$trigs[0].StartBoundary).ToString('HH:mm') } catch { $hhmm = '' }
            foreach ($b in $bits.Keys) { if (([int]$trigs[0].DaysOfWeek -band $b) -ne 0) { $days += $bits[$b] } }
            $days = @($days | Sort-Object)
            foreach ($g in $groups) {
                $gd = @($g.Days | Sort-Object); $key = "$($g.TimeHHmm)|$($gd -join ',')"
                if ($g.TimeHHmm -eq $hhmm -and ($gd -join ',') -eq ($days -join ',') -and -not $satisfied.ContainsKey($key)) { $keep = $true; $satisfied[$key] = $true; break }
            }
        }
        $desc = if ($hhmm) { "$hhmm on $((@($days) | ForEach-Object { $dn[$_] }) -join ',')" } else { 'other trigger' }
        if ($scope -eq 'group') {
            # shared task (runs as the Users group): it covers this account if it matches; never delete or edit it here
            if ($keep) { [void]$out.Add("[OK] Shared group task '$($t.TaskName)' ($desc) already covers '$userName' - kept, no per-account copy made.") }
            continue
        }
        if ($keep) {
            if ($t.State -eq 'Disabled') {
                try { Enable-ScheduledTask -TaskName $t.TaskName -TaskPath $t.TaskPath | Out-Null; [void]$out.Add("[OK] Logout task '$($t.TaskName)' ($desc) was disabled - re-enabled.") }
                catch { [void]$out.Add("[FAIL] Logout task '$($t.TaskName)' is disabled and could not be enabled: $($_.Exception.Message)") }
            } else { [void]$out.Add("[OK] Logout task '$($t.TaskName)' ($desc) already right - kept.") }
        } else {
            try { Unregister-ScheduledTask -TaskName $t.TaskName -TaskPath $t.TaskPath -Confirm:$false -ErrorAction Stop; [void]$out.Add("[OK] Removed old logout task '$($t.TaskName)' ($desc) - no longer matches the hours.") }
            catch { [void]$out.Add("[FAIL] Could not remove old logout task '$($t.TaskName)' ($desc): $($_.Exception.Message)") }
        }
    }
    foreach ($g in $groups) {
        $gd = @($g.Days | Sort-Object); $key = "$($g.TimeHHmm)|$($gd -join ',')"
        if ($satisfied.ContainsKey($key)) { continue }
        $dayAbbrevs = ($gd | ForEach-Object { $abbr[$_] }) -join ','
        $dayTag = $gd -join ''
        $taskName = "StudentTimes-Logout-$userName-$($g.TimeHHmm -replace ':','')-$dayTag"
        $command = if ($action -eq 'LogOff') { 'shutdown.exe /l' } else { 'rundll32.exe user32.dll,LockWorkStation' }
        $cred = if ($password) { @('/rp', $password) } else { @('/it') }
        $argv = @('/create', '/tn', $taskName, '/tr', $command, '/sc', 'weekly', '/d', $dayAbbrevs, '/st', $g.TimeHHmm, '/ru', $userName) + $cred + @('/rl', 'LIMITED', '/f')
        $res = (& schtasks.exe @argv 2>&1 | ForEach-Object { if ($_ -is [System.Management.Automation.ErrorRecord]) { $_.Exception.Message } else { "$_" } } | Where-Object { $_ -and $_.Trim() }) -join ' '
        $dstr = ($gd | ForEach-Object { $dn[$_] }) -join ','
        if ($LASTEXITCODE -eq 0) { [void]$out.Add("[OK] Created logout task $($g.TimeHHmm) on $dstr ($taskName).") }
        else { [void]$out.Add("[FAIL] Could not create logout task $($g.TimeHHmm) on ${dstr}: $res") }
    }
    if ($groups.Count -eq 0) { [void]$out.Add('[OK] The hours have no lockout points (fully open or fully blocked) - no logout tasks are needed.') }
    return @($out)
}

function Sync-LogoutTasks([string]$userName, $grid) {
    # Local run of the shared block, using the action / password chosen on the Setup tab.
    $groups = @(Get-LogoutKickoffGroups $grid)
    $action = if ($cmbLogoutAction.SelectedItem -eq 'Full log off') { 'LogOff' } else { 'Lock' }
    return @(& $SyncLogoutTasksBlock $userName $groups $action $txtLogoutPassword.Text)
}

# ---------------------------------------------------------------------------
# Fixed lock-screen / sign-in picture for the WHOLE PC (the picture you see before you pick a user),
# so nobody can ever put anything inappropriate there.  Self-contained scriptblock (runs locally and
# remotely inside Network Push).   $mode = 'Check' | 'Apply' | 'Undo'
#   $opts = @{ ImageBytes = [byte[]]; ImageExt = '.jpg' | '.png' }   (only needed for Apply)
# The picture is copied to C:\ProgramData\StudentTimes\LockScreen (Users: read-only), then pointed to by
# the lock-screen policy + PersonalizationCSP values, and changing it is blocked (NoChangingLockScreen).
# ---------------------------------------------------------------------------
$LockScreenPictureBlock = {
    param([string]$mode, $opts)
    $out = New-Object System.Collections.ArrayList
    $dir = Join-Path $env:ProgramData 'StudentTimes\LockScreen'
    $pol = 'HKLM:\SOFTWARE\Policies\Microsoft\Windows\Personalization'
    $csp = 'HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\PersonalizationCSP'
    $sys = 'HKLM:\SOFTWARE\Policies\Microsoft\Windows\System'
    $cloud = 'HKLM:\SOFTWARE\Policies\Microsoft\Windows\CloudContent'
    function Set-RegValue($path, $name, $value, $type) {
        if (-not (Test-Path -LiteralPath $path)) { New-Item -Path $path -Force -ErrorAction Stop | Out-Null }
        New-ItemProperty -LiteralPath $path -Name $name -Value $value -PropertyType $type -Force -ErrorAction Stop | Out-Null
    }
    function Get-RegValue($path, $name) {
        if (-not (Test-Path -LiteralPath $path)) { return $null }
        return (Get-ItemProperty -LiteralPath $path -Name $name -ErrorAction SilentlyContinue).$name
    }
    function Remove-RegValue($path, $name) {
        if (Test-Path -LiteralPath $path) { try { Remove-ItemProperty -LiteralPath $path -Name $name -ErrorAction Stop } catch { if ($_.Exception.Message -notmatch 'does not exist') { throw } } }
    }
    if ($mode -ne 'Check') {
        $isAdm = ([Security.Principal.WindowsPrincipal][Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)
        if (-not $isAdm) { return @('[FAIL] This needs administrator rights (run the tool as administrator). Nothing was changed.') }
    }
    try {
        if ($mode -eq 'Apply') {
            $imgBytes = $null; $imgExt = ''
            if ($opts -and $opts.WindowsDefault) {
                # Windows' own default lock-screen picture, taken from THIS machine (so it also works in a push with no big upload)
                $webDir = Join-Path $env:windir 'Web\Screen'
                $srcFile = Join-Path $webDir 'img100.jpg'
                if (-not (Test-Path -LiteralPath $srcFile)) { $srcFile = (Get-ChildItem -Path $webDir -Filter 'img10*.jpg' -File -ErrorAction SilentlyContinue | Sort-Object Length -Descending | Select-Object -First 1).FullName }
                if (-not $srcFile) { return @("[FAIL] Windows' default lock-screen pictures were not found in $webDir on this PC.") }
                $imgBytes = [System.IO.File]::ReadAllBytes($srcFile); $imgExt = '.jpg'
            } elseif ($opts -and $opts.ImageBytes) { $imgBytes = [byte[]]$opts.ImageBytes; $imgExt = $opts.ImageExt }
            if (-not $imgBytes -or @($imgBytes).Count -lt 1000) { return @('[FAIL] No picture was chosen (or the file is empty).') }
            if (-not (Test-Path $dir)) { New-Item -ItemType Directory -Path $dir -Force -ErrorAction Stop | Out-Null }
            Get-ChildItem -Path $dir -Filter 'LockScreen.*' -File -ErrorAction SilentlyContinue | Remove-Item -Force -ErrorAction SilentlyContinue
            $file = Join-Path $dir ('LockScreen' + $imgExt)
            [System.IO.File]::WriteAllBytes($file, $imgBytes)
            & icacls.exe $dir /inheritance:r /grant:r '*S-1-5-18:(OI)(CI)F' '*S-1-5-32-544:(OI)(CI)F' '*S-1-5-32-545:(OI)(CI)RX' /Q 2>&1 | Out-Null
            if ($LASTEXITCODE -ne 0) { [void]$out.Add('[WARN] Could not make the picture folder read-only for ordinary users.') }
            Set-RegValue $pol 'LockScreenImage' $file 'String'
            Set-RegValue $pol 'NoChangingLockScreen' 1 'DWord'
            Set-RegValue $csp 'LockScreenImagePath' $file 'String'
            Set-RegValue $csp 'LockScreenImageUrl' $file 'String'
            Set-RegValue $csp 'LockScreenImageStatus' 1 'DWord'
            Set-RegValue $sys 'DisableLogonBackgroundImage' 0 'DWord'
            Set-RegValue $cloud 'DisableWindowsSpotlightFeatures' 1 'DWord'
            [void]$out.Add("[OK] Fixed lock-screen / sign-in picture set: $file ($(@($imgBytes).Count) bytes$(if ($opts.WindowsDefault) { ', the Windows default picture' })). Changing it is blocked.")
            [void]$out.Add('Takes effect at the next lock / sign-out / restart. Check it by signing out (not just locking).')
        } elseif ($mode -eq 'Undo') {
            foreach ($n in 'LockScreenImage', 'NoChangingLockScreen') { Remove-RegValue $pol $n }
            foreach ($n in 'LockScreenImagePath', 'LockScreenImageUrl', 'LockScreenImageStatus') { Remove-RegValue $csp $n }
            Remove-RegValue $sys 'DisableLogonBackgroundImage'
            Remove-RegValue $cloud 'DisableWindowsSpotlightFeatures'
            if (Test-Path $dir) { Remove-Item -Path $dir -Recurse -Force -ErrorAction SilentlyContinue }
            [void]$out.Add('[OK] The fixed lock-screen picture was removed; Windows is back to its normal lock screen.')
        } else {
            $img = Get-RegValue $pol 'LockScreenImage'
            $block = Get-RegValue $pol 'NoChangingLockScreen'
            $cspPath = Get-RegValue $csp 'LockScreenImagePath'
            if ($img -and (Test-Path -LiteralPath $img)) { [void]$out.Add("[PASS] A fixed picture is set and the file exists: $img ($((Get-Item -LiteralPath $img).Length) bytes).") }
            elseif ($img) { [void]$out.Add("[FAIL] A fixed picture is set ($img) but the file is MISSING.") }
            else { [void]$out.Add('[FAIL] No fixed lock-screen picture is set (Windows shows its normal / Spotlight picture).') }
            if ($block -eq 1) { [void]$out.Add('[PASS] Changing the lock screen is blocked.') } else { [void]$out.Add('[FAIL] Changing the lock screen is NOT blocked.') }
            if ($cspPath -and $img -and $cspPath -eq $img) { [void]$out.Add('[PASS] The second (PersonalizationCSP) setting points at the same picture.') } elseif ($img) { [void]$out.Add('[WARN] The PersonalizationCSP setting is missing or different (Windows Pro may need it).') }
        }
    } catch {
        [void]$out.Add("[FAIL] $mode failed: $($_.Exception.Message)")
    }
    return @($out)
}

function New-PlainLockPicture {
    # A plain, safe default picture: dark blue with the company name. Returns the saved file path.
    Add-Type -AssemblyName System.Drawing
    $dir = Join-Path $env:LOCALAPPDATA 'StudentTimes'
    if (-not (Test-Path $dir)) { New-Item -ItemType Directory -Path $dir -Force | Out-Null }
    $path = Join-Path $dir 'LockScreen-plain.png'
    $bmp = New-Object System.Drawing.Bitmap 1920, 1080
    $g = [System.Drawing.Graphics]::FromImage($bmp)
    $g.SmoothingMode = 'AntiAlias'; $g.TextRenderingHint = 'AntiAlias'
    $g.Clear([System.Drawing.Color]::FromArgb(18, 40, 76))
    $fmt = New-Object System.Drawing.StringFormat; $fmt.Alignment = 'Center'; $fmt.LineAlignment = 'Center'
    $font = New-Object System.Drawing.Font('Segoe UI', 54, [System.Drawing.FontStyle]::Bold)
    $g.DrawString('Rubin IT Computers and Telecom', $font, [System.Drawing.Brushes]::White, (New-Object System.Drawing.RectangleF(0, 0, 1920, 1080)), $fmt)
    $g.Dispose(); $bmp.Save($path, [System.Drawing.Imaging.ImageFormat]::Png); $bmp.Dispose()
    return $path
}

function Get-LockPictureOptions([string]$file) {
    if ($file -like 'Windows default*') { return @{ WindowsDefault = $true } }
    if (-not $file -or -not (Test-Path -LiteralPath $file)) { return $null }
    $ext = ([System.IO.Path]::GetExtension($file)).ToLower()
    if ($ext -notin @('.jpg', '.jpeg', '.png', '.bmp')) { return $null }
    if ($ext -eq '.jpeg') { $ext = '.jpg' }
    return @{ ImageBytes = [System.IO.File]::ReadAllBytes($file); ImageExt = $ext }
}

# ---------------------------------------------------------------------------
# Desktop lock-down: stop a (student) account changing the desktop background / theme / colors and turn
# off Windows Spotlight (rotating lock-screen + desktop pictures, "Learn about this picture", tips,
# suggestions and auto-installed promo apps).  Self-contained scriptblock (same pattern as
# $SyncLogoutTasksBlock) so it runs locally, on a remote PC inside Network Push, anywhere.
#   $mode = 'Check' | 'Apply' | 'Undo'      $opts = @{ Wallpaper = $true; Spotlight = $true; LockScreenMachine = $false }
# Works on the account's registry hive even when it is not signed in (loads NTUSER.DAT, unloads after).
# Every result line starts with [OK] / [PASS] / [FAIL] / [WARN].
# ---------------------------------------------------------------------------
$DesktopLockBlock = {
    param([string]$userName, [string]$mode, $opts)
    $out = New-Object System.Collections.ArrayList
    try { $sid = (New-Object System.Security.Principal.NTAccount($userName)).Translate([System.Security.Principal.SecurityIdentifier]).Value }
    catch { return @("[FAIL] There is no account named '$userName' on this PC.") }

    $loadedHere = $false; $tmp = 'StudentTimesLockTmp'
    $root = "Registry::HKEY_USERS\$sid"
    if (-not (Test-Path $root)) {
        $prof = (Get-ItemProperty "HKLM:\SOFTWARE\Microsoft\Windows NT\CurrentVersion\ProfileList\$sid" -ErrorAction SilentlyContinue).ProfileImagePath
        if (-not $prof -or -not (Test-Path (Join-Path $prof 'NTUSER.DAT'))) {
            return @("[FAIL] '$userName' has never signed in on this PC, so it has no profile to lock down yet. Sign in once as that account, then run this again.")
        }
        $r = (& reg.exe load "HKU\$tmp" (Join-Path $prof 'NTUSER.DAT') 2>&1 | ForEach-Object { "$_" }) -join ' '
        if ($LASTEXITCODE -ne 0) { return @("[FAIL] Could not open the profile of '$userName': $r") }
        $loadedHere = $true; $root = "Registry::HKEY_USERS\$tmp"
    }

    $cdm  = 'Software\Microsoft\Windows\CurrentVersion\ContentDeliveryManager'
    $cloud = 'Software\Policies\Microsoft\Windows\CloudContent'
    $items = New-Object System.Collections.ArrayList
    function Add-Item($group, $scope, $path, $name, $value) { [void]$items.Add([pscustomobject]@{ Group = $group; Scope = $scope; Path = $path; Name = $name; Value = $value }) }
    if ($opts.Wallpaper) {
        $g = 'Block changing the desktop background / theme / colors'
        Add-Item $g 'U' 'Software\Microsoft\Windows\CurrentVersion\Policies\ActiveDesktop' 'NoChangingWallPaper' 1
        Add-Item $g 'U' 'Software\Microsoft\Windows\CurrentVersion\Policies\Explorer' 'NoThemesTab' 1
        Add-Item $g 'U' 'Software\Microsoft\Windows\CurrentVersion\Policies\System' 'NoDispBackgroundPage' 1
        Add-Item $g 'U' 'Software\Microsoft\Windows\CurrentVersion\Policies\System' 'NoDispAppearancePage' 1
    }
    if ($opts.Spotlight) {
        $g = 'Turn off Windows Spotlight, tips and suggestions'
        foreach ($n in 'DisableWindowsSpotlightFeatures','DisableWindowsSpotlightOnActionCenter','DisableWindowsSpotlightOnSettings','DisableWindowsSpotlightWindowsWelcomeExperience','DisableThirdPartySuggestions','DisableTailoredExperiencesWithDiagnosticData') { Add-Item $g 'U' $cloud $n 1 }
        foreach ($n in 'RotatingLockScreenEnabled','RotatingLockScreenOverlayEnabled','SubscribedContent-338387Enabled','SubscribedContent-338388Enabled','SubscribedContent-338389Enabled','SubscribedContent-353694Enabled','SubscribedContent-353696Enabled','SoftLandingEnabled','SystemPaneSuggestionsEnabled','ContentDeliveryAllowed','OemPreInstalledAppsEnabled','PreInstalledAppsEnabled','SilentInstalledAppsEnabled') { Add-Item $g 'U' $cdm $n 0 }
        Add-Item $g 'U' 'Software\Microsoft\Windows\CurrentVersion\Explorer\HideDesktopIcons\NewStartPanel' '{2cc5ca98-6485-489a-920e-b3e88a6ccce3}' 1
        Add-Item $g 'M' 'SOFTWARE\Policies\Microsoft\Windows\CloudContent' 'DisableWindowsConsumerFeatures' 1
    }
    if ($opts.LockScreenMachine) {
        Add-Item 'Block changing the lock-screen picture (whole PC)' 'M' 'SOFTWARE\Policies\Microsoft\Windows\Personalization' 'NoChangingLockScreen' 1
    }
    if ($items.Count -eq 0) { return @('[WARN] Nothing selected.') }

    try {
        foreach ($grp in ($items | Group-Object Group)) {
            $good = 0; $bad = 0; $errs = @()
            foreach ($it in $grp.Group) {
                $base = if ($it.Scope -eq 'M') { "HKLM:\$($it.Path)" } else { "$root\$($it.Path)" }
                try {
                    if ($mode -eq 'Apply') {
                        if (-not (Test-Path -LiteralPath $base)) { New-Item -Path $base -Force -ErrorAction Stop | Out-Null }
                        New-ItemProperty -LiteralPath $base -Name $it.Name -Value $it.Value -PropertyType DWord -Force -ErrorAction Stop | Out-Null
                        $good++
                    } elseif ($mode -eq 'Undo') {
                        if (Test-Path -LiteralPath $base) { try { Remove-ItemProperty -LiteralPath $base -Name $it.Name -ErrorAction Stop } catch { if ($_.Exception.Message -notmatch 'does not exist') { throw } } }
                        $good++
                    } else {
                        $cur = $null
                        if (Test-Path -LiteralPath $base) { $cur = (Get-ItemProperty -LiteralPath $base -Name $it.Name -ErrorAction SilentlyContinue).($it.Name) }
                        if ($null -ne $cur -and [int]$cur -eq [int]$it.Value) { $good++ } else { $bad++ }
                    }
                } catch { $bad++; $errs += $_.Exception.Message }
            }
            $total = $grp.Count
            if ($mode -eq 'Check') {
                if ($bad -eq 0) { [void]$out.Add("[PASS] $($grp.Name) - all $total settings are in place for '$userName'.") }
                else { [void]$out.Add("[FAIL] $($grp.Name) - only $good of $total settings are in place for '$userName' (run Apply).") }
            } elseif ($bad -eq 0) {
                [void]$out.Add("[OK] $($grp.Name) - $(if ($mode -eq 'Apply') { 'applied' } else { 'removed' }) ($total settings) for '$userName'.")
            } else {
                [void]$out.Add("[FAIL] $($grp.Name) - $bad of $total settings failed: $(($errs | Select-Object -Unique) -join '; ')")
            }
        }
        # Win11 "Windows Spotlight" desktop background (BackgroundType 3) -> plain picture
        if ($opts.Spotlight -and $mode -ne 'Undo') {
            $wp = "$root\Software\Microsoft\Windows\CurrentVersion\Explorer\Wallpapers"
            $bt = if (Test-Path -LiteralPath $wp) { (Get-ItemProperty -LiteralPath $wp -Name BackgroundType -ErrorAction SilentlyContinue).BackgroundType } else { $null }
            if ($bt -eq 3) {
                if ($mode -eq 'Apply') { Set-ItemProperty -LiteralPath $wp -Name BackgroundType -Value 0 -ErrorAction Stop; [void]$out.Add("[OK] The desktop background was set to 'Windows Spotlight' - switched to a normal picture.") }
                else { [void]$out.Add("[FAIL] The desktop background is still set to 'Windows Spotlight' (run Apply).") }
            }
        }
    } finally {
        if ($loadedHere) {
            [gc]::Collect(); [gc]::WaitForPendingFinalizers(); Start-Sleep -Milliseconds 400
            & reg.exe unload "HKU\$tmp" 2>&1 | Out-Null
            if ($LASTEXITCODE -ne 0) { Start-Sleep -Seconds 2; [gc]::Collect(); & reg.exe unload "HKU\$tmp" 2>&1 | Out-Null }
            if ($LASTEXITCODE -ne 0) { [void]$out.Add("[WARN] Could not unload the profile hive of '$userName' (it will unload by itself at the next restart; the settings were saved).") }
        }
    }
    if ($mode -eq 'Apply' -and -not @($out | Where-Object { "$_" -match '^\[FAIL\]' }).Count) { [void]$out.Add("Takes effect at the next sign-in of '$userName' (or sign them out now).") }
    return @($out)
}

$setupTop = 20

$lblSetup = New-Object System.Windows.Forms.Label
$lblSetup.Text = 'Initial setup (one time on each PC):'
$lblSetup.Location = New-Object System.Drawing.Point(15, $setupTop)
$lblSetup.Size = New-Object System.Drawing.Size(1290, 18)
$setupPanel.Controls.Add($lblSetup)

$lblSetupWarning = New-Object System.Windows.Forms.Label
$lblSetupWarning.Text = 'The weekly logout tasks now follow the hours automatically: applying hours (Run it immediately), Unlock/Restore, slots, future-dated changes and Network Push all update them too. This button re-syncs them by hand (e.g. after pasting a copied command).'
$lblSetupWarning.Location = New-Object System.Drawing.Point(15, ($setupTop + 18))
$lblSetupWarning.Size = New-Object System.Drawing.Size(1290, 18)
$lblSetupWarning.ForeColor = [System.Drawing.Color]::DarkRed
$setupPanel.Controls.Add($lblSetupWarning)

$lblLogoutAction = New-Object System.Windows.Forms.Label
$lblLogoutAction.Text = 'When kicked off:'
$lblLogoutAction.Location = New-Object System.Drawing.Point(15, ($setupTop + 60))
$lblLogoutAction.AutoSize = $true
$setupPanel.Controls.Add($lblLogoutAction)

$cmbLogoutAction = New-Object System.Windows.Forms.ComboBox
$cmbLogoutAction.DropDownStyle = 'DropDownList'
$cmbLogoutAction.Items.AddRange(@('Lock workstation', 'Full log off'))
$cmbLogoutAction.SelectedIndex = 0
$cmbLogoutAction.Location = New-Object System.Drawing.Point(130, ($setupTop + 56))
$cmbLogoutAction.Size = New-Object System.Drawing.Size(140, 24)
$setupPanel.Controls.Add($cmbLogoutAction)

$lblLogoutPassword = New-Object System.Windows.Forms.Label
$lblLogoutPassword.Text = "Password (blank = none):"
$lblLogoutPassword.Location = New-Object System.Drawing.Point(280, ($setupTop + 60))
$lblLogoutPassword.AutoSize = $true
$setupPanel.Controls.Add($lblLogoutPassword)

$txtLogoutPassword = New-Object System.Windows.Forms.TextBox
$txtLogoutPassword.UseSystemPasswordChar = $true
$txtLogoutPassword.Location = New-Object System.Drawing.Point(430, ($setupTop + 56))
$txtLogoutPassword.Size = New-Object System.Drawing.Size(125, 24)
$setupPanel.Controls.Add($txtLogoutPassword)

$btnCreateLogoutTasks = New-Object System.Windows.Forms.Button
$btnCreateLogoutTasks.Text = '2. Sync Logout Tasks To Grid'
$btnCreateLogoutTasks.Location = New-Object System.Drawing.Point(565, ($setupTop + 54))
$btnCreateLogoutTasks.Size = New-Object System.Drawing.Size(320, 30)
$setupPanel.Controls.Add($btnCreateLogoutTasks)

$btnCheckSecpol = New-Object System.Windows.Forms.Button
$btnCheckSecpol.Text = '3+4. Check/Remove Time-Change Rights'
$btnCheckSecpol.Location = New-Object System.Drawing.Point(15, ($setupTop + 92))
$btnCheckSecpol.Size = New-Object System.Drawing.Size(260, 30)
$setupPanel.Controls.Add($btnCheckSecpol)

$btnVerifyAll = New-Object System.Windows.Forms.Button
$btnVerifyAll.Text = 'Verify Everything For This Account'
$btnVerifyAll.Location = New-Object System.Drawing.Point(285, ($setupTop + 92))
$btnVerifyAll.Size = New-Object System.Drawing.Size(230, 30)
$btnVerifyAll.BackColor = [System.Drawing.Color]::LightSteelBlue
$setupPanel.Controls.Add($btnVerifyAll)

# --- Time settings row: is the clock trustworthy, and can the student move it? ---
$lblTimeRow = New-Object System.Windows.Forms.Label
$lblTimeRow.Text = 'Time settings:'
$lblTimeRow.Location = New-Object System.Drawing.Point(15, ($setupTop + 136))
$lblTimeRow.AutoSize = $true
$setupPanel.Controls.Add($lblTimeRow)

$lblAutoZone = New-Object System.Windows.Forms.Label
$lblAutoZone.Text = 'Automatic time zone:'
$lblAutoZone.Location = New-Object System.Drawing.Point(110, ($setupTop + 136))
$lblAutoZone.AutoSize = $true
$setupPanel.Controls.Add($lblAutoZone)

$cmbAutoZone = New-Object System.Windows.Forms.ComboBox
$cmbAutoZone.DropDownStyle = 'DropDownList'
$cmbAutoZone.Items.AddRange(@('Fixed - turn OFF (recommended)', 'Automatic - turn ON', 'Leave as is'))
$cmbAutoZone.SelectedIndex = 0
$cmbAutoZone.Location = New-Object System.Drawing.Point(235, ($setupTop + 132))
$cmbAutoZone.Size = New-Object System.Drawing.Size(215, 24)
$setupPanel.Controls.Add($cmbAutoZone)

$chkRemoveBroad = New-Object System.Windows.Forms.CheckBox
$chkRemoveBroad.Text = 'Also remove Users/Everyone from the time + time-zone rights (affects ALL standard accounts on this PC)'
$chkRemoveBroad.Checked = $true
$chkRemoveBroad.Location = New-Object System.Drawing.Point(465, ($setupTop + 133))
$chkRemoveBroad.AutoSize = $true
$setupPanel.Controls.Add($chkRemoveBroad)

$btnCheckTime = New-Object System.Windows.Forms.Button
$btnCheckTime.Text = 'Check Time Settings'
$btnCheckTime.Location = New-Object System.Drawing.Point(15, ($setupTop + 162))
$btnCheckTime.Size = New-Object System.Drawing.Size(180, 30)
$setupPanel.Controls.Add($btnCheckTime)

$btnFixTime = New-Object System.Windows.Forms.Button
$btnFixTime.Text = 'Fix Time Settings'
$btnFixTime.Location = New-Object System.Drawing.Point(205, ($setupTop + 162))
$btnFixTime.Size = New-Object System.Drawing.Size(180, 30)
$btnFixTime.BackColor = [System.Drawing.Color]::Khaki
$setupPanel.Controls.Add($btnFixTime)

$lblNtp = New-Object System.Windows.Forms.Label
$lblNtp.Text = 'Time servers (backups if one is down; blank = keep what is there):'
$lblNtp.Location = New-Object System.Drawing.Point(400, ($setupTop + 160))
$lblNtp.AutoSize = $true
$setupPanel.Controls.Add($lblNtp)

$txtNtpServers = New-Object System.Windows.Forms.TextBox
$txtNtpServers.Location = New-Object System.Drawing.Point(400, ($setupTop + 178))
$txtNtpServers.Size = New-Object System.Drawing.Size(905, 24)
$txtNtpServers.Font = New-Object System.Drawing.Font('Consolas', 9)
$setupPanel.Controls.Add($txtNtpServers)   # filled with the recommended list once that is defined, below

$lblAutoNew = New-Object System.Windows.Forms.Label
$lblAutoNew.Text = 'Accounts created LATER (user5, user6 ...): keep EVERY non-admin account on the grid shown on the Daily Use tab, automatically:'
$lblAutoNew.Location = New-Object System.Drawing.Point(15, ($setupTop + 206))
$lblAutoNew.Size = New-Object System.Drawing.Size(1290, 18)
$lblAutoNew.ForeColor = [System.Drawing.Color]::DarkSlateBlue
$setupPanel.Controls.Add($lblAutoNew)

$btnAutoOn = New-Object System.Windows.Forms.Button
$btnAutoOn.Text = 'Auto-Lock New Accounts: Turn ON (this grid)'
$btnAutoOn.Location = New-Object System.Drawing.Point(15, ($setupTop + 230))
$btnAutoOn.Size = New-Object System.Drawing.Size(330, 30)
$btnAutoOn.BackColor = [System.Drawing.Color]::Khaki
$setupPanel.Controls.Add($btnAutoOn)

$btnAutoOff = New-Object System.Windows.Forms.Button
$btnAutoOff.Text = 'Turn OFF'
$btnAutoOff.Location = New-Object System.Drawing.Point(355, ($setupTop + 230))
$btnAutoOff.Size = New-Object System.Drawing.Size(110, 30)
$setupPanel.Controls.Add($btnAutoOff)

$btnAutoStatus = New-Object System.Windows.Forms.Button
$btnAutoStatus.Text = 'Status'
$btnAutoStatus.Location = New-Object System.Drawing.Point(475, ($setupTop + 230))
$btnAutoStatus.Size = New-Object System.Drawing.Size(110, 30)
$setupPanel.Controls.Add($btnAutoStatus)

$lblAutoExclude = New-Object System.Windows.Forms.Label
$lblAutoExclude.Text = 'Never touch these accounts (names, comma separated):'
$lblAutoExclude.Location = New-Object System.Drawing.Point(600, ($setupTop + 237))
$lblAutoExclude.AutoSize = $true
$setupPanel.Controls.Add($lblAutoExclude)

$txtAutoExclude = New-Object System.Windows.Forms.TextBox
$txtAutoExclude.Location = New-Object System.Drawing.Point(930, ($setupTop + 233))
$txtAutoExclude.Size = New-Object System.Drawing.Size(375, 24)
$setupPanel.Controls.Add($txtAutoExclude)

$lblSetupReport = New-Object System.Windows.Forms.Label
$lblSetupReport.Text = 'Report (from the buttons above):'
$lblSetupReport.Location = New-Object System.Drawing.Point(15, ($setupTop + 272))
$lblSetupReport.AutoSize = $true
$lblSetupReport.ForeColor = [System.Drawing.Color]::DimGray
$setupPanel.Controls.Add($lblSetupReport)

function Update-StatusBanner($banner, [string]$text) {
    # Big red / amber / green bar above a report, worked out from the report's own [FAIL]/[WARN]/[PASS] lines.
    $fails = ([regex]::Matches($text, '(?m)^\[(FAIL|FAILED|NOT READY)\]')).Count
    $warns = ([regex]::Matches($text, '(?m)^\[WARN\]')).Count
    $goods = ([regex]::Matches($text, '(?m)^\[(PASS|OK|READY ?)\]')).Count
    if ($fails -gt 0) {
        $banner.Text = "  PROBLEM: $fails thing(s) FAILED - see the [FAIL] lines below" + $(if ($warns) { " (+ $warns warning(s))" } else { '' })
        $banner.BackColor = [System.Drawing.Color]::Firebrick; $banner.ForeColor = [System.Drawing.Color]::White; $banner.Visible = $true
    } elseif ($warns -gt 0) {
        $banner.Text = "  Works, but $warns warning(s) - see the [WARN] lines below"
        $banner.BackColor = [System.Drawing.Color]::DarkOrange; $banner.ForeColor = [System.Drawing.Color]::Black; $banner.Visible = $true
    } elseif ($goods -gt 0) {
        $banner.Text = "  ALL GOOD - nothing failed"
        $banner.BackColor = [System.Drawing.Color]::SeaGreen; $banner.ForeColor = [System.Drawing.Color]::White; $banner.Visible = $true
    } else { $banner.Visible = $false }
}

$txtSetupReport = New-Object System.Windows.Forms.TextBox
$txtSetupReport.Text = '(Nothing yet - click "Create Weekly Logout Tasks", "Check/Remove Time-Change Rights", "Check Time Settings", or "Verify Everything" above to see a report here.)'
$txtSetupReport.Location = New-Object System.Drawing.Point(15, ($setupTop + 334))
$txtSetupReport.Size = New-Object System.Drawing.Size(1290, 170)
$txtSetupReport.Multiline = $true
$txtSetupReport.ReadOnly = $true
$txtSetupReport.BackColor = [System.Drawing.Color]::White
$txtSetupReport.ScrollBars = 'Vertical'
$txtSetupReport.Font = New-Object System.Drawing.Font('Consolas', 11)
$setupPanel.Controls.Add($txtSetupReport)

$lblSetupBanner = New-Object System.Windows.Forms.Label
$lblSetupBanner.Location = New-Object System.Drawing.Point(15, ($setupTop + 292))
$lblSetupBanner.Size = New-Object System.Drawing.Size(1290, 36)
$lblSetupBanner.Font = New-Object System.Drawing.Font('Segoe UI', 15, [System.Drawing.FontStyle]::Bold)
$lblSetupBanner.TextAlign = 'MiddleLeft'
$lblSetupBanner.Visible = $false
$setupPanel.Controls.Add($lblSetupBanner)
$txtSetupReport.Add_TextChanged({ Update-StatusBanner $lblSetupBanner $txtSetupReport.Text })

function New-CopyReportButton($panel, $reportBox, [int]$x, [int]$y) {
    # One click copies the whole report text to the clipboard (no scrolling / selecting).
    $b = New-Object System.Windows.Forms.Button
    $b.Text = 'Copy Report'
    $b.Location = New-Object System.Drawing.Point($x, $y)
    $b.Size = New-Object System.Drawing.Size(130, 26)
    $b.BackColor = [System.Drawing.Color]::LightSteelBlue
    $panel.Controls.Add($b)
    $b.Tag = $reportBox
    $b.Add_Click({
        $box = $this.Tag
        if ($box.Text) {
            [System.Windows.Forms.Clipboard]::SetText($box.Text)
            $this.Text = 'Copied!'
            $timer = New-Object System.Windows.Forms.Timer
            $timer.Interval = 1500
            $btn = $this
            $timer.Add_Tick({ $btn.Text = 'Copy Report'; $timer.Stop(); $timer.Dispose() }.GetNewClosure())
            $timer.Start()
        }
    })
    return $b
}

$btnCopySetupReport = New-CopyReportButton $setupPanel $txtSetupReport 1175 ($setupTop + 266)

# --- secpol (Local Security Policy) helpers for steps 3+4 ---

function Get-AccountSid([string]$userName) {
    try {
        return (New-Object System.Security.Principal.NTAccount($userName)).Translate([System.Security.Principal.SecurityIdentifier]).Value
    } catch {
        return $null
    }
}

function Get-AccountLocalGroups([string]$userName) {
    # Friendly names of local groups the account belongs to (e.g. "Users",
    # "Administrators") -- used to explain a privilege granted via group
    # membership rather than directly to the account.
    try {
        $adsiUser = [ADSI]"WinNT://./$userName,user"
        $groupNames = @()
        foreach ($group in $adsiUser.Groups()) {
            $groupNames += $group.GetType().InvokeMember('Name', 'GetProperty', $null, $group, $null)
        }
        return $groupNames
    } catch {
        return @()
    }
}

function Export-PrivilegeMembers([string]$privilegeName) {
    # Returns the raw member tokens (e.g. "*S-1-5-32-545", "sptest4") for one
    # User Rights Assignment entry, plus the temp inf path used (caller may
    # reuse it to write a modified version back).
    $infPath = Join-Path $env:TEMP ("secpol_{0}.inf" -f ([guid]::NewGuid().ToString('N')))
    $result = Run-ProcessCaptured 'secedit.exe' ('/export /cfg "' + $infPath + '" /areas USER_RIGHTS')
    if (-not $result.Success -or -not (Test-Path $infPath)) {
        return [PSCustomObject]@{ Members = @(); InfPath = $null }
    }
    $lines = Get-Content -Path $infPath
    $line = $lines | Where-Object { $_ -match "^$privilegeName\s*=" }
    $members = @()
    if ($line -and $line -match '^\S+\s*=\s*(.*)$') {
        $members = @($matches[1] -split ',' | ForEach-Object { $_.Trim() } | Where-Object { $_ })
    }
    return [PSCustomObject]@{ Members = $members; InfPath = $infPath }
}

function Resolve-PrivilegeMember([string]$token) {
    if ($token.StartsWith('*')) {
        try {
            return (New-Object System.Security.Principal.SecurityIdentifier($token.Substring(1))).Translate([System.Security.Principal.NTAccount]).Value
        } catch {
            return "(unresolvable SID: $token)"
        }
    }
    return $token
}

function Test-AccountPrivilegeStatus([string]$userName, [string]$privilegeName) {
    # Returns one of: 'Direct' (account itself is a member -- removable),
    # 'Group:<name>' (granted via a group the account belongs to -- NOT
    # auto-removed, since that affects every member of that group), 'None'.
    $export = Export-PrivilegeMembers $privilegeName
    if ($export.InfPath) { Remove-Item $export.InfPath -ErrorAction SilentlyContinue }
    $sid = Get-AccountSid $userName
    $resolvedMembers = $export.Members | ForEach-Object { Resolve-PrivilegeMember $_ }

    foreach ($m in $export.Members) {
        if ($m -eq "*$sid" -or $m -ieq $userName -or $m -ieq "*$userName") { return 'Direct' }
    }
    $groups = Get-AccountLocalGroups $userName
    foreach ($resolved in $resolvedMembers) {
        $shortName = ($resolved -split '\\')[-1]
        if ($groups -icontains $shortName) { return "Group:$shortName" }
    }
    return 'None'
}

function Remove-AccountFromPrivilege([string]$userName, [string]$privilegeName) {
    # Removes ONLY a direct grant to this specific account -- never touches
    # a group entry (e.g. BUILTIN\Users), since that would affect every
    # other account in that group, not just this one.
    $sid = Get-AccountSid $userName
    $export = Export-PrivilegeMembers $privilegeName
    if (-not $export.InfPath) { return [PSCustomObject]@{ Success = $false; Output = 'Could not export current security policy.' } }

    $filtered = $export.Members | Where-Object { $_ -ne "*$sid" -and $_ -notmatch ("^\*?" + [regex]::Escape($userName) + '$') }
    if ($filtered.Count -eq $export.Members.Count) {
        Remove-Item $export.InfPath -ErrorAction SilentlyContinue
        return [PSCustomObject]@{ Success = $true; Output = 'Account was not directly listed -- nothing to remove.' }
    }

    $lines = Get-Content -Path $export.InfPath
    $updated = $lines | ForEach-Object {
        if ($_ -match "^$privilegeName\s*=") { "$privilegeName = " + ($filtered -join ',') } else { $_ }
    }
    $updated | Set-Content -Path $export.InfPath
    $dbPath = Join-Path $env:TEMP ("secpol_apply_{0}.sdb" -f ([guid]::NewGuid().ToString('N')))
    $result = Run-ProcessCaptured 'secedit.exe' ('/configure /db "' + $dbPath + '" /cfg "' + $export.InfPath + '" /areas USER_RIGHTS /quiet')
    Remove-Item $export.InfPath, $dbPath -ErrorAction SilentlyContinue
    return [PSCustomObject]@{ Success = $result.Success; Output = $result.Output }
}

function Add-AccountToPrivilege([string]$userName, [string]$privilegeName) {
    # Grants a privilege directly to this account (used only for
    # SeBatchLogonRight, "Log on as a batch job" -- needed for a "run
    # whether user is logged on or not" task, same as Windows' own Task
    # Scheduler wizard grants silently when you tick that box by hand).
    $sid = Get-AccountSid $userName
    if (-not $sid) { return [PSCustomObject]@{ Success = $false; Output = "Could not resolve SID for '$userName'." } }
    $export = Export-PrivilegeMembers $privilegeName
    if (-not $export.InfPath) { return [PSCustomObject]@{ Success = $false; Output = 'Could not export current security policy.' } }

    $alreadyGranted = $export.Members | Where-Object { $_ -eq "*$sid" -or $_ -ieq $userName -or $_ -ieq "*$userName" }
    if ($alreadyGranted) {
        Remove-Item $export.InfPath -ErrorAction SilentlyContinue
        return [PSCustomObject]@{ Success = $true; Output = 'Already granted.' }
    }

    $newMembers = @($export.Members) + "*$sid"
    $lines = Get-Content -Path $export.InfPath
    $hasLine = $lines | Where-Object { $_ -match "^$privilegeName\s*=" }
    if ($hasLine) {
        $updated = $lines | ForEach-Object { if ($_ -match "^$privilegeName\s*=") { "$privilegeName = " + ($newMembers -join ',') } else { $_ } }
    } else {
        $updated = $lines | ForEach-Object { $_; if ($_ -match '^\[Privilege Rights\]') { "$privilegeName = *$sid" } }
    }
    $updated | Set-Content -Path $export.InfPath
    $dbPath = Join-Path $env:TEMP ("secpol_apply_{0}.sdb" -f ([guid]::NewGuid().ToString('N')))
    $result = Run-ProcessCaptured 'secedit.exe' ('/configure /db "' + $dbPath + '" /cfg "' + $export.InfPath + '" /areas USER_RIGHTS /quiet')
    Remove-Item $export.InfPath, $dbPath -ErrorAction SilentlyContinue
    return [PSCustomObject]@{ Success = $result.Success; Output = $result.Output }
}

# ---------------------------------------------------------------------------
# Time settings: is the PC's clock trustworthy, and can the student move it?
# Logon hours and the lock tasks both run off this PC's clock, so a clock that
# drifts, skips DST, or can be changed by the student defeats the schedule.
# ---------------------------------------------------------------------------

# Groups every student session belongs to. Any of them in a time right means
# the student can use it. Administrators / LocalService are deliberately NOT
# here: they keep the right (the time-zone service runs as LocalService).
$BroadRightsGroups = @{
    '*S-1-5-32-545' = 'Users'
    '*S-1-1-0'      = 'Everyone'
    '*S-1-5-11'     = 'Authenticated Users'
    '*S-1-5-4'      = 'Interactive'
}
$TimeRights = @{
    'SeTimeZonePrivilege'   = 'Change the time zone'
    'SeSystemtimePrivilege' = 'Change the system time'
}

# Windows' default (time.windows.com) is the one every Windows PC on earth
# uses and is often slow/unreliable from client networks, so a single server is
# a single point of failure: Windows falls back to the next one in this list
# when one doesn't answer. All of these answered from the dev PC and agreed to
# within a few ms. (Google is left out on purpose: it "smears" leap seconds,
# which is discouraged when mixed with ordinary servers.) ",0x9" = poll on the
# special interval, as a client.
$RecommendedNtpServers = 'time.cloudflare.com,0x9 il.pool.ntp.org,0x9 pool.ntp.org,0x9 time.windows.com,0x9'
$txtNtpServers.Text = $RecommendedNtpServers

function Get-NtpServerSet([string]$list) {
    # Order and case don't matter when deciding whether two lists are the same.
    return @(($list -split '\s+') | Where-Object { $_ } | ForEach-Object { $_.Trim().ToLower() } | Sort-Object -Unique)
}

function Get-BroadGroupsInMembers([string[]]$members) {
    $found = @()
    foreach ($m in $members) {
        foreach ($sid in $BroadRightsGroups.Keys) {
            $name = $BroadRightsGroups[$sid]
            if ($m -ieq $sid -or $m -ieq $name -or $m -ieq "BUILTIN\$name" -or $m -ieq "NT AUTHORITY\$name") { $found += $name }
        }
    }
    return @($found | Select-Object -Unique)
}

function Get-TimeSettingsState([string]$userName) {
    $s = [ordered]@{}

    # --- clock sync (Windows Time service) ---
    $svc = Get-CimInstance Win32_Service -Filter "Name='W32Time'" -ErrorAction SilentlyContinue
    $s.W32Present   = [bool]$svc
    $s.W32State     = if ($svc) { [string]$svc.State } else { '' }
    $s.W32StartMode = if ($svc) { [string]$svc.StartMode } else { '' }
    $wp = 'HKLM:\SYSTEM\CurrentControlSet\Services\W32Time\Parameters'
    $s.SyncType  = [string](Get-ItemProperty $wp -Name Type -ErrorAction SilentlyContinue).Type
    $s.NtpServer = [string](Get-ItemProperty $wp -Name NtpServer -ErrorAction SilentlyContinue).NtpServer
    $s.NtpClientEnabled = (Get-ItemProperty 'HKLM:\SYSTEM\CurrentControlSet\Services\W32Time\TimeProviders\NtpClient' -Name Enabled -ErrorAction SilentlyContinue).Enabled
    $status = Run-ProcessCaptured 'w32tm.exe' '/query /status'
    $s.SyncSource = ''; $s.LastSync = ''; $s.Leap = ''
    foreach ($line in ($status.Output -split "`r?`n")) {
        if ($line -match '^\s*Source:\s*(.+)$') { $s.SyncSource = $matches[1].Trim() }
        elseif ($line -match '^\s*Last Successful Sync Time:\s*(.+)$') { $s.LastSync = $matches[1].Trim() }
        elseif ($line -match '^\s*Leap Indicator:\s*(.+)$') { $s.Leap = $matches[1].Trim() }
    }
    # "Configured for NTP" is not "actually syncing": a PC can have every
    # setting right and still run on its own CMOS clock if the network blocks
    # the time server, so judge it by what w32tm says it is really doing.
    $s.IsSynced = [bool]($status.Success -and $s.SyncSource -and
        $s.SyncSource -notmatch 'Local CMOS|Free-running' -and
        $s.LastSync -notmatch 'unspecified' -and $s.Leap -notmatch '^3')
    # Right after the service restarts or a resync, Windows reports "Leap Indicator: 3" (not
    # synchronized) for a minute or two even though the sync just succeeded. Don't call that a failure.
    $s.JustSynced = $false
    if (-not $s.IsSynced -and $status.Success -and $s.SyncSource -and $s.SyncSource -notmatch 'Local CMOS|Free-running' -and $s.LastSync -notmatch 'unspecified') {
        $when = [datetime]::MinValue
        if ([datetime]::TryParse($s.LastSync, [ref]$when) -and ((Get-Date) - $when).TotalMinutes -ge -2 -and ((Get-Date) - $when).TotalMinutes -le 15) { $s.JustSynced = $true }
    }

    # --- daylight saving ---
    $dst = Get-ItemProperty 'HKLM:\SYSTEM\CurrentControlSet\Control\TimeZoneInformation' -Name DynamicDaylightTimeDisabled -ErrorAction SilentlyContinue
    $s.DstDisabledFlag = if ($dst) { [int]$dst.DynamicDaylightTimeDisabled } else { 0 }
    $s.ZoneId = [TimeZoneInfo]::Local.Id
    $s.ZoneSupportsDst = [TimeZoneInfo]::Local.SupportsDaylightSavingTime

    # --- automatic time zone (location-based) ---
    $tz = Get-CimInstance Win32_Service -Filter "Name='tzautoupdate'" -ErrorAction SilentlyContinue
    $s.AutoZonePresent   = [bool]$tz
    $s.AutoZoneStartMode = if ($tz) { [string]$tz.StartMode } else { '' }
    $s.LocationConsent   = [string](Get-ItemProperty 'HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\CapabilityAccessManager\ConsentStore\location' -Name Value -ErrorAction SilentlyContinue).Value

    # --- who may change the time / time zone ---
    $s.RightsReadable = $true
    $s.BroadRights = @{}
    foreach ($priv in $TimeRights.Keys) {
        $export = Export-PrivilegeMembers $priv
        if ($export.InfPath) { Remove-Item $export.InfPath -ErrorAction SilentlyContinue } else { $s.RightsReadable = $false }
        $s.BroadRights[$priv] = Get-BroadGroupsInMembers $export.Members
    }

    # --- is the student an administrator? (an admin can change anything) ---
    $s.StudentKnown = $false; $s.StudentIsAdmin = $false
    if ($userName) {
        $exists = $false
        try { $exists = [ADSI]::Exists("WinNT://./$userName,user") } catch { }
        if ($exists) { $s.StudentKnown = $true; $s.StudentIsAdmin = [bool]((Get-AccountLocalGroups $userName) -icontains 'Administrators') }
    }
    return [PSCustomObject]$s
}

function Format-TimeSettingsReport($s, [string]$userName) {
    $r = New-Object System.Collections.Generic.List[string]
    $r.Add('--- Time settings (clock sync, daylight saving, who can change the time) ---')

    # clock sync
    if (-not $s.W32Present) {
        $r.Add('[FAIL] Windows Time service is not present on this PC.')
    } else {
        $r.Add("Windows Time service: $($s.W32State), start type $($s.W32StartMode); sync type '$($s.SyncType)'; server '$($s.NtpServer)'")
        if ($s.W32StartMode -eq 'Auto') { $r.Add('[PASS] Service starts automatically.') }
        else { $r.Add("[WARN] Service start type is '$($s.W32StartMode)', not Automatic - it only starts when Windows decides it should.") }
        if ($s.SyncType -eq 'NoSync') { $r.Add('[FAIL] Sync type is NoSync - network time sync is turned OFF.') }
        elseif ($s.SyncType -eq 'NTP' -and -not $s.NtpServer) { $r.Add('[FAIL] Sync type is NTP but no time server is set.') }
        elseif ($s.SyncType -in @('NTP', 'NT5DS')) { $r.Add("[PASS] Network time sync is on ($($s.SyncType)).") }
        else { $r.Add("[WARN] Unrecognized sync type '$($s.SyncType)'.") }
        if ($s.SyncType -eq 'NTP' -and $s.NtpServer -and (Get-NtpServerSet $s.NtpServer).Count -lt 2) {
            $r.Add("[WARN] Only one time server is set ($($s.NtpServer.Trim())) - if it can't be reached the clock stops syncing. 'Fix Time Settings' adds backups.")
        }
        if ($s.NtpClientEnabled -ne 1 -and $s.SyncType -eq 'NTP') { $r.Add('[FAIL] The NTP client is disabled.') }
        if ($s.W32State -ne 'Running') {
            $r.Add("[FAIL] The Windows Time service is not running (it is $($s.W32State)) - the clock is not being synced.")
        } elseif ($s.IsSynced) {
            $r.Add("[PASS] Clock is actually syncing: source $($s.SyncSource), last sync $($s.LastSync).")
        } elseif ($s.JustSynced) {
            $r.Add("[WARN] Clock synced just now (source $($s.SyncSource), at $($s.LastSync)) but Windows still reports it as settling. Run the check again in a couple of minutes - this should turn into a PASS.")
        } else {
            $r.Add("[FAIL] Clock is NOT synchronized (source '$($s.SyncSource)', last successful sync '$($s.LastSync)').")
            $r.Add('       Settings can look right and it still never syncs: the usual cause is the network blocking the time server (UDP port 123), which some content filters do.')
        }
    }

    # daylight saving
    if ($s.DstDisabledFlag -eq 0) {
        $r.Add("[PASS] Automatic daylight-saving adjustment is ON (time zone '$($s.ZoneId)'$(if (-not $s.ZoneSupportsDst) { ' - this zone has no daylight saving' })).")
    } else {
        $r.Add("[FAIL] Automatic daylight-saving adjustment is OFF - the clock will be an hour wrong for half the year.")
    }

    # automatic time zone
    if (-not $s.AutoZonePresent) {
        $r.Add('Automatic time zone: service not present on this Windows edition (zone stays as set).')
    } elseif ($s.AutoZoneStartMode -eq 'Disabled') {
        $r.Add("[PASS] Automatic time zone is OFF - the zone stays fixed at '$($s.ZoneId)'.")
    } else {
        $r.Add("[WARN] Automatic time zone is ON (location consent: $($s.LocationConsent)). The zone can change by itself, which shifts every lock time.")
    }

    # who can change it
    if (-not $s.RightsReadable) {
        $r.Add('[WARN] Could not read the local security policy (needs administrator) - skipping the "who can change the time" check.')
    } else {
        foreach ($priv in $TimeRights.Keys) {
            $found = @($s.BroadRights[$priv])
            if ($found.Count -eq 0) { $r.Add("[PASS] '$($TimeRights[$priv])' is limited to administrators/system.") }
            else { $r.Add("[FAIL] '$($TimeRights[$priv])' is granted to: $($found -join ', ') - so a student can use it.") }
        }
    }
    if ($s.StudentKnown) {
        if ($s.StudentIsAdmin) { $r.Add("[FAIL] '$userName' is a member of Administrators - an admin can change the time, the hours and the tasks. Remove it from that group by hand.") }
        else { $r.Add("[PASS] '$userName' is not an administrator.") }
    }
    return $r
}

function Get-TimeSettingsFixPlan($s, [string]$autoZoneMode, [bool]$removeBroad, [string]$ntpServers = '') {
    # $ntpServers: the server list the tech wants; blank = leave an existing
    # list alone (a missing one is still filled in with the recommended list).
    # Ordered list of step ids. Nothing is planned for something already right.
    $plan = New-Object System.Collections.Generic.List[string]
    # Rights go FIRST and the clock steps last: in a live test the Windows Time
    # service was found stopped after the fix ran, and the clock sync (the step
    # that matters most) should be the last thing touched.
    if ($removeBroad -and $s.RightsReadable) {
        foreach ($priv in $TimeRights.Keys) { if (@($s.BroadRights[$priv]).Count -gt 0) { $plan.Add("RIGHTS:$priv") } }
    }
    $clockStepPlanned = $false
    if ($s.W32Present) {
        foreach ($step in @(
                @('W32_AUTOSTART', ($s.W32StartMode -ne 'Auto')),
                @('W32_START',     ($s.W32State -ne 'Running')),
                @('W32_TYPE_NTP',  ($s.SyncType -eq 'NoSync')),
                @('W32_SERVER',    ($s.SyncType -ne 'NT5DS' -and (-not $s.NtpServer -or
                                    ($ntpServers.Trim() -and ((Get-NtpServerSet $s.NtpServer) -join ' ') -ne ((Get-NtpServerSet $ntpServers) -join ' '))))),
                @('W32_PROVIDER',  ($s.SyncType -ne 'NT5DS' -and $s.NtpClientEnabled -ne 1)))) {
            if ($step[1]) { $plan.Add($step[0]); $clockStepPlanned = $true }
        }
    }
    if ($s.DstDisabledFlag -ne 0) { $plan.Add('DST_ON') }
    if ($s.AutoZonePresent) {
        if ($autoZoneMode -eq 'Fixed' -and $s.AutoZoneStartMode -ne 'Disabled') { $plan.Add('AUTOZONE_OFF') }
        if ($autoZoneMode -eq 'Automatic' -and $s.AutoZoneStartMode -eq 'Disabled') { $plan.Add('AUTOZONE_ON') }
    }
    # Resync only when the clock was touched or isn't actually syncing - and last.
    if ($s.W32Present -and ($clockStepPlanned -or -not $s.IsSynced -and -not $s.JustSynced)) { $plan.Add('W32_RESYNC') }
    return @($plan)
}

function Remove-BroadGroupsFromPrivilege([string]$privilegeName) {
    # Strips Users / Everyone / Authenticated Users / Interactive from one
    # right, leaving administrators and system services alone.
    $export = Export-PrivilegeMembers $privilegeName
    if (-not $export.InfPath) { return [PSCustomObject]@{ Success = $false; Output = 'Could not read the security policy (run as administrator).'; Removed = @() } }
    $kept = @(); $removed = @()
    foreach ($m in $export.Members) {
        if ((Get-BroadGroupsInMembers @($m)).Count -gt 0) { $removed += $m } else { $kept += $m }
    }
    if ($removed.Count -eq 0) {
        Remove-Item $export.InfPath -ErrorAction SilentlyContinue
        return [PSCustomObject]@{ Success = $true; Output = 'Nothing to remove.'; Removed = @() }
    }
    $lines = Get-Content -Path $export.InfPath
    $updated = $lines | ForEach-Object { if ($_ -match "^$privilegeName\s*=") { "$privilegeName = " + ($kept -join ',') } else { $_ } }
    $updated | Set-Content -Path $export.InfPath
    $dbPath = Join-Path $env:TEMP ("secpol_apply_{0}.sdb" -f ([guid]::NewGuid().ToString('N')))
    $result = Run-ProcessCaptured 'secedit.exe' ('/configure /db "' + $dbPath + '" /cfg "' + $export.InfPath + '" /areas USER_RIGHTS /quiet')
    Remove-Item $export.InfPath, $dbPath -ErrorAction SilentlyContinue
    return [PSCustomObject]@{ Success = $result.Success; Output = $result.Output; Removed = $removed }
}

function Resolve-NtpServers([string]$ntpServers) {
    if ($ntpServers.Trim()) { return (($ntpServers -split '\s+') | Where-Object { $_ }) -join ' ' }
    return $RecommendedNtpServers
}

function Describe-TimeFixStep([string]$step, $s, [string]$ntpServers = '') {
    switch -Regex ($step) {
        '^W32_AUTOSTART$' { return 'Set the Windows Time service to start automatically' }
        '^W32_START$'     { return 'Start the Windows Time service' }
        '^W32_TYPE_NTP$'  { return 'Turn network time sync ON (sync type NTP)' }
        '^W32_SERVER$'    { return "Set the time servers to: $((Resolve-NtpServers $ntpServers) -replace ',0x9', '')" }
        '^W32_PROVIDER$'  { return 'Enable the NTP client' }
        '^W32_RESYNC$'    { return 'Force a sync with the time server now' }
        '^DST_ON$'        { return 'Turn automatic daylight-saving adjustment ON (needs a restart to take effect)' }
        '^AUTOZONE_OFF$'  { return "Turn automatic time zone OFF (zone stays fixed at '$($s.ZoneId)')" }
        '^AUTOZONE_ON$'   { return 'Turn automatic time zone ON' }
        '^RIGHTS:(.+)$'   { $p = $matches[1]; return "Remove $((@($s.BroadRights[$p])) -join ', ') from '$($TimeRights[$p])' - affects EVERY standard account on this PC, not just the student" }
        default           { return $step }
    }
}

function Invoke-TimeSettingsFix([string[]]$plan, $s, [string]$ntpServers = '') {
    $out = New-Object System.Collections.Generic.List[string]
    $wp = 'HKLM:\SYSTEM\CurrentControlSet\Services\W32Time\Parameters'
    foreach ($step in $plan) {
        $desc = Describe-TimeFixStep $step $s $ntpServers
        try {
            switch -Regex ($step) {
                '^W32_AUTOSTART$' { $r = Run-ProcessCaptured 'sc.exe' 'config W32Time start= auto'; if (-not $r.Success) { throw $r.Output } }
                '^W32_START$'     { Start-Service -Name W32Time -ErrorAction Stop }
                '^W32_TYPE_NTP$'  { Set-ItemProperty -Path $wp -Name Type -Value 'NTP' -ErrorAction Stop }
                '^W32_SERVER$'    { Set-ItemProperty -Path $wp -Name NtpServer -Value (Resolve-NtpServers $ntpServers) -ErrorAction Stop }
                '^W32_PROVIDER$'  { Set-ItemProperty -Path 'HKLM:\SYSTEM\CurrentControlSet\Services\W32Time\TimeProviders\NtpClient' -Name Enabled -Value 1 -Type DWord -ErrorAction Stop }
                '^W32_RESYNC$' {
                    [void](Run-ProcessCaptured 'w32tm.exe' '/config /update')
                    $r = Run-ProcessCaptured 'w32tm.exe' '/resync /force'
                    if (-not $r.Success) { throw ($r.Output -replace '\s+', ' ') }
                }
                '^DST_ON$'        { Set-ItemProperty -Path 'HKLM:\SYSTEM\CurrentControlSet\Control\TimeZoneInformation' -Name DynamicDaylightTimeDisabled -Value 0 -Type DWord -ErrorAction Stop }
                '^AUTOZONE_OFF$'  { $r = Run-ProcessCaptured 'sc.exe' 'config tzautoupdate start= disabled'; if (-not $r.Success) { throw $r.Output } }
                '^AUTOZONE_ON$'   { $r = Run-ProcessCaptured 'sc.exe' 'config tzautoupdate start= demand'; if (-not $r.Success) { throw $r.Output } }
                '^RIGHTS:(.+)$'   { $r = Remove-BroadGroupsFromPrivilege $matches[1]; if (-not $r.Success) { throw $r.Output } }
            }
            $out.Add("[OK]     $desc")
        } catch {
            $out.Add("[FAILED] $desc -> $($_.Exception.Message)")
        }
    }

    # Safety net, whatever the cause: in a live test the Windows Time service
    # was found stopped after this fix had run (cause not pinned down). If it
    # was running when we started - or we were asked to set it up - and it is
    # stopped now, start it and sync again rather than report a "fixed" PC
    # whose clock isn't running.
    if ($s.W32Present -and ($s.W32State -eq 'Running' -or @($plan | Where-Object { $_ -like 'W32_*' }).Count -gt 0)) {
        $svc = Get-Service -Name W32Time -ErrorAction SilentlyContinue
        $mode = (Get-CimInstance Win32_Service -Filter "Name='W32Time'" -ErrorAction SilentlyContinue).StartMode
        if ($svc -and $svc.Status -ne 'Running' -and $mode -ne 'Disabled') {
            try {
                Start-Service -Name W32Time -ErrorAction Stop
                [void](Run-ProcessCaptured 'w32tm.exe' '/resync /force')
                $out.Add('[OK]     The Windows Time service had stopped during the fix - started it again and synced')
            } catch {
                $out.Add("[FAILED] The Windows Time service is stopped and could not be started -> $($_.Exception.Message)")
            }
        }
    }
    if (@($plan | Where-Object { $_ -eq 'W32_RESYNC' }).Count -gt 0) { Start-Sleep -Seconds 3 }  # let the sync settle before the re-check
    return $out
}

$btnCreateLogoutTasks.Add_Click({
    $userName = $txtUser.Text.Trim()
    if (-not $userName) {
        [System.Windows.Forms.MessageBox]::Show('Enter a Windows account name first.', 'Missing account') | Out-Null
        return
    }
    # Blank password is the normal case (these accounts have none) = interactive-only tasks.
    $password = $txtLogoutPassword.Text
    $modeText = if ($password) { 'run whether logged on or not' } else { 'interactive-only - no password needed; fires while that account is logged on' }
    $groups = @(Get-LogoutKickoffGroups $Grid)
    $action = if ($cmbLogoutAction.SelectedItem -eq 'Full log off') { 'LogOff' } else { 'Lock' }
    $preview = if ($groups.Count -eq 0) { '(no lockout points - any existing logout tasks for this account will be removed)' } else { ($groups | ForEach-Object { "$($_.TimeHHmm) on $(($_.Days | ForEach-Object { $DayNames[$_] }) -join ',')" }) -join "`n" }
    $confirm = [System.Windows.Forms.MessageBox]::Show(
        "This makes the weekly logout tasks on $($env:COMPUTERNAME) match the grid shown on the Daily Use tab. Tasks run as '$userName' ($modeText) and $($action.ToLower()) the session at:`n`n$preview`n`nTasks that already match are kept; this account's other lock/log-off tasks (old or stale ones) are REMOVED. Continue?",
        'Sync logout tasks', 'YesNo', 'Question')
    if ($confirm -ne 'Yes') { return }
    $report = New-Object System.Text.StringBuilder
    [void]$report.AppendLine("=== Logout tasks synced to the grid for '$userName' on $($env:COMPUTERNAME) ($(Get-Date -Format 'yyyy-MM-dd HH:mm')) ===")
    [void]$report.AppendLine("Mode: $modeText")
    if ($password) {
        $batchResult = Add-AccountToPrivilege -userName $userName -privilegeName 'SeBatchLogonRight'
        Write-DebugLog "Add-AccountToPrivilege '$userName' SeBatchLogonRight: success=$($batchResult.Success) output=$($batchResult.Output)"
        [void]$report.AppendLine("$(if ($batchResult.Success) { '[OK]' } else { '[FAIL]' }) Granted 'Log on as a batch job' right$(if (-not $batchResult.Success) { ': ' + $batchResult.Output })")
    }
    $lines = @()
    foreach ($acct in @(Get-TargetAccounts)) { $lines += "-- $acct --"; $lines += @(Sync-LogoutTasks -userName $acct -grid $Grid) }
    foreach ($l in $lines) { [void]$report.AppendLine($l) }
    Write-DebugLog "Sync-LogoutTasks '$userName':`n$($lines -join "`n")"
    $txtSetupReport.Text = $report.ToString()
    $txtLogoutPassword.Text = ''
    Refresh-ScheduledTasksList
})

function Get-SelectedAutoZoneMode {
    switch ($cmbAutoZone.SelectedIndex) { 0 { 'Fixed' } 1 { 'Automatic' } default { 'Leave' } }
}

$btnCheckTime.Add_Click({
    $userName = $txtUser.Text.Trim()
    $state = Get-TimeSettingsState $userName
    $lines = Format-TimeSettingsReport $state $userName
    $txtSetupReport.Text = ("=== Time settings on $($env:COMPUTERNAME) ($(Get-Date -Format 'yyyy-MM-dd HH:mm')) ===" + "`r`n" + ($lines -join "`r`n"))
    Write-DebugLog ("Check Time Settings for '$userName':`n" + ($lines -join "`n"))
})

$btnFixTime.Add_Click({
    $userName = $txtUser.Text.Trim()
    $state = Get-TimeSettingsState $userName
    $ntp = $txtNtpServers.Text.Trim()
    $plan = Get-TimeSettingsFixPlan $state (Get-SelectedAutoZoneMode) ([bool]$chkRemoveBroad.Checked) $ntp
    if ($plan.Count -eq 0) {
        $lines = Format-TimeSettingsReport $state $userName
        $txtSetupReport.Text = ("=== Time settings on $($env:COMPUTERNAME): nothing to fix ===" + "`r`n" + ($lines -join "`r`n"))
        return
    }
    $desc = ($plan | ForEach-Object { '  - ' + (Describe-TimeFixStep $_ $state $ntp) }) -join "`n"
    $confirm = [System.Windows.Forms.MessageBox]::Show(
        "This will make these changes on $($env:COMPUTERNAME):`n`n$desc`n`nContinue?",
        'Confirm time-settings fix', 'YesNo', 'Warning')
    if ($confirm -ne 'Yes') { return }

    $results = Invoke-TimeSettingsFix $plan $state $ntp
    Write-DebugLog ("Fix Time Settings for '$userName':`n" + ($results -join "`n"))
    $after = Get-TimeSettingsState $userName
    $lines = Format-TimeSettingsReport $after $userName
    $txtSetupReport.Text = ("=== Time settings fix on $($env:COMPUTERNAME) ($(Get-Date -Format 'yyyy-MM-dd HH:mm')) ===" + "`r`n" +
        ($results -join "`r`n") + "`r`n`r`n=== Re-check after the fix ===`r`n" + ($lines -join "`r`n"))
})

$btnCheckSecpol.Add_Click({
    $targets = @(Get-TargetAccounts)
    if ($targets.Count -eq 0) {
        [System.Windows.Forms.MessageBox]::Show('Enter a Windows account name (or pick accounts) first.', 'Missing account') | Out-Null
        return
    }
    $privileges = @('SeTimeZonePrivilege', 'SeSystemtimePrivilege')
    $report = New-Object System.Text.StringBuilder
    [void]$report.AppendLine("=== Time-change rights on $($env:COMPUTERNAME) ($(Get-Date -Format 'yyyy-MM-dd HH:mm')) - $($targets.Count) account(s) ===")
    $anyGroup = $false
    foreach ($userName in $targets) {
        $statuses = @{}
        foreach ($p in $privileges) { $statuses[$p] = Test-AccountPrivilegeStatus -userName $userName -privilegeName $p }
        $directOnes = @($privileges | Where-Object { $statuses[$_] -eq 'Direct' })
        $groupOnes = @($privileges | Where-Object { $statuses[$_] -like 'Group:*' })
        [void]$report.AppendLine('')
        [void]$report.AppendLine("--- $userName ---")
        foreach ($p in $privileges) {
            $label = if ($p -eq 'SeTimeZonePrivilege') { 'Change the time zone' } else { 'Change the system time' }
            [void]$report.AppendLine("$label ($p): $($statuses[$p])")
        }
        if ($directOnes.Count -gt 0) {
            $confirm = [System.Windows.Forms.MessageBox]::Show(
                "'$userName' is DIRECTLY granted: $($directOnes -join ', ')`n`nRemove '$userName' from these (matches checklist step 3/4)? This does not touch any group-level entries.",
                'Remove direct grants?', 'YesNo', 'Question')
            if ($confirm -eq 'Yes') {
                foreach ($p in $directOnes) {
                    $result = Remove-AccountFromPrivilege -userName $userName -privilegeName $p
                    [void]$report.AppendLine("Removed from $p : $(if ($result.Success) { 'OK' } else { 'FAILED - ' + $result.Output })")
                    Write-DebugLog "Remove-AccountFromPrivilege '$userName' '$p': success=$($result.Success) output=$($result.Output)"
                }
                [void]$report.AppendLine('IMPORTANT (checklist step): sign the user out now for this to take effect for their session.')
            }
        }
        if ($groupOnes.Count -gt 0) { $anyGroup = $true }
        if ($directOnes.Count -eq 0 -and $groupOnes.Count -eq 0) { [void]$report.AppendLine('Not granted either way -- already locked down.') }
    }
    if ($anyGroup) {
        [void]$report.AppendLine('')
        [void]$report.AppendLine('NOTE: the "Group:" entries above mean the right comes from a GROUP these accounts belong to (e.g. BUILTIN\Users), not the account directly.')
        [void]$report.AppendLine('This button never removes a group-wide entry, since that affects every account in that group on this PC.')
        [void]$report.AppendLine("The 'Fix Time Settings' button (with its Users/Everyone box ticked) does remove it - use that on a dedicated student PC.")
    }
    $txtSetupReport.Text = $report.ToString()
})

$btnVerifyAll.Add_Click({
    $targets = @(Get-TargetAccounts)
    if ($targets.Count -eq 0) {
        [System.Windows.Forms.MessageBox]::Show('Enter a Windows account name (or pick accounts) first.', 'Missing account') | Out-Null
        return
    }
    $report = New-Object System.Text.StringBuilder
    [void]$report.AppendLine("=== Verify setup for $($targets -join ', ') on $($env:COMPUTERNAME) ($(Get-Date -Format 'yyyy-MM-dd HH:mm')) ===")
    [void]$report.AppendLine('')
    foreach ($userName in $targets) {
    [void]$report.AppendLine("######## Account: $userName ########")

    # 1) net user times vs. current grid
    [void]$report.AppendLine('--- Step 2: net user /times ---')
    $existing = Get-CurrentLoginHours $userName
    if ($null -eq $existing) {
        [void]$report.AppendLine("[FAIL] Could not read account '$userName' on this PC -- does it exist?")
    } else {
        $existingGrid = New-Object 'bool[,]' 7,24
        for ($d = 0; $d -lt 7; $d++) { for ($h = 0; $h -lt 24; $h++) { $existingGrid[$d,$h] = $existing[$d,$h] } }
        $tempGridBackup = New-Object 'bool[,]' 7,24
        for ($d = 0; $d -lt 7; $d++) { for ($h = 0; $h -lt 24; $h++) { $tempGridBackup[$d,$h] = $Grid[$d,$h] } }
        $gridTimes = Build-TimesString
        for ($d = 0; $d -lt 7; $d++) { for ($h = 0; $h -lt 24; $h++) { $Grid[$d,$h] = $existing[$d,$h] } }
        $liveTimes = Build-TimesString
        for ($d = 0; $d -lt 7; $d++) { for ($h = 0; $h -lt 24; $h++) { $Grid[$d,$h] = $tempGridBackup[$d,$h] } }

        [void]$report.AppendLine("Currently set on this PC: net user $userName /times:$liveTimes")
        [void]$report.AppendLine("Grid shown in this editor: net user $userName /times:$gridTimes")
        if ($liveTimes -eq $gridTimes) {
            [void]$report.AppendLine('[PASS] Matches the grid shown above.')
        } else {
            [void]$report.AppendLine('[WARN] Does NOT match the grid currently shown in the editor -- load current schedule to compare, or this account may need re-applying.')
        }
    }

    # 2) Weekly logout tasks vs. expected kickoff points from the grid
    [void]$report.AppendLine('')
    [void]$report.AppendLine('--- Step 3: weekly logout Scheduled Tasks ---')
    [void]$report.AppendLine('(Matched by run-as user + lock/logoff action, not by task name - finds manually-created tasks too, e.g. "Lock PC".)')
    $expected = Get-LogoutKickoffGroups $Grid
    $actualTasks = Get-AccountLogoutLikeTasks $userName
    foreach ($ln in (Get-LogoutTaskCoverageLines $Grid $expected $actualTasks $userName)) { [void]$report.AppendLine($ln) }

    # 3) secpol status
    [void]$report.AppendLine('')
    [void]$report.AppendLine('--- Steps 4+6: time-change rights (secpol) ---')
    foreach ($p in @('SeTimeZonePrivilege', 'SeSystemtimePrivilege')) {
        $status = Test-AccountPrivilegeStatus -userName $userName -privilegeName $p
        $label = if ($p -eq 'SeTimeZonePrivilege') { 'Change the time zone' } else { 'Change the system time' }
        if ($status -eq 'None') {
            [void]$report.AppendLine("[PASS] $label - not granted.")
        } elseif ($status -eq 'Direct') {
            [void]$report.AppendLine("[FAIL] $label - still directly granted to '$userName'. Use the 3+4 button above to remove it.")
        } else {
            [void]$report.AppendLine("[WARN] $label - granted via group ($status), not directly to this account. Use 'Fix Time Settings' (with the Users/Everyone box ticked) to close it.")
        }
    }

    [void]$report.AppendLine('')
    }   # end of the per-account loop

    # clock sync + DST + who can change the time (same for the whole PC; checked once, against the first account)
    $userName = $targets[0]
    foreach ($line in (Format-TimeSettingsReport (Get-TimeSettingsState $userName) $userName)) { [void]$report.AppendLine($line) }

    $txtSetupReport.Text = $report.ToString()
    Write-DebugLog "Verify Everything report for $($targets -join ','):`n$($report.ToString())"
})

# ---------------------------------------------------------------------------
# Network Push -- push the current grid's schedule from THIS PC to other
# student PCs on the same local network via PowerShell Remoting (WinRM).
#
# These are standalone (non-domain) lab PCs, so the setup has two halves:
#   - Every PC that might ever be pushed TO needs WinRM enabled on itself
#     (step 1 below) -- including whichever PC is acting as the "controller"
#     today, since any of them could be the target another day.
#   - The PC doing the pushing needs its targets in WSMan TrustedHosts (a
#     workgroup requirement, handled automatically per push) and a shared
#     admin credential valid on all of them.
# ---------------------------------------------------------------------------

$PushTargetsPath = Join-Path $ScriptDir 'PushTargets.txt'
$script:pushCredential = $null

function Enable-ThisPCForPushes {
    # Target-side one-time setup: turns on the WinRM listener + firewall
    # rule, and lets a local (non-domain) admin account authenticate with a
    # full token over the network -- without this, a local admin connecting
    # remotely gets silently downgraded to a non-elevated token and "net
    # user" fails with access denied even though the credential is correct.
    $report = New-Object System.Text.StringBuilder
    try {
        Enable-PSRemoting -Force -SkipNetworkProfileCheck -ErrorAction Stop | Out-Null
        [void]$report.AppendLine('[OK] WinRM listener + firewall rule enabled.')
    } catch {
        [void]$report.AppendLine("[FAIL] Enable-PSRemoting: $($_.Exception.Message)")
    }
    try {
        Set-ItemProperty -Path 'HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Policies\System' -Name LocalAccountTokenFilterPolicy -Value 1 -Type DWord -Force -ErrorAction Stop
        [void]$report.AppendLine('[OK] Local admin accounts can now authenticate remotely with a full token.')
    } catch {
        [void]$report.AppendLine("[FAIL] LocalAccountTokenFilterPolicy: $($_.Exception.Message)")
    }
    return $report.ToString()
}

function ConvertFrom-PushTargetText([string]$text) {
    # One target per line (a name or an IP). Anything after a "#" is a note, so
    # a scanned entry can read "192.168.1.51   # DESKTOP-ABC" and still push to the IP.
    return @(($text -split "`r?`n") | ForEach-Object { ($_ -replace '\s*#.*$', '').Trim() } | Where-Object { $_ })
}

function Get-PushTargetText {
    if (Test-Path $PushTargetsPath) { return ((Get-Content -Path $PushTargetsPath) -join "`r`n") }
    return ''
}

function Save-PushTargetText([string]$text) {
    # Saved exactly as typed, notes included.
    ($text -split "`r?`n") | Set-Content -Path $PushTargetsPath -Encoding UTF8
}

function Ensure-TrustedHost([string]$targetName) {
    $current = (Get-Item WSMan:\localhost\Client\TrustedHosts -ErrorAction SilentlyContinue).Value
    $list = @()
    # @() matters: a single existing entry would otherwise collapse to a plain string
    # and "+=" would glue the new host onto it ("192.168.1.2" + "192.168.1.3" = "192.168.1.2192.168.1.3").
    if ($current) { $list = @($current -split ',' | ForEach-Object { $_.Trim() } | Where-Object { $_ }) }
    if ($list -icontains $targetName -or $list -contains '*') { return }
    $list += $targetName
    Set-Item WSMan:\localhost\Client\TrustedHosts -Value ($list -join ',') -Force
}

function Push-TimesToMachine([string]$targetName, [System.Management.Automation.PSCredential]$credential, [string]$userName, [string]$timesValue, $groups = @(), [string]$action = 'Lock', $lockOpts = $null, $pictureOpts = $null) {
    try {
        Ensure-TrustedHost $targetName
    } catch {
        return [PSCustomObject]@{ Success = $false; Output = "Could not update TrustedHosts: $($_.Exception.Message)" }
    }
    try {
        $r = Invoke-Command -ComputerName $targetName -Credential $credential -ErrorAction Stop -ScriptBlock {
            param($u, $t, $groups, $action, $syncText, $lockText, $lockOpts, $picText, $picOpts)
            # Some PCs name the student account differently (Students / Student / Stu / Bochurim). If the typed
            # account is one of those standard names and THIS PC has no such account but has exactly ONE of the
            # others, use that one (and say so in the result). Anything else (user1, ...) is never swapped.
            $stdNames = @('Students', 'Student', 'Stu', 'Bochurim')
            $accountNote = ''
            $hasAcct = { param($n) try { [ADSI]::Exists("WinNT://./$n,user") } catch { $false } }
            if (($stdNames -icontains $u) -and -not (& $hasAcct $u)) {
                $alts = @($stdNames | Where-Object { $_ -ine $u -and (& $hasAcct $_) })
                if ($alts.Count -eq 1) { $accountNote = "this PC has no account '$u' - used '$($alts[0])' instead"; $u = $alts[0] }
                elseif ($alts.Count -gt 1) { $accountNote = "this PC has no account '$u' and several look like student accounts ($($alts -join ', ')) - not guessing" }
            }
            # stderr arrives as error records; keep just each line's message (no PowerShell position noise)
            $text = ((& net.exe user $u "/times:$t" 2>&1) | ForEach-Object { if ($_ -is [System.Management.Automation.ErrorRecord]) { $_.Exception.Message } else { "$_" } } | Where-Object { $_ -and $_.Trim() }) -join ' '
            $code = $LASTEXITCODE
            if ($accountNote) { $text = "[$accountNote] $text" }
            $syncLines = @()
            if ($code -eq 0) {
                # hours changed -> the logout tasks on THIS PC must follow (same code the local button runs)
                $syncLines = @(& ([scriptblock]::Create($syncText)) $u $groups $action '')
            }
            if ($code -eq 0 -and $lockOpts) {
                # optional: the desktop lock-down / Spotlight-off settings go to this PC too
                $syncLines += @(& ([scriptblock]::Create($lockText)) $u 'Apply' $lockOpts)
            }
            if ($code -eq 0 -and $picOpts) {
                # optional: the fixed lock-screen / sign-in picture (whole PC)
                $syncLines += @(& ([scriptblock]::Create($picText)) 'Apply' $picOpts)
            }
            [PSCustomObject]@{ Computer = $env:COMPUTERNAME; ExitCode = $code; Text = $text; Sync = $syncLines }
        } -ArgumentList $userName, $timesValue, @($groups), $action, $SyncLogoutTasksBlock.ToString(), $DesktopLockBlock.ToString(), $lockOpts, $LockScreenPictureBlock.ToString(), $pictureOpts
        # net.exe failing (e.g. that PC has no such account) does not throw, so
        # judge success by its exit code, and say which PC answered.
        $syncLines = @($r.Sync)
        $syncBad = @($syncLines | Where-Object { "$_" -match '^\[FAIL\]' })
        $ok = ($r.ExitCode -eq 0 -and $syncBad.Count -eq 0)
        $msg = "on $($r.Computer): hours: $($r.Text -replace '\s+', ' ')"
        if ($r.ExitCode -eq 0) { $msg += " | logout tasks: " + $(if ($syncBad.Count -eq 0) { "synced ($($syncLines.Count) step(s))" } else { "PROBLEM - " + (($syncBad | ForEach-Object { "$_" }) -join ' ') }) }
        return [PSCustomObject]@{ Success = $ok; Output = $msg }
    } catch {
        return [PSCustomObject]@{ Success = $false; Output = $_.Exception.Message }
    }
}

# ---- finding the other PCs on the network ----

function ConvertTo-IpNumber([string]$ip) {
    $b = ([System.Net.IPAddress]::Parse($ip)).GetAddressBytes(); [Array]::Reverse($b)
    return [long][BitConverter]::ToUInt32($b, 0)
}
function ConvertFrom-IpNumber([long]$n) {
    $b = [BitConverter]::GetBytes([uint32]$n); [Array]::Reverse($b)
    return ([System.Net.IPAddress]::new($b)).ToString()
}

function Get-LocalNetworkInfo {
    # This PC's own IPv4 addresses, and the address ranges to sweep (one per
    # real network card; never wider than /22 = ~1000 addresses).
    $own = @(); $nets = @{}
    $skip = 'vEthernet|VirtualBox|VMware|Loopback|Bluetooth|Hyper-V|Tailscale|ZeroTier|WireGuard|VPN'
    foreach ($a in (Get-NetIPAddress -AddressFamily IPv4 -ErrorAction SilentlyContinue)) {
        $ip = [string]$a.IPAddress
        if ($ip -like '127.*' -or $ip -like '169.254.*') { continue }
        $own += $ip
        if ($a.InterfaceAlias -match $skip -or $a.AddressState -ne 'Preferred') { continue }
        $prefix = [int]$a.PrefixLength
        if ($prefix -gt 30) { continue }
        if ($prefix -lt 22) { $prefix = 22 }
        $mask = ([long]0xFFFFFFFF -shl (32 - $prefix)) -band [long]0xFFFFFFFF
        $net = (ConvertTo-IpNumber $ip) -band $mask
        $nets["$net/$prefix"] = [PSCustomObject]@{ Network = $net; Prefix = $prefix; Mask = $mask }
    }
    $gw = @(Get-NetRoute -DestinationPrefix '0.0.0.0/0' -ErrorAction SilentlyContinue | ForEach-Object { [string]$_.NextHop }) | Where-Object { $_ -and $_ -ne '0.0.0.0' }
    return [PSCustomObject]@{ OwnIps = $own; Networks = @($nets.Values); Gateways = @($gw) }
}

function Test-TcpPortMany([string[]]$ips, [int]$port, [int]$timeoutMs) {
    # Tries every address at once. 'open' = something is listening; 'refused'
    # = the PC is there but nothing listens on that port (so it is alive);
    # 'none' = no answer (dead, or a firewall quietly dropping it).
    $state = @{}; $clients = @{}; $handles = @{}
    foreach ($ip in $ips) {
        $c = New-Object System.Net.Sockets.TcpClient
        $clients[$ip] = $c
        try { $handles[$ip] = $c.BeginConnect($ip, $port, $null, $null) } catch { }
    }
    $deadline = (Get-Date).AddMilliseconds($timeoutMs)
    while ((Get-Date) -lt $deadline) {
        if (@($handles.Values | Where-Object { -not $_.IsCompleted }).Count -eq 0) { break }
        Start-Sleep -Milliseconds 100
    }
    foreach ($ip in $ips) {
        $state[$ip] = 'none'
        if ($handles.ContainsKey($ip) -and $handles[$ip].IsCompleted) {
            try {
                $clients[$ip].EndConnect($handles[$ip])
                if ($clients[$ip].Connected) { $state[$ip] = 'open' }
            } catch {
                $ex = $_.Exception; while ($ex.InnerException) { $ex = $ex.InnerException }
                if ($ex -is [System.Net.Sockets.SocketException] -and $ex.SocketErrorCode -eq 'ConnectionRefused') { $state[$ip] = 'refused' }
            }
        }
        try { $clients[$ip].Close() } catch { }
    }
    return $state
}

function Get-ArpAliveHosts([long]$lo, [long]$hi) {
    # Any PC on this network segment that sent us even one packet back is in the
    # ARP table - including ones whose firewall ignores pings and port probes.
    $alive = @()
    foreach ($n in (Get-NetNeighbor -AddressFamily IPv4 -ErrorAction SilentlyContinue)) {
        if ($n.State -notin @('Reachable', 'Stale', 'Delay', 'Probe', 'Permanent')) { continue }
        $mac = [string]$n.LinkLayerAddress
        if (-not $mac -or $mac -match '^(FF-FF-FF-FF-FF-FF|00-00-00-00-00-00|01-00-5E)') { continue }
        $num = ConvertTo-IpNumber ([string]$n.IPAddress)
        if ($num -ge $lo -and $num -le $hi) { $alive += [string]$n.IPAddress }
    }
    return $alive
}

function Resolve-HostNamesMany([string[]]$ips, [string[]]$netbiosIps = @(), [int]$timeoutMs = 4000) {
    # Name for each IP: reverse DNS first, then - for the PCs in $netbiosIps -
    # the NetBIOS name (the one a Windows PC calls itself) when DNS has none.
    $names = @{}; $handles = @{}
    foreach ($ip in $ips) { try { $handles[$ip] = [System.Net.Dns]::BeginGetHostEntry($ip, $null, $null) } catch { } }
    $deadline = (Get-Date).AddMilliseconds($timeoutMs)
    while ((Get-Date) -lt $deadline -and @($handles.Values | Where-Object { -not $_.IsCompleted }).Count -gt 0) { Start-Sleep -Milliseconds 100 }
    foreach ($ip in @($handles.Keys)) {
        if ($handles[$ip].IsCompleted) {
            try {
                $h = [System.Net.Dns]::EndGetHostEntry($handles[$ip]).HostName
                # With no DNS record Windows hands the address itself back as the "name".
                $parsed = $null
                if ($h -and -not [System.Net.IPAddress]::TryParse($h, [ref]$parsed)) { $names[$ip] = ($h -split '\.')[0] }
            } catch { }
        }
    }
    $rest = @($netbiosIps | Where-Object { -not $names.ContainsKey($_) } | Select-Object -First 60)
    $procs = @{}
    foreach ($ip in $rest) {
        try {
            $psi = New-Object System.Diagnostics.ProcessStartInfo; $psi.FileName = 'nbtstat.exe'; $psi.Arguments = "-A $ip"
            $psi.RedirectStandardOutput = $true; $psi.UseShellExecute = $false; $psi.CreateNoWindow = $true
            $p = New-Object System.Diagnostics.Process; $p.StartInfo = $psi; [void]$p.Start(); $procs[$ip] = $p
        } catch { }
    }
    $nbDeadline = (Get-Date).AddMilliseconds(6000)
    foreach ($ip in @($procs.Keys)) {
        $left = [int][Math]::Max(200, ($nbDeadline - (Get-Date)).TotalMilliseconds)
        if ($procs[$ip].WaitForExit($left)) {
            $out = $procs[$ip].StandardOutput.ReadToEnd()
            if ($out -match '(?m)^\s*(\S+)\s+<00>\s+UNIQUE') { $names[$ip] = $matches[1] }
        } else { try { $procs[$ip].Kill() } catch { } }
    }
    return $names
}

function Find-LanPcs {
    # Returns one row per PC found on this PC's network(s):
    #   Ip, Name, Ready ($true only when WinRM actually answered), Status.
    $info = Get-LocalNetworkInfo
    $all = New-Object System.Collections.Generic.List[string]
    $ranges = @()
    foreach ($n in $info.Networks) {
        $first = $n.Network + 1
        $last  = ($n.Network -bor ((-bnot $n.Mask) -band [long]0xFFFFFFFF)) - 1
        $ranges += [PSCustomObject]@{ Lo = $first; Hi = $last }
        for ($i = $first; $i -le $last; $i++) { $all.Add((ConvertFrom-IpNumber $i)) }
    }
    $ips = @($all | Select-Object -Unique)
    if ($ips.Count -eq 0) { return @() }

    $winrm = Test-TcpPortMany $ips 5985 2500
    $alive = New-Object System.Collections.Generic.HashSet[string]
    foreach ($ip in $ips) { if ($winrm[$ip] -ne 'none') { [void]$alive.Add($ip) } }
    foreach ($r in $ranges) { foreach ($ip in (Get-ArpAliveHosts $r.Lo $r.Hi)) { [void]$alive.Add($ip) } }

    # Which of the live devices are Windows PCs? A Windows PC normally answers on
    # the file-sharing / RPC ports; printers, phones, cameras and routers don't.
    $aliveList = @($alive)
    $p445 = Test-TcpPortMany $aliveList 445 2500
    $p135 = Test-TcpPortMany $aliveList 135 2500
    $windowsLike = @($aliveList | Where-Object { $p445[$_] -eq 'open' -or $p135[$_] -eq 'open' -or $winrm[$_] -eq 'open' })

    $names = Resolve-HostNamesMany $aliveList $windowsLike
    $rows = @()
    foreach ($ip in $aliveList) {
        $isMe = $info.OwnIps -contains $ip
        $isRouter = $info.Gateways -contains $ip
        $isWin = ($windowsLike -contains $ip) -or $isMe
        $ready = $false
        if ($winrm[$ip] -eq 'open' -and -not $isMe) {
            try { $null = Test-WSMan -ComputerName $ip -ErrorAction Stop; $ready = $true } catch { }
        }
        $status = if ($isMe) { '(this PC)' }
            elseif ($isRouter) { '(router)' }
            elseif ($ready) { 'READY - WinRM answered' }
            elseif ($winrm[$ip] -eq 'open') { 'Windows PC - port open but WinRM did not answer' }
            elseif ($isWin) { 'Windows PC - WinRM is OFF: run Step 1 on it' }
            else { 'other device (printer / phone / camera?)' }
        $order = if ($ready) { 0 } elseif ($isWin -and -not $isMe) { 1 } elseif ($isMe) { 2 } elseif ($isRouter) { 4 } else { 3 }
        $rows += [PSCustomObject]@{ Ip = $ip; Name = [string]$names[$ip]; Ready = $ready; IsWindows = $isWin; IsOther = (-not $isWin -and -not $isRouter); Status = $status; IsThisPc = $isMe; Order = $order }
    }
    return @($rows | Sort-Object Order, @{ Expression = { ConvertTo-IpNumber $_.Ip } })
}

function Test-PushReadiness([string]$target, [System.Management.Automation.PSCredential]$credential) {
    # The only real proof a push will work: log in with the shared credential
    # and ask the PC who it is and whether it got a full administrator token.
    try { Ensure-TrustedHost $target } catch {
        return [PSCustomObject]@{ Ok = $false; Name = ''; Text = "could not update TrustedHosts: $($_.Exception.Message)" }
    }
    try {
        $r = Invoke-Command -ComputerName $target -Credential $credential -ErrorAction Stop -ScriptBlock {
            $id = [Security.Principal.WindowsIdentity]::GetCurrent()
            [PSCustomObject]@{
                Computer = $env:COMPUTERNAME; User = $id.Name
                IsAdmin  = ([Security.Principal.WindowsPrincipal]$id).IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)
            }
        }
        if ($r.IsAdmin) { return [PSCustomObject]@{ Ok = $true; Name = $r.Computer; Text = "logged in as $($r.User) with full administrator rights" } }
        return [PSCustomObject]@{ Ok = $false; Name = $r.Computer; Text = "logged in as $($r.User) but WITHOUT full administrator rights - run 'Enable This PC To Receive Pushes' on it (Step 1)" }
    } catch {
        $m = ($_.Exception.Message -replace '\s+', ' ')
        if ($m -match 'TrustedHosts') { $m += "  [HINT: this PC does not trust '$target' - check (Get-Item WSMan:\localhost\Client\TrustedHosts).Value is a comma-separated list]" }
        return [PSCustomObject]@{ Ok = $false; Name = ''; Text = $m }
    }
}

# --- Network Push tab UI ---
$netTop = 20

$lblNetHeader = New-Object System.Windows.Forms.Label
$lblNetHeader.Text = 'Push the current grid (from the Daily Use tab) to other student PCs on this same local network.'
$lblNetHeader.Location = New-Object System.Drawing.Point(15, $netTop)
$lblNetHeader.Size = New-Object System.Drawing.Size(1290, 18)
$networkPanel.Controls.Add($lblNetHeader)

$lblNetStep1 = New-Object System.Windows.Forms.Label
$lblNetStep1.Text = 'Step 1 (run ONCE on EVERY PC in the group, including whichever one is acting as controller today):'
$lblNetStep1.Location = New-Object System.Drawing.Point(15, ($netTop + 30))
$lblNetStep1.AutoSize = $true
$networkPanel.Controls.Add($lblNetStep1)

$btnEnablePush = New-Object System.Windows.Forms.Button
$btnEnablePush.Text = 'Enable This PC To Receive Pushes'
$btnEnablePush.Location = New-Object System.Drawing.Point(15, ($netTop + 54))
$btnEnablePush.Size = New-Object System.Drawing.Size(260, 30)
$networkPanel.Controls.Add($btnEnablePush)

$lblNetStep2 = New-Object System.Windows.Forms.Label
$lblNetStep2.Text = "Step 2: scan for PCs, tick the ones to push to (ready ones are ticked for you), set the shared admin credential, then push:"
$lblNetStep2.Location = New-Object System.Drawing.Point(15, ($netTop + 100))
$lblNetStep2.Size = New-Object System.Drawing.Size(1290, 18)
$networkPanel.Controls.Add($lblNetStep2)

# --- ONE list: the PCs found on the network (plus any added by hand). Ticked rows are what gets pushed to. ---
$btnScanNet = New-Object System.Windows.Forms.Button
$btnScanNet.Text = 'Scan Network For PCs'
$btnScanNet.Location = New-Object System.Drawing.Point(15, ($netTop + 122))
$btnScanNet.Size = New-Object System.Drawing.Size(190, 30)
$networkPanel.Controls.Add($btnScanNet)

$lblScanStatus = New-Object System.Windows.Forms.Label
$lblScanStatus.Text = 'Finds every PC on this PC''s network and shows which are ready to receive a push.'
$lblScanStatus.Location = New-Object System.Drawing.Point(215, ($netTop + 129))
$lblScanStatus.Size = New-Object System.Drawing.Size(930, 18)
$lblScanStatus.ForeColor = [System.Drawing.Color]::DimGray
$networkPanel.Controls.Add($lblScanStatus)

$chkShowOther = New-Object System.Windows.Forms.CheckBox
$chkShowOther.Text = 'Show other devices'
$chkShowOther.Location = New-Object System.Drawing.Point(1160, ($netTop + 128))
$chkShowOther.AutoSize = $true
$networkPanel.Controls.Add($chkShowOther)

$lstFound = New-Object System.Windows.Forms.CheckedListBox
$lstFound.CheckOnClick = $true
$lstFound.HorizontalScrollbar = $true
$lstFound.Font = New-Object System.Drawing.Font('Consolas', 9)
$lstFound.Location = New-Object System.Drawing.Point(15, ($netTop + 158))
$lstFound.Size = New-Object System.Drawing.Size(1290, 130)
$networkPanel.Controls.Add($lstFound)

$btnSetCredential = New-Object System.Windows.Forms.Button
$btnSetCredential.Text = 'Set Shared Admin Credential'
$btnSetCredential.Location = New-Object System.Drawing.Point(15, ($netTop + 296))
$btnSetCredential.Size = New-Object System.Drawing.Size(220, 30)
$networkPanel.Controls.Add($btnSetCredential)

$lblCredStatus = New-Object System.Windows.Forms.Label
$lblCredStatus.Text = 'No credential set yet.'
$lblCredStatus.Location = New-Object System.Drawing.Point(430, ($netTop + 304))
$lblCredStatus.Size = New-Object System.Drawing.Size(155, 18)
$lblCredStatus.ForeColor = [System.Drawing.Color]::DimGray
$networkPanel.Controls.Add($lblCredStatus)

$chkRememberCred = New-Object System.Windows.Forms.CheckBox
$chkRememberCred.Text = 'Remember on this PC'
$chkRememberCred.Location = New-Object System.Drawing.Point(245, ($netTop + 301))
$chkRememberCred.AutoSize = $true
$networkPanel.Controls.Add($chkRememberCred)

$btnTestFound = New-Object System.Windows.Forms.Button
$btnTestFound.Text = 'Test Ticked PCs (changes nothing)'
$btnTestFound.Location = New-Object System.Drawing.Point(590, ($netTop + 296))
$btnTestFound.Size = New-Object System.Drawing.Size(260, 30)
$networkPanel.Controls.Add($btnTestFound)

$btnPushToAll = New-Object System.Windows.Forms.Button
$btnPushToAll.Text = 'Push Current Grid To Ticked PCs'
$btnPushToAll.Location = New-Object System.Drawing.Point(865, ($netTop + 293))
$btnPushToAll.Size = New-Object System.Drawing.Size(260, 36)
$btnPushToAll.BackColor = [System.Drawing.Color]::Khaki
$networkPanel.Controls.Add($btnPushToAll)

$lblManual = New-Object System.Windows.Forms.Label
$lblManual.Text = 'PC not in the list? Add it by name or IP:'
$lblManual.Location = New-Object System.Drawing.Point(15, ($netTop + 340))
$lblManual.AutoSize = $true
$networkPanel.Controls.Add($lblManual)

$txtManualPc = New-Object System.Windows.Forms.TextBox
$txtManualPc.Location = New-Object System.Drawing.Point(255, ($netTop + 336))
$txtManualPc.Size = New-Object System.Drawing.Size(220, 24)
$networkPanel.Controls.Add($txtManualPc)

$chkPushLock = New-Object System.Windows.Forms.CheckBox
$chkPushLock.Text = 'Also apply the Desktop Lock tab settings'
$chkPushLock.Location = New-Object System.Drawing.Point(640, ($netTop + 338))
$chkPushLock.AutoSize = $true
$networkPanel.Controls.Add($chkPushLock)

$chkPushPicture = New-Object System.Windows.Forms.CheckBox
$chkPushPicture.Text = 'Also set the fixed lock-screen picture (Desktop Lock tab)'
$chkPushPicture.Location = New-Object System.Drawing.Point(900, ($netTop + 338))
$chkPushPicture.AutoSize = $true
$networkPanel.Controls.Add($chkPushPicture)

$btnAddManual = New-Object System.Windows.Forms.Button
$btnAddManual.Text = 'Add To List'
$btnAddManual.Location = New-Object System.Drawing.Point(485, ($netTop + 333))
$btnAddManual.Size = New-Object System.Drawing.Size(110, 28)
$networkPanel.Controls.Add($btnAddManual)

$lblNetReport = New-Object System.Windows.Forms.Label
$lblNetReport.Text = 'Report:'
$lblNetReport.Location = New-Object System.Drawing.Point(15, ($netTop + 420))
$lblNetReport.AutoSize = $true
$lblNetReport.ForeColor = [System.Drawing.Color]::DimGray
$networkPanel.Controls.Add($lblNetReport)

$txtNetReport = New-Object System.Windows.Forms.TextBox
$txtNetReport.Text = '(Nothing yet - use the buttons above.)'
$txtNetReport.Location = New-Object System.Drawing.Point(15, ($netTop + 440))
$txtNetReport.Size = New-Object System.Drawing.Size(1290, 200)
$txtNetReport.Multiline = $true
$txtNetReport.ReadOnly = $true
$txtNetReport.BackColor = [System.Drawing.Color]::White
$txtNetReport.ScrollBars = 'Vertical'
$txtNetReport.Font = New-Object System.Drawing.Font('Consolas', 11)
$networkPanel.Controls.Add($txtNetReport)

$lblNetBanner = New-Object System.Windows.Forms.Label
$lblNetBanner.Location = New-Object System.Drawing.Point(15, ($netTop + 376))
$lblNetBanner.Size = New-Object System.Drawing.Size(1290, 36)
$lblNetBanner.Font = New-Object System.Drawing.Font('Segoe UI', 15, [System.Drawing.FontStyle]::Bold)
$lblNetBanner.TextAlign = 'MiddleLeft'
$lblNetBanner.Visible = $false
$networkPanel.Controls.Add($lblNetBanner)
$txtNetReport.Add_TextChanged({ Update-StatusBanner $lblNetBanner $txtNetReport.Text })
$btnCopyNetReport = New-CopyReportButton $networkPanel $txtNetReport 1175 ($netTop + 414)

$btnEnablePush.Add_Click({
    $confirm = [System.Windows.Forms.MessageBox]::Show(
        "This turns on Windows Remote Management (WinRM) on THIS PC ($($env:COMPUTERNAME)) so other machines can push schedule changes to it, and lets local admin accounts authenticate remotely with full rights. Run this once per PC. Continue?",
        'Confirm', 'YesNo', 'Question')
    if ($confirm -ne 'Yes') { return }
    $result = Enable-ThisPCForPushes
    $txtNetReport.Text = "=== Enable This PC To Receive Pushes ($(Get-Date -Format 'yyyy-MM-dd HH:mm')) ===`n$result"
    Write-DebugLog "Enable-ThisPCForPushes:`n$result"
})

# Saved credential: encrypted by Windows (DPAPI) so only THIS Windows user on THIS PC can read it,
# and kept outside the repo folder (%LOCALAPPDATA%), so it is never committed or copied around.
$CredentialPath = Join-Path $env:LOCALAPPDATA 'StudentTimes\PushCredential.xml'
$script:loadingCred = $false

function Save-PushCredential {
    try {
        $dir = Split-Path $CredentialPath
        if (-not (Test-Path $dir)) { New-Item -ItemType Directory -Path $dir -Force | Out-Null }
        $script:pushCredential | Export-Clixml -Path $CredentialPath
        return $true
    } catch { Write-DebugLog "Save-PushCredential failed: $($_.Exception.Message)"; return $false }
}

$btnSetCredential.Add_Click({
    $cred = Get-Credential -Message "Enter the shared admin account used on these PCs (e.g. .\AdminAccountName)"
    if ($cred) {
        $script:pushCredential = $cred
        $lblCredStatus.Text = "Credential: $($cred.UserName)"
        $lblCredStatus.ForeColor = [System.Drawing.Color]::DarkGreen
        if ($chkRememberCred.Checked) {
            if (-not (Save-PushCredential)) { [System.Windows.Forms.MessageBox]::Show('Could not save the credential on this PC (see the debug log). It is still set for this session.', 'Not saved') | Out-Null }
        }
    }
})

$chkRememberCred.Add_CheckedChanged({
    if ($script:loadingCred) { return }
    if ($chkRememberCred.Checked) {
        if ($script:pushCredential) { [void](Save-PushCredential) }   # no credential yet: saved when you set one
    } else {
        Remove-Item -Path $CredentialPath -Force -ErrorAction SilentlyContinue   # unticking = forget it
    }
})

# load a saved one at startup
try {
    if (Test-Path $CredentialPath) {
        $saved = Import-Clixml -Path $CredentialPath
        if ($saved -is [System.Management.Automation.PSCredential]) {
            $script:pushCredential = $saved
            $script:loadingCred = $true; $chkRememberCred.Checked = $true; $script:loadingCred = $false
            $lblCredStatus.Text = "Saved: $($saved.UserName)"
            $lblCredStatus.ForeColor = [System.Drawing.Color]::DarkGreen
        }
    }
} catch { Write-DebugLog "Saved credential could not be loaded (will ask again): $($_.Exception.Message)" }

$script:foundPcs  = @()   # everything the last scan found
$script:manualPcs = @()   # names/IPs typed in by hand (remembered in PushTargets.txt)
$script:shownPcs  = @()   # what the list is showing right now (the rows the checkboxes refer to)
$script:tickedKeys = @{}  # remembered ticks, so re-filtering the list doesn't lose them

foreach ($m in (ConvertFrom-PushTargetText (Get-PushTargetText))) {
    $script:manualPcs += [PSCustomObject]@{ Ip = $m; Name = ''; Ready = $false; IsWindows = $true; IsOther = $false; Status = 'added by hand - use Test to check it'; IsThisPc = $false; Manual = $true }
    $script:tickedKeys[$m] = $true
}

function Save-ManualPcs { Save-PushTargetText (($script:manualPcs | ForEach-Object { $_.Ip }) -join "`r`n") }

function Remember-Ticks {
    for ($i = 0; $i -lt $script:shownPcs.Count; $i++) {
        $script:tickedKeys[$script:shownPcs[$i].Ip] = $lstFound.GetItemChecked($i)
    }
}

function Show-FoundPcs {
    $lstFound.Items.Clear()
    $all = @($script:foundPcs)
    foreach ($m in $script:manualPcs) {
        if (-not (@($all | Where-Object { $_.Ip -eq $m.Ip -or ($_.Name -and $_.Name -eq $m.Ip) }).Count)) { $all += $m }
    }
    $script:shownPcs = @($all | Where-Object { $chkShowOther.Checked -or -not $_.IsOther })
    foreach ($pc in $script:shownPcs) {
        $label = '{0,-18} {1,-15} {2}' -f $(if ($pc.Name) { $pc.Name } else { '(name unknown)' }), $pc.Ip, $pc.Status
        $idx = $lstFound.Items.Add($label)
        $tick = if ($script:tickedKeys.ContainsKey($pc.Ip)) { [bool]$script:tickedKeys[$pc.Ip] } else { [bool]($pc.Ready -and -not $pc.IsThisPc) }
        if ($tick) { $lstFound.SetItemChecked($idx, $true) }
    }
}
$chkShowOther.Add_CheckedChanged({ Remember-Ticks; Show-FoundPcs })

function Get-TickedPcs {
    $picked = @()
    foreach ($i in $lstFound.CheckedIndices) {
        $pc = $script:shownPcs[$i]
        if ($pc -and -not $pc.IsThisPc) { $picked += $pc }
    }
    return $picked
}

Show-FoundPcs

$btnScanNet.Add_Click({
    Remember-Ticks
    $btnScanNet.Enabled = $false
    $lblScanStatus.ForeColor = [System.Drawing.Color]::DimGray
    $lblScanStatus.Text = 'Scanning this network... (takes about 10-15 seconds)'
    $lstFound.Items.Clear()
    $form.Cursor = [System.Windows.Forms.Cursors]::WaitCursor
    [System.Windows.Forms.Application]::DoEvents()
    try {
        $script:foundPcs = @(Find-LanPcs)
        $script:tickedKeys = @{}   # a fresh scan re-ticks the ready ones
        foreach ($m in $script:manualPcs) { $script:tickedKeys[$m.Ip] = $true }
        Show-FoundPcs
        $ready = @($script:foundPcs | Where-Object { $_.Ready }).Count
        $win   = @($script:foundPcs | Where-Object { $_.IsWindows -and -not $_.IsThisPc }).Count
        $other = @($script:foundPcs | Where-Object { $_.IsOther }).Count
        $lblScanStatus.Text = "Found $win other Windows PC(s): $ready ready, $($win - $ready) need Step 1. ($other other device(s) hidden.)"
        $lblScanStatus.ForeColor = if ($ready -gt 0) { [System.Drawing.Color]::DarkGreen } else { [System.Drawing.Color]::DarkRed }
        Write-DebugLog ("Scan Network: " + (($script:foundPcs | ForEach-Object { "$($_.Ip)=$($_.Name) [$($_.Status)]" }) -join '; '))
    } catch {
        $lblScanStatus.Text = "Scan failed: $($_.Exception.Message)"
        $lblScanStatus.ForeColor = [System.Drawing.Color]::DarkRed
    } finally {
        $form.Cursor = [System.Windows.Forms.Cursors]::Default
        $btnScanNet.Enabled = $true
    }
})

$btnAddManual.Add_Click({
    $entry = ($txtManualPc.Text -replace '\s*#.*$', '').Trim()
    if (-not $entry) { return }
    if (-not (@($script:manualPcs | Where-Object { $_.Ip -eq $entry }).Count)) {
        $script:manualPcs += [PSCustomObject]@{ Ip = $entry; Name = ''; Ready = $false; IsWindows = $true; IsOther = $false; Status = 'added by hand - use Test to check it'; IsThisPc = $false; Manual = $true }
        Save-ManualPcs
    }
    Remember-Ticks
    $script:tickedKeys[$entry] = $true
    $txtManualPc.Text = ''
    Show-FoundPcs
})

$btnTestFound.Add_Click({
    if (-not $script:pushCredential) {
        [System.Windows.Forms.MessageBox]::Show('Set the shared admin credential first (the test logs in to each PC with it).', 'No credential') | Out-Null
        return
    }
    $picked = @(Get-TickedPcs)
    if ($picked.Count -eq 0) {
        [System.Windows.Forms.MessageBox]::Show('Tick at least one PC in the list first.', 'Nothing ticked') | Out-Null
        return
    }
    $form.Cursor = [System.Windows.Forms.Cursors]::WaitCursor
    [System.Windows.Forms.Application]::DoEvents()
    $report = New-Object System.Text.StringBuilder
    [void]$report.AppendLine("=== Can each PC receive a push? ($(Get-Date -Format 'yyyy-MM-dd HH:mm')) - logs in with $($script:pushCredential.UserName), changes nothing ===")
    foreach ($pc in $picked) {
        $t = Test-PushReadiness -target $pc.Ip -credential $script:pushCredential
        $tag = if ($t.Ok) { 'READY ' } else { 'NOT READY' }
        $who = if ($t.Name) { $t.Name } elseif ($pc.Name) { $pc.Name } else { '?' }
        [void]$report.AppendLine("[$tag] $($pc.Ip)  ($who) -> $($t.Text)")
        Write-DebugLog "Test-PushReadiness '$($pc.Ip)': ok=$($t.Ok) name=$who text=$($t.Text)"
    }
    $txtNetReport.Text = $report.ToString()
    $form.Cursor = [System.Windows.Forms.Cursors]::Default
})

$btnPushToAll.Add_Click({
    $targets = @(Get-TargetAccounts)
    if ($targets.Count -eq 0) {
        [System.Windows.Forms.MessageBox]::Show('Enter a Windows account name (or pick accounts) on the Daily Use tab first.', 'Missing account') | Out-Null
        return
    }
    $userName = $targets -join ', '
    $picked = @(Get-TickedPcs)
    if ($picked.Count -eq 0) {
        [System.Windows.Forms.MessageBox]::Show('Tick at least one PC in the list first (use Scan Network For PCs, or add one by name/IP).', 'No PCs ticked') | Out-Null
        return
    }
    if (-not $script:pushCredential) {
        [System.Windows.Forms.MessageBox]::Show('Set the shared admin credential first.', 'No credential') | Out-Null
        return
    }
    $pushPictureOpts = $null
    if ($chkPushPicture.Checked) {
        $pushPictureOpts = Get-LockPictureOptions $txtLockPicture.Text.Trim()
        if ($null -eq $pushPictureOpts) { [System.Windows.Forms.MessageBox]::Show('You ticked "set the fixed lock-screen picture" but no valid picture is chosen on the Desktop Lock tab.', 'No picture') | Out-Null; return }
        if (-not $pushPictureOpts.WindowsDefault -and @($pushPictureOpts.ImageBytes).Count -gt 350000) { [System.Windows.Forms.MessageBox]::Show('That picture is too big to push (350 KB max over remote management). Use a smaller jpg, "Make a plain Rubin IT picture", or "Use the Windows default picture".', 'Picture too big') | Out-Null; return }
    }
    $timesValue = Build-TimesString
    if (-not $timesValue) {
        # An empty /times: means "may NEVER log on" - pushing that locks everyone out of every target PC.
        [System.Windows.Forms.MessageBox]::Show("The grid on the Daily Use tab is EMPTY (no hours allowed). Pushing that would lock '$userName' out of every ticked PC completely.`n`nGo to the Daily Use tab, click 'Load Current Schedule From This PC' (or paint the hours you want), then push again.", 'Empty schedule - not pushed') | Out-Null
        return
    }
    $command = ($targets | ForEach-Object { "net user $_ /times:$timesValue" }) -join "`n"
    $allowedHours = 0; for ($d = 0; $d -lt 7; $d++) { for ($h = 0; $h -lt 24; $h++) { if ($Grid[$d,$h]) { $allowedHours++ } } }
    $names = @($picked | ForEach-Object { if ($_.Name) { "$($_.Name)  ($($_.Ip))" } else { $_.Ip } })
    $confirm = [System.Windows.Forms.MessageBox]::Show(
        "This pushes the grid EXACTLY as shown on the Daily Use tab ($allowedHours of 168 hours per week allowed) by running:`n`n$command`n`non these $($picked.Count) PC(s):`n`n$($names -join "`n")`n`nContinue?",
        'Confirm push', 'YesNo', 'Question')
    if ($confirm -ne 'Yes') { return }

    $report = New-Object System.Text.StringBuilder
    [void]$report.AppendLine("=== Push '$command' ($(Get-Date -Format 'yyyy-MM-dd HH:mm')) ===")
    foreach ($pc in $picked) {
        foreach ($acct in $targets) {
            $result = Push-TimesToMachine -targetName $pc.Ip -credential $script:pushCredential -userName $acct -timesValue $timesValue -groups @(Get-LogoutKickoffGroups $Grid) -action $(if ($cmbLogoutAction.SelectedItem -eq 'Full log off') { 'LogOff' } else { 'Lock' }) -lockOpts $(if ($chkPushLock.Checked) { Get-LockOptions } else { $null }) -pictureOpts $(if ($chkPushPicture.Checked -and $acct -eq $targets[0]) { $pushPictureOpts } else { $null })
            $status = if ($result.Success) { 'OK' } else { 'FAILED' }
            $label = if ($pc.Name) { "$($pc.Name) ($($pc.Ip))" } else { $pc.Ip }
            [void]$report.AppendLine("[$status] $label  account '$acct' -> $($result.Output)")
            Write-DebugLog "Push-TimesToMachine '$($pc.Ip)' '$acct': success=$($result.Success) output=$($result.Output)"
        }
    }
    $txtNetReport.Text = $report.ToString()
    Write-LogEntry -userName $userName -oldTimes '' -newTimes $command -note "Pushed to: $(($picked | ForEach-Object { $_.Ip }) -join ', ')"
})

# ---------------------------------------------------------------------------
# Desktop Lock tab: block changing the desktop + Windows Spotlight off (per student account).
# ---------------------------------------------------------------------------
$lockTop = 20

$lblLockHead = New-Object System.Windows.Forms.Label
$lblLockHead.Text = "Stop the account in the 'Windows account name' box (Daily Use tab) from changing the desktop, and turn off Windows Spotlight:"
$lblLockHead.Location = New-Object System.Drawing.Point(15, $lockTop)
$lblLockHead.Size = New-Object System.Drawing.Size(1290, 18)
$lockPanel.Controls.Add($lblLockHead)

$chkLockWallpaper = New-Object System.Windows.Forms.CheckBox
$chkLockWallpaper.Text = 'Block changing the desktop background, theme and colors'
$chkLockWallpaper.Checked = $true
$chkLockWallpaper.Location = New-Object System.Drawing.Point(15, ($lockTop + 30))
$chkLockWallpaper.AutoSize = $true
$lockPanel.Controls.Add($chkLockWallpaper)

$chkLockSpotlight = New-Object System.Windows.Forms.CheckBox
$chkLockSpotlight.Text = 'Turn off Windows Spotlight (rotating lock-screen / desktop pictures, "Learn about this picture", tips, suggestions, auto-installed promo apps)'
$chkLockSpotlight.Checked = $true
$chkLockSpotlight.Location = New-Object System.Drawing.Point(15, ($lockTop + 56))
$chkLockSpotlight.AutoSize = $true
$lockPanel.Controls.Add($chkLockSpotlight)

$chkLockScreenMachine = New-Object System.Windows.Forms.CheckBox
$chkLockScreenMachine.Text = 'Also block changing the lock-screen picture for EVERYONE on this PC (machine-wide, admins too)'
$chkLockScreenMachine.Checked = $false
$chkLockScreenMachine.Location = New-Object System.Drawing.Point(15, ($lockTop + 82))
$chkLockScreenMachine.AutoSize = $true
$lockPanel.Controls.Add($chkLockScreenMachine)

$btnLockCheck = New-Object System.Windows.Forms.Button
$btnLockCheck.Text = 'Check'
$btnLockCheck.Location = New-Object System.Drawing.Point(15, ($lockTop + 116))
$btnLockCheck.Size = New-Object System.Drawing.Size(150, 30)
$lockPanel.Controls.Add($btnLockCheck)

$btnLockApply = New-Object System.Windows.Forms.Button
$btnLockApply.Text = 'Apply Lock-Down'
$btnLockApply.Location = New-Object System.Drawing.Point(175, ($lockTop + 116))
$btnLockApply.Size = New-Object System.Drawing.Size(170, 30)
$btnLockApply.BackColor = [System.Drawing.Color]::Khaki
$lockPanel.Controls.Add($btnLockApply)

$btnLockUndo = New-Object System.Windows.Forms.Button
$btnLockUndo.Text = 'Undo (remove these settings)'
$btnLockUndo.Location = New-Object System.Drawing.Point(355, ($lockTop + 116))
$btnLockUndo.Size = New-Object System.Drawing.Size(200, 30)
$lockPanel.Controls.Add($btnLockUndo)

$lblLockNote = New-Object System.Windows.Forms.Label
$lblLockNote.Text = 'Takes effect at the next sign-in of that account. Works even if the account is not signed in. The account must have signed in at least once.'
$lblLockNote.Location = New-Object System.Drawing.Point(15, ($lockTop + 152))
$lblLockNote.Size = New-Object System.Drawing.Size(1290, 18)
$lblLockNote.ForeColor = [System.Drawing.Color]::DimGray
$lockPanel.Controls.Add($lblLockNote)

$lblLockBanner = New-Object System.Windows.Forms.Label
$lblLockBanner.Location = New-Object System.Drawing.Point(15, ($lockTop + 180))
$lblLockBanner.Size = New-Object System.Drawing.Size(1290, 36)
$lblLockBanner.Font = New-Object System.Drawing.Font('Segoe UI', 15, [System.Drawing.FontStyle]::Bold)
$lblLockBanner.TextAlign = 'MiddleLeft'
$lblLockBanner.Visible = $false
$lockPanel.Controls.Add($lblLockBanner)

$txtLockReport = New-Object System.Windows.Forms.TextBox
$txtLockReport.Text = '(Nothing yet - use the buttons above.)'
$txtLockReport.Location = New-Object System.Drawing.Point(15, ($lockTop + 222))
$txtLockReport.Size = New-Object System.Drawing.Size(1290, 150)
$txtLockReport.Multiline = $true
$txtLockReport.ReadOnly = $true
$txtLockReport.BackColor = [System.Drawing.Color]::White
$txtLockReport.ScrollBars = 'Vertical'
$txtLockReport.Font = New-Object System.Drawing.Font('Consolas', 11)
$lockPanel.Controls.Add($txtLockReport)
$txtLockReport.Add_TextChanged({ Update-StatusBanner $lblLockBanner $txtLockReport.Text })
$btnCopyLockReport = New-CopyReportButton $lockPanel $txtLockReport 1175 ($lockTop + 150)

function Get-LockOptions {
    return @{ Wallpaper = $chkLockWallpaper.Checked; Spotlight = $chkLockSpotlight.Checked; LockScreenMachine = $chkLockScreenMachine.Checked }
}

function Invoke-LockDown([string]$mode) {
    $targets = @(Get-TargetAccounts)
    if ($targets.Count -eq 0) { [System.Windows.Forms.MessageBox]::Show('Enter a Windows account name (or pick accounts) first.', 'Missing account') | Out-Null; return }
    $lines = @()
    foreach ($acct in $targets) {
        $lines += "-- $acct --"
        $lines += @(& $DesktopLockBlock $acct $mode (Get-LockOptions))
    }
    Write-DebugLog "Desktop lock ($mode) $($targets -join ','):`n$($lines -join "`n")"
    $title = switch ($mode) { 'Apply' { 'Lock-down applied' } 'Undo' { 'Lock-down removed' } default { 'Lock-down check' } }
    $txtLockReport.Text = "=== $title on $($env:COMPUTERNAME) ($(Get-Date -Format 'yyyy-MM-dd HH:mm')) ===`r`n" + ($lines -join "`r`n")
}
$btnLockCheck.Add_Click({ Invoke-LockDown 'Check' })
$btnLockApply.Add_Click({ Invoke-LockDown 'Apply' })
$btnLockUndo.Add_Click({ Invoke-LockDown 'Undo' })

# --- Fixed lock-screen / sign-in picture (whole PC) ---
$picTop = $lockTop + 395

$lblPicHead = New-Object System.Windows.Forms.Label
$lblPicHead.Text = 'Fixed lock-screen / sign-in picture (the picture you see BEFORE choosing a user) - the same on every user, and nobody can change it:'
$lblPicHead.Location = New-Object System.Drawing.Point(15, $picTop)
$lblPicHead.Size = New-Object System.Drawing.Size(1290, 18)
$lblPicHead.ForeColor = [System.Drawing.Color]::DarkSlateBlue
$lockPanel.Controls.Add($lblPicHead)

$txtLockPicture = New-Object System.Windows.Forms.TextBox
$txtLockPicture.Location = New-Object System.Drawing.Point(15, ($picTop + 26))
$txtLockPicture.Size = New-Object System.Drawing.Size(640, 24)
$lockPanel.Controls.Add($txtLockPicture)

$btnBrowsePicture = New-Object System.Windows.Forms.Button
$btnBrowsePicture.Text = 'Browse for a picture...'
$btnBrowsePicture.Location = New-Object System.Drawing.Point(665, ($picTop + 23))
$btnBrowsePicture.Size = New-Object System.Drawing.Size(160, 28)
$lockPanel.Controls.Add($btnBrowsePicture)

$btnPlainPicture = New-Object System.Windows.Forms.Button
$btnPlainPicture.Text = 'Make a plain Rubin IT picture'
$btnPlainPicture.Location = New-Object System.Drawing.Point(835, ($picTop + 23))
$btnPlainPicture.Size = New-Object System.Drawing.Size(210, 28)
$lockPanel.Controls.Add($btnPlainPicture)

$btnWinDefaultPicture = New-Object System.Windows.Forms.Button
$btnWinDefaultPicture.Text = 'Use the Windows default picture'
$btnWinDefaultPicture.Location = New-Object System.Drawing.Point(1055, ($picTop + 23))
$btnWinDefaultPicture.Size = New-Object System.Drawing.Size(230, 28)
$lockPanel.Controls.Add($btnWinDefaultPicture)
$btnWinDefaultPicture.Add_Click({ $txtLockPicture.Text = 'Windows default picture (Windows'' own default lock screen)' })

$btnPicCheck = New-Object System.Windows.Forms.Button
$btnPicCheck.Text = 'Check Picture'
$btnPicCheck.Location = New-Object System.Drawing.Point(15, ($picTop + 60))
$btnPicCheck.Size = New-Object System.Drawing.Size(150, 30)
$lockPanel.Controls.Add($btnPicCheck)

$btnPicApply = New-Object System.Windows.Forms.Button
$btnPicApply.Text = 'Set + Lock This Picture'
$btnPicApply.Location = New-Object System.Drawing.Point(175, ($picTop + 60))
$btnPicApply.Size = New-Object System.Drawing.Size(190, 30)
$btnPicApply.BackColor = [System.Drawing.Color]::Khaki
$lockPanel.Controls.Add($btnPicApply)

$btnPicUndo = New-Object System.Windows.Forms.Button
$btnPicUndo.Text = 'Remove (back to normal)'
$btnPicUndo.Location = New-Object System.Drawing.Point(375, ($picTop + 60))
$btnPicUndo.Size = New-Object System.Drawing.Size(190, 30)
$lockPanel.Controls.Add($btnPicUndo)

$lblPicNote = New-Object System.Windows.Forms.Label
$lblPicNote.Text = 'Applies to the whole PC (not one account). Results appear in the report above. Check it by SIGNING OUT (a plain lock shows the signed-in user''s own picture).'
$lblPicNote.Location = New-Object System.Drawing.Point(15, ($picTop + 96))
$lblPicNote.Size = New-Object System.Drawing.Size(1290, 18)
$lblPicNote.ForeColor = [System.Drawing.Color]::DimGray
$lockPanel.Controls.Add($lblPicNote)

$btnBrowsePicture.Add_Click({
    $dlg = New-Object System.Windows.Forms.OpenFileDialog
    $dlg.Filter = 'Pictures (*.jpg;*.jpeg;*.png;*.bmp)|*.jpg;*.jpeg;*.png;*.bmp'
    $dlg.Title = 'Choose the fixed lock-screen picture'
    if ($dlg.ShowDialog() -eq 'OK') { $txtLockPicture.Text = $dlg.FileName }
})
$btnPlainPicture.Add_Click({
    try { $txtLockPicture.Text = New-PlainLockPicture } catch { [System.Windows.Forms.MessageBox]::Show("Could not create the picture: $($_.Exception.Message)", 'Failed') | Out-Null }
})

function Invoke-LockPicture([string]$mode) {
    $opts = $null
    if ($mode -eq 'Apply') {
        $opts = Get-LockPictureOptions $txtLockPicture.Text.Trim()
        if ($null -eq $opts) { [System.Windows.Forms.MessageBox]::Show('Choose a picture first (jpg, png or bmp), or click one of the default-picture buttons.', 'No picture') | Out-Null; return }
        $c = [System.Windows.Forms.MessageBox]::Show("Set this picture as the lock-screen / sign-in picture for EVERY user of $($env:COMPUTERNAME), and block changing it?`n`n$($txtLockPicture.Text.Trim())", 'Confirm', 'YesNo', 'Question')
        if ($c -ne 'Yes') { return }
    }
    $lines = @(& $LockScreenPictureBlock $mode $opts)
    Write-DebugLog "Lock-screen picture ($mode):`n$($lines -join "`n")"
    $title = switch ($mode) { 'Apply' { 'Fixed lock-screen picture set' } 'Undo' { 'Fixed lock-screen picture removed' } default { 'Fixed lock-screen picture check' } }
    $txtLockReport.Text = "=== $title on $($env:COMPUTERNAME) ($(Get-Date -Format 'yyyy-MM-dd HH:mm')) ===`r`n" + ($lines -join "`r`n")
}
$btnPicCheck.Add_Click({ Invoke-LockPicture 'Check' })
$btnPicApply.Add_Click({ Invoke-LockPicture 'Apply' })
$btnPicUndo.Add_Click({ Invoke-LockPicture 'Undo' })


# ---------------------------------------------------------------------------
# Which accounts do the actions apply to?  Either the one typed in the box, or the accounts ticked in the
# "Accounts on this PC" picker (e.g. user1..user6).  Every action that sets hours / tasks / lock-down /
# unlock / push loops over Get-TargetAccounts.
# ---------------------------------------------------------------------------
$script:selectedAccounts = @()

function Get-TargetAccounts {
    if (@($script:selectedAccounts).Count -gt 0) { return @($script:selectedAccounts) }
    $u = $txtUser.Text.Trim()
    if ($u) { return @($u) }
    return @()
}

function Get-LocalAccountCandidates {
    # Enabled local accounts except the Windows built-ins; flags administrators and the account running this tool.
    $adminNames = @()
    try { $adminNames = @(Get-LocalGroupMember -SID 'S-1-5-32-544' -ErrorAction Stop | ForEach-Object { ($_.Name -split '\\')[-1] }) } catch { }
    $builtin = @('Administrator', 'DefaultAccount', 'Guest', 'WDAGUtilityAccount')
    $list = @()
    try {
        foreach ($u in @(Get-LocalUser -ErrorAction Stop)) {
            if (-not $u.Enabled) { continue }
            if ($builtin -contains $u.Name) { continue }
            $list += [pscustomobject]@{ Name = $u.Name; IsAdmin = ($adminNames -icontains $u.Name); IsMe = ($u.Name -ieq $env:USERNAME) }
        }
    } catch { Write-DebugLog "Get-LocalAccountCandidates failed: $($_.Exception.Message)" }
    return @($list | Sort-Object Name)
}

function Update-ApplyToLabel {
    $t = @(Get-TargetAccounts)
    if (@($script:selectedAccounts).Count -gt 0) {
        $lblApplyTo.Text = "Applying to $($t.Count) account(s): $($t -join ', ')"
        $lblApplyTo.ForeColor = [System.Drawing.Color]::DarkGreen
    } else {
        $lblApplyTo.Text = 'Applying to the account typed on the left (click "Accounts on this PC..." to tick several, e.g. user1-user6)'
        $lblApplyTo.ForeColor = [System.Drawing.Color]::DimGray
    }
}

function Show-AccountPicker {
    $dlg = New-Object System.Windows.Forms.Form
    $dlg.Text = 'Accounts on this PC'
    $dlg.StartPosition = 'CenterParent'
    $dlg.Size = New-Object System.Drawing.Size(460, 470)
    $dlg.FormBorderStyle = 'FixedDialog'; $dlg.MaximizeBox = $false; $dlg.MinimizeBox = $false

    $lbl = New-Object System.Windows.Forms.Label
    $lbl.Text = 'Tick the student accounts to apply the next actions to:'
    $lbl.Location = New-Object System.Drawing.Point(12, 10); $lbl.AutoSize = $true
    $dlg.Controls.Add($lbl)

    $lst = New-Object System.Windows.Forms.CheckedListBox
    $lst.CheckOnClick = $true
    $lst.Location = New-Object System.Drawing.Point(12, 34); $lst.Size = New-Object System.Drawing.Size(420, 300)
    $dlg.Controls.Add($lst)

    $chkAdmins = New-Object System.Windows.Forms.CheckBox
    $chkAdmins.Text = 'Also show administrators and the account I am signed in with'
    $chkAdmins.Location = New-Object System.Drawing.Point(12, 340); $chkAdmins.AutoSize = $true
    $dlg.Controls.Add($chkAdmins)

    $note = New-Object System.Windows.Forms.Label
    $note.Text = 'Pre-ticked: accounts named student(s), stu, bochurim or userN. Administrators are never pre-ticked.'
    $note.Location = New-Object System.Drawing.Point(12, 366); $note.Size = New-Object System.Drawing.Size(420, 32)
    $note.ForeColor = [System.Drawing.Color]::DimGray
    $dlg.Controls.Add($note)

    $btnOk = New-Object System.Windows.Forms.Button
    $btnOk.Text = 'Use the ticked accounts'; $btnOk.Location = New-Object System.Drawing.Point(12, 400); $btnOk.Size = New-Object System.Drawing.Size(170, 28)
    $btnOk.DialogResult = 'OK'
    $dlg.Controls.Add($btnOk); $dlg.AcceptButton = $btnOk

    $btnNone = New-Object System.Windows.Forms.Button
    $btnNone.Text = 'Clear (use typed name)'; $btnNone.Location = New-Object System.Drawing.Point(190, 400); $btnNone.Size = New-Object System.Drawing.Size(140, 28)
    $dlg.Controls.Add($btnNone)

    $btnCancel = New-Object System.Windows.Forms.Button
    $btnCancel.Text = 'Cancel'; $btnCancel.Location = New-Object System.Drawing.Point(338, 400); $btnCancel.Size = New-Object System.Drawing.Size(94, 28)
    $btnCancel.DialogResult = 'Cancel'
    $dlg.Controls.Add($btnCancel); $dlg.CancelButton = $btnCancel

    $all = @(Get-LocalAccountCandidates)
    $ticked = @{}
    foreach ($n in @($script:selectedAccounts)) { $ticked[$n.ToLower()] = $true }
    $fill = {
        for ($i = 0; $i -lt $lst.Items.Count; $i++) { $ticked[$script:pickerRows[$i].Name.ToLower()] = $lst.GetItemChecked($i) }
        $lst.Items.Clear()
        $script:pickerRows = @($all | Where-Object { $chkAdmins.Checked -or (-not $_.IsAdmin -and -not $_.IsMe) })
        foreach ($r in $script:pickerRows) {
            $tag = ''
            if ($r.IsAdmin) { $tag += '  (administrator)' }
            if ($r.IsMe) { $tag += '  (you are signed in as this)' }
            $idx = $lst.Items.Add($r.Name + $tag)
            $key = $r.Name.ToLower()
            $on = $false
            if ($ticked.ContainsKey($key)) { $on = [bool]$ticked[$key] }
            elseif (@($script:selectedAccounts).Count -eq 0 -and -not $r.IsAdmin -and -not $r.IsMe -and $r.Name -match '^(?i)(students?|stu|bochurim|user\d+)$') { $on = $true }
            if ($on) { $lst.SetItemChecked($idx, $true) }
        }
    }
    $script:pickerRows = @()
    $chkAdmins.Add_CheckedChanged($fill)
    & $fill
    if ($script:pickerRows.Count -eq 0 -and $all.Count -eq 0) { $lbl.Text = 'No local accounts could be listed on this PC - type the name in the main window instead.' }

    $script:pickerCleared = $false
    $btnNone.Add_Click({ $script:pickerCleared = $true; $dlg.DialogResult = 'OK'; $dlg.Close() })

    if ($dlg.ShowDialog($form) -ne 'OK') { return }
    if ($script:pickerCleared) { $script:selectedAccounts = @(); Update-ApplyToLabel; return }
    $names = @()
    for ($i = 0; $i -lt $script:pickerRows.Count; $i++) { if ($lst.GetItemChecked($i)) { $names += $script:pickerRows[$i].Name } }
    $risky = @($script:pickerRows | Where-Object { ($_.IsAdmin -or $_.IsMe) -and ($names -contains $_.Name) })
    if ($risky.Count -gt 0) {
        $c = [System.Windows.Forms.MessageBox]::Show("You ticked an administrator / your own account ($(($risky | ForEach-Object { $_.Name }) -join ', ')). Restricting its hours or locking its desktop could lock YOU out of this PC.`n`nKeep it ticked anyway?", 'Careful', 'YesNo', 'Warning')
        if ($c -ne 'Yes') { $names = @($names | Where-Object { $risky.Name -notcontains $_ }) }
    }
    $script:selectedAccounts = @($names)
    if ($names.Count -gt 0) { $txtUser.Text = $names[0] }   # the first one is also the one shown/loaded in the grid
    Update-ApplyToLabel
}

$btnAccounts.Add_Click({ Show-AccountPicker })

# ---------------------------------------------------------------------------
# Auto-lock NEW accounts: a SYSTEM scheduled task keeps every NON-ADMIN local account on the chosen hours,
# including accounts created later (user5, user6 ...).  It fires on "a user account was created" (security
# event 4720), at every boot and every 10 minutes (safety net).  For each enabled non-admin account that is
# still completely UNRESTRICTED (logon hours = all) it applies the saved schedule + the weekly logout tasks.
# It never touches: administrators, Windows built-in accounts, accounts on the "never touch" list, accounts
# that already have restricted hours (someone set them on purpose), or an account that is in a temporary
# unlock (pending Relock task).  Files live in C:\ProgramData\StudentTimes\AutoApply (admins/SYSTEM write only).
# ---------------------------------------------------------------------------
$AutoApplyDir = Join-Path $env:ProgramData 'StudentTimes\AutoApply'
$AutoApplyTask = 'StudentTimes-AutoApply'

function Protect-StudentTimesDir {
    # Scripts in here run as SYSTEM, so ordinary users must only be able to READ them.
    $d = Join-Path $env:ProgramData 'StudentTimes'
    if (-not (Test-Path $d)) { New-Item -ItemType Directory -Path $d -Force | Out-Null }
    & icacls.exe $d /inheritance:r /grant:r '*S-1-5-18:(OI)(CI)F' '*S-1-5-32-544:(OI)(CI)F' '*S-1-5-32-545:(OI)(CI)RX' /Q 2>&1 | Out-Null
    return ($LASTEXITCODE -eq 0)
}

function Get-AutoApplyScriptText {
    $syncText = $SyncLogoutTasksBlock.ToString()
    return @"
# Generated by Set-StudentTimes - do not edit by hand (turn it off/on again from the tool instead).
`$ErrorActionPreference = 'Continue'
`$base = Join-Path `$env:ProgramData 'StudentTimes\AutoApply'
`$log = Join-Path `$base 'autoapply.log'
function Log(`$m) { try { Add-Content -Path `$log -Value ('{0:s} {1}' -f (Get-Date), `$m) -Encoding UTF8 } catch { } }
try { if ((Get-Item `$log -ErrorAction SilentlyContinue).Length -gt 400KB) { Move-Item `$log (`$log + '.old') -Force } } catch { }
`$sync = {
$syncText
}
try { `$pol = Get-Content (Join-Path `$base 'policy.json') -Raw -ErrorAction Stop | ConvertFrom-Json } catch { Log "policy.json unreadable: `$(`$_.Exception.Message)"; exit 1 }
if (-not `$pol.Enabled) { exit 0 }
`$builtin = @('Administrator', 'DefaultAccount', 'Guest', 'WDAGUtilityAccount')
`$admins = @()
try { `$admins = @(Get-LocalGroupMember -SID 'S-1-5-32-544' -ErrorAction Stop | ForEach-Object { (`$_.Name -split '\\')[-1] }) } catch { Log "could not list administrators: `$(`$_.Exception.Message) - doing nothing this run"; exit 1 }
`$excluded = @(`$pol.Excluded | Where-Object { `$_ })
foreach (`$u in @(Get-LocalUser -ErrorAction SilentlyContinue | Where-Object { `$_.Enabled })) {
    `$n = `$u.Name
    if (`$builtin -contains `$n -or `$admins -icontains `$n -or `$excluded -icontains `$n) { continue }
    if (@(Get-ScheduledTask -TaskName "StudentTimes-`$n-*-Relock" -ErrorAction SilentlyContinue).Count -gt 0) { continue }
    try { `$b = ([ADSI]"WinNT://./`$n,user").Properties['LoginHours'].Value } catch { Log "cannot read hours of '`$n': `$(`$_.Exception.Message)"; continue }
    `$allOpen = (-not `$b) -or (@(`$b | Where-Object { `$_ -ne 255 }).Count -eq 0)
    if (-not `$allOpen) { continue }
    `$res = (& net.exe user `$n "/times:`$(`$pol.TimesValue)" 2>&1 | ForEach-Object { "`$_" }) -join ' '
    if (`$LASTEXITCODE -ne 0) { Log "FAILED to restrict '`$n': `$res"; continue }
    Log "restricted account '`$n' to `$(`$pol.TimesValue)"
    foreach (`$line in @(& `$sync `$n `$pol.Groups `$pol.Action '')) { Log "  `$n : `$line" }
}
"@
}

function Install-AutoApply([string]$timesValue, $groups, [string]$action, [string[]]$excluded) {
    $out = @()
    try {
        $isAdm = ([Security.Principal.WindowsPrincipal][Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)
        if (-not $isAdm) { return @('[FAIL] This needs administrator rights (run the tool as administrator). Nothing was changed.') }
        if (-not (Protect-StudentTimesDir)) { $out += '[WARN] Could not tighten the permissions of C:\ProgramData\StudentTimes.' }
        if (-not (Test-Path $AutoApplyDir)) { New-Item -ItemType Directory -Path $AutoApplyDir -Force | Out-Null }
        $policy = [pscustomobject]@{
            Enabled = $true; TimesValue = $timesValue; Action = $action
            Groups = @($groups | ForEach-Object { [pscustomobject]@{ TimeHHmm = $_.TimeHHmm; Days = @($_.Days) } })
            Excluded = @($excluded | Where-Object { $_ })
            Updated = (Get-Date).ToString('s'); UpdatedOn = $env:COMPUTERNAME
        }
        ($policy | ConvertTo-Json -Depth 5) | Set-Content -Path (Join-Path $AutoApplyDir 'policy.json') -Encoding UTF8
        (Get-AutoApplyScriptText) | Set-Content -Path (Join-Path $AutoApplyDir 'AutoApply.ps1') -Encoding UTF8
        $scriptPath = Join-Path $AutoApplyDir 'AutoApply.ps1'
        $now = (Get-Date).ToString('s')
        $xml = @"
<?xml version="1.0" encoding="UTF-16"?>
<Task version="1.2" xmlns="http://schemas.microsoft.com/windows/2004/02/mit/task">
  <RegistrationInfo><Description>Locks every non-admin local account (including new ones) to the Student Times schedule.</Description></RegistrationInfo>
  <Triggers>
    <EventTrigger>
      <Enabled>true</Enabled>
      <Subscription>&lt;QueryList&gt;&lt;Query Id="0" Path="Security"&gt;&lt;Select Path="Security"&gt;*[System[(EventID=4720)]]&lt;/Select&gt;&lt;/Query&gt;&lt;/QueryList&gt;</Subscription>
    </EventTrigger>
    <BootTrigger><Enabled>true</Enabled></BootTrigger>
    <TimeTrigger>
      <StartBoundary>$now</StartBoundary>
      <Enabled>true</Enabled>
      <Repetition><Interval>PT10M</Interval></Repetition>
    </TimeTrigger>
  </Triggers>
  <Principals><Principal id="Author"><UserId>S-1-5-18</UserId><RunLevel>HighestAvailable</RunLevel></Principal></Principals>
  <Settings>
    <MultipleInstancesPolicy>IgnoreNew</MultipleInstancesPolicy>
    <DisallowStartIfOnBatteries>false</DisallowStartIfOnBatteries>
    <StopIfGoingOnBatteries>false</StopIfGoingOnBatteries>
    <StartWhenAvailable>true</StartWhenAvailable>
    <ExecutionTimeLimit>PT5M</ExecutionTimeLimit>
    <Hidden>true</Hidden>
  </Settings>
  <Actions Context="Author"><Exec><Command>powershell.exe</Command><Arguments>-NoProfile -ExecutionPolicy Bypass -File "$scriptPath"</Arguments></Exec></Actions>
</Task>
"@
        Register-ScheduledTask -TaskName $AutoApplyTask -Xml $xml -Force -ErrorAction Stop | Out-Null
        $out += "[OK] Auto-lock is ON: task '$AutoApplyTask' runs when an account is created, at boot and every 10 minutes."
        $out += "[OK] Saved schedule: net user <account> /times:$timesValue"
        if (@($excluded | Where-Object { $_ }).Count -gt 0) { $out += "[OK] Never touched: $(($excluded | Where-Object { $_ }) -join ', ')" }
    } catch { $out += "[FAIL] Could not turn auto-lock on: $($_.Exception.Message)" }
    return $out
}

function Uninstall-AutoApply {
    $out = @()
    try {
        $isAdm = ([Security.Principal.WindowsPrincipal][Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)
        if (-not $isAdm) { return @('[FAIL] This needs administrator rights. Nothing was changed.') }
        if (Get-ScheduledTask -TaskName $AutoApplyTask -ErrorAction SilentlyContinue) { Unregister-ScheduledTask -TaskName $AutoApplyTask -Confirm:$false -ErrorAction Stop }
        $pf = Join-Path $AutoApplyDir 'policy.json'
        if (Test-Path $pf) {
            $p = Get-Content $pf -Raw | ConvertFrom-Json; $p.Enabled = $false
            ($p | ConvertTo-Json -Depth 5) | Set-Content -Path $pf -Encoding UTF8
        }
        $out += '[OK] Auto-lock is OFF (the task was removed). Accounts that were already locked stay locked; new accounts are no longer touched.'
    } catch { $out += "[FAIL] Could not turn auto-lock off: $($_.Exception.Message)" }
    return $out
}

function Get-AutoApplyStatus {
    $out = @()
    $pf = Join-Path $AutoApplyDir 'policy.json'
    $task = Get-ScheduledTask -TaskName $AutoApplyTask -ErrorAction SilentlyContinue
    $pol = $null
    if (Test-Path $pf) { try { $pol = Get-Content $pf -Raw | ConvertFrom-Json } catch { } }
    if ($task -and $pol -and $pol.Enabled) {
        $info = $task | Get-ScheduledTaskInfo -ErrorAction SilentlyContinue
        $out += "[PASS] Auto-lock is ON ($($task.State)). Schedule: /times:$($pol.TimesValue)   (saved $($pol.Updated) on $($pol.UpdatedOn))"
        $out += "       Last run: $($info.LastRunTime)  result $($info.LastTaskResult)   Next run: $($info.NextRunTime)"
        if (@($pol.Excluded).Count -gt 0) { $out += "       Never touched: $(@($pol.Excluded) -join ', ')" }
    } elseif ($task) { $out += '[WARN] The auto-lock task exists but its policy is off or unreadable.' }
    else { $out += '[FAIL] Auto-lock is OFF - accounts created later will NOT get the time restrictions.' }
    $builtin = @('Administrator', 'DefaultAccount', 'Guest', 'WDAGUtilityAccount')
    $admins = @(); try { $admins = @(Get-LocalGroupMember -SID 'S-1-5-32-544' -ErrorAction Stop | ForEach-Object { ($_.Name -split '\\')[-1] }) } catch { }
    foreach ($u in @(Get-LocalUser -ErrorAction SilentlyContinue | Where-Object { $_.Enabled -and $builtin -notcontains $_.Name })) {
        if ($admins -icontains $u.Name) { $out += "       $($u.Name): administrator - never touched"; continue }
        $b = $null; try { $b = ([ADSI]"WinNT://./$($u.Name),user").Properties['LoginHours'].Value } catch { }
        $open = (-not $b) -or (@($b | Where-Object { $_ -ne 255 }).Count -eq 0)
        $ex = ($pol -and (@($pol.Excluded) -icontains $u.Name))
        $pend = (@(Get-ScheduledTask -TaskName "StudentTimes-$($u.Name)-*-Relock" -ErrorAction SilentlyContinue).Count -gt 0)
        if ($ex) { $out += "       $($u.Name): on the never-touch list" }
        elseif ($pend) { $out += "       $($u.Name): temporarily unlocked (re-locks by itself)" }
        elseif ($open) { $out += $(if ($task) { "[WARN] $($u.Name): hours are UNRESTRICTED right now (will be locked within 10 minutes)" } else { "[FAIL] $($u.Name): hours are UNRESTRICTED" }) }
        else { $out += "[PASS] $($u.Name): hours are restricted" }
    }
    $lg = Join-Path $AutoApplyDir 'autoapply.log'
    if (Test-Path $lg) { $out += '--- last log lines ---'; $out += @(Get-Content $lg -Tail 8) }
    return $out
}

$btnAutoOn.Add_Click({
    $times = Build-TimesString
    if (-not $times) {
        [System.Windows.Forms.MessageBox]::Show('The grid on the Daily Use tab is EMPTY (no hours allowed) - that would lock every new account out completely. Load or paint the hours first.', 'Empty schedule') | Out-Null
        return
    }
    $groups = @(Get-LogoutKickoffGroups $Grid)
    $action = if ($cmbLogoutAction.SelectedItem -eq 'Full log off') { 'LogOff' } else { 'Lock' }
    $excl = @($txtAutoExclude.Text -split '[,;\s]+' | Where-Object { $_ })
    $c = [System.Windows.Forms.MessageBox]::Show(
        "Turn ON automatic locking on $($env:COMPUTERNAME)?`n`nEvery NON-ADMIN local account whose hours are still completely unrestricted - including accounts created later, like user5 - gets:`n  net user <account> /times:$times`nplus the matching weekly logout tasks.`n`nRight now this will also lock any existing unrestricted non-admin account (within a minute).`nNever touched: administrators, Windows built-in accounts, accounts that already have restricted hours, accounts in a temporary unlock$(if ($excl.Count) { ", and: $($excl -join ', ')" }).`n`nContinue?",
        'Auto-lock new accounts', 'YesNo', 'Question')
    if ($c -ne 'Yes') { return }
    $lines = @(Install-AutoApply $times $groups $action $excl)
    if (@($lines | Where-Object { $_ -match '^\[FAIL\]' }).Count -eq 0) {
        try { Start-ScheduledTask -TaskName $AutoApplyTask -ErrorAction Stop; Start-Sleep -Seconds 8 } catch { $lines += "[WARN] Could not start the first run now: $($_.Exception.Message)" }
        $lines += ''
        $lines += @(Get-AutoApplyStatus)
    }
    Write-DebugLog "Auto-lock new accounts ON:`n$($lines -join "`n")"
    $txtSetupReport.Text = "=== Auto-lock new accounts: ON ($(Get-Date -Format 'yyyy-MM-dd HH:mm')) ===`r`n" + ($lines -join "`r`n")
})

$btnAutoOff.Add_Click({
    $c = [System.Windows.Forms.MessageBox]::Show('Turn OFF automatic locking? Accounts that are already locked stay locked, but accounts created later will NOT get the restrictions.', 'Turn off', 'YesNo', 'Question')
    if ($c -ne 'Yes') { return }
    $lines = @(Uninstall-AutoApply)
    Write-DebugLog "Auto-lock new accounts OFF:`n$($lines -join "`n")"
    $txtSetupReport.Text = "=== Auto-lock new accounts: OFF ($(Get-Date -Format 'yyyy-MM-dd HH:mm')) ===`r`n" + ($lines -join "`r`n")
})

$btnAutoStatus.Add_Click({
    $lines = @(Get-AutoApplyStatus)
    $txtSetupReport.Text = "=== Auto-lock new accounts: status on $($env:COMPUTERNAME) ($(Get-Date -Format 'yyyy-MM-dd HH:mm')) ===`r`n" + ($lines -join "`r`n")
})

# --- Shared apply/emit logic (Generate, Unlock Today, Restore, Slot Load all use this) ---
function Invoke-NetUserTimes([string]$userName, [string]$timesValue) {
    # Runs the actual net.exe command and returns {Success, ExitCode, Output}.
    $argTimes = "/times:$timesValue"
    $psi = New-Object System.Diagnostics.ProcessStartInfo
    $psi.FileName = 'net.exe'
    $psi.Arguments = 'user "' + $userName + '" "' + $argTimes + '"'
    $psi.RedirectStandardOutput = $true
    $psi.RedirectStandardError = $true
    $psi.UseShellExecute = $false
    $psi.CreateNoWindow = $true
    $proc = New-Object System.Diagnostics.Process
    $proc.StartInfo = $psi
    [void]$proc.Start()
    $stdout = $proc.StandardOutput.ReadToEnd()
    $stderr = $proc.StandardError.ReadToEnd()
    $proc.WaitForExit()
    $combined = ($stdout + $stderr).Trim()
    return [PSCustomObject]@{
        Success  = ($proc.ExitCode -eq 0)
        ExitCode = $proc.ExitCode
        Output   = $combined
    }
}

function Apply-Grid-AndEmit([string]$userName, [string]$actionLabel) {
    # Applies / emits the grid for EVERY target account (the ticked accounts, or the typed one).
    $targets = @(Get-TargetAccounts)
    if ($targets.Count -eq 0) { [System.Windows.Forms.MessageBox]::Show('Enter a Windows account name (or pick accounts) first.', 'Missing account') | Out-Null; return }
    $timesValue = Build-TimesString
    $commands = @($targets | ForEach-Object { "net user $_ /times:$timesValue" })
    $txtOut.Text = $commands -join "`r`n"

    if ($chkRunNow.Checked) {
        $emptyWarn = if (-not $timesValue) { "`n`nWARNING: the grid is EMPTY - this means these accounts may NEVER log on." } else { '' }
        $confirm = [System.Windows.Forms.MessageBox]::Show(
            "This will run these command(s) RIGHT NOW on $($env:COMPUTERNAME):`n`n$($commands -join "`n")$emptyWarn`n`nContinue?",
            'Confirm immediate change', 'YesNo', 'Warning')
        if ($confirm -ne 'Yes') {
            Write-DebugLog "Apply-Grid-AndEmit: user CANCELLED immediate run of: $($commands -join ' ; ')"
            return
        }
        $lines = @(); $bad = 0
        foreach ($acct in $targets) {
            $command = "net user $acct /times:$timesValue"
            Write-DebugLog "Apply-Grid-AndEmit: running immediately -> $command"
            $result = Invoke-NetUserTimes -userName $acct -timesValue $timesValue
            Write-DebugLog "Apply-Grid-AndEmit: exitcode=$($result.ExitCode) output=$($result.Output)"
            if ($result.Success) {
                Write-LogEntry -userName $acct -oldTimes $script:lastOldTimes -newTimes $command -note "RAN IMMEDIATELY - OK. $actionLabel"
                $syncLines = @(Sync-LogoutTasks -userName $acct -grid $Grid)
                Write-DebugLog "Apply-Grid-AndEmit: logout task sync ($acct):`n$($syncLines -join "`n")"
                $lines += "-- $acct --"
                $lines += '[OK] net user /times applied.'
                $lines += $syncLines
                if (@($syncLines | Where-Object { $_ -match '^\[FAIL\]' }).Count -gt 0) { $bad++ }
            } else {
                Write-LogEntry -userName $acct -oldTimes $script:lastOldTimes -newTimes $command -note "RAN IMMEDIATELY - FAILED (exit $($result.ExitCode)): $($result.Output). $actionLabel"
                $lines += "-- $acct --"
                $lines += "[FAIL] net user failed (exit $($result.ExitCode)): $($result.Output) - nothing changed on this account."
                $bad++
            }
        }
        $script:lastOldTimes = $timesValue
        $script:lastNewTimes = $commands[0]
        $btnCopyRow.Enabled = $true
        $txtSetupReport.Text = "=== Hours applied + logout tasks synced on $($env:COMPUTERNAME) ($(Get-Date -Format 'yyyy-MM-dd HH:mm')) ===`r`n" + ($lines -join "`r`n")
        if ($bad -gt 0) {
            [System.Windows.Forms.MessageBox]::Show("$bad of $($targets.Count) account(s) had a PROBLEM - see the red report on the Setup / Verify tab. (Check the account names and that this window is running elevated.)", 'PROBLEM', 'OK', 'Warning') | Out-Null
        } else {
            [System.Windows.Forms.MessageBox]::Show("Applied to $($targets.Count) account(s) on $($env:COMPUTERNAME) and the logout tasks were synced:`n`n$($targets -join ', ')`n`n$($commands[0])", 'Applied') | Out-Null
        }
        return
    }

    [System.Windows.Forms.Clipboard]::SetText($commands -join "`r`n")
    foreach ($acct in $targets) { Write-LogEntry -userName $acct -oldTimes $script:lastOldTimes -newTimes "net user $acct /times:$timesValue" -note $actionLabel }
    $script:lastOldTimes = $timesValue
    $script:lastNewTimes = $commands[0]
    $btnCopyRow.Enabled = $true
    [System.Windows.Forms.MessageBox]::Show("$($commands.Count) command(s) copied to clipboard:`n`n$($commands -join "`n")`n`nPaste into an elevated Command Prompt on this PC ($($env:COMPUTERNAME)) to apply.`n`nIMPORTANT: this only COPIED the command(s) - the logout tasks are NOT updated yet. After pasting, click 'Sync Logout Tasks To Grid' on the Setup / Verify tab (or tick 'Run it immediately here' next time and both happen together).", 'Copied') | Out-Null
}

function Update-SlotLabels {
    $userName = $txtUser.Text.Trim()
    if (-not $userName) { return }

    $auto = Get-UserSlot -userName $userName -slotKey 'AutoBackup'
    if ($null -ne $auto) {
        $lblAutoBackupInfo.Text = "Saved before last unlock: $($auto.Timestamp)"
        $pend = @(Get-PendingRelockTasks $userName)
        if ($pend.Count -gt 0) {
            $nr = $pend[0] | Get-ScheduledTaskInfo -ErrorAction SilentlyContinue
            if ($nr -and $nr.NextRunTime) { $lblAutoBackupInfo.Text += "   |   AUTO RE-LOCK: $($nr.NextRunTime.ToString('ddd MMM d, h:mm tt'))" }
        }
        $btnRestorePrevious.Enabled = $true
    } else {
        $lblAutoBackupInfo.Text = 'No auto-saved previous schedule yet for this account.'
        $btnRestorePrevious.Enabled = $false
    }

    for ($i = 1; $i -le $slotCount; $i++) {
        $slot = Get-UserSlot -userName $userName -slotKey "Slot$i"
        if ($null -ne $slot) {
            $slotTimestampLabels[$i - 1].Text = "saved $($slot.Timestamp)"
            $slotTextBoxes[$i - 1].Text = $slot.TimesString
            if ($slot.Name) { $slotNameBoxes[$i - 1].Text = $slot.Name }
        } else {
            $slotTimestampLabels[$i - 1].Text = '(empty)'
            $slotTextBoxes[$i - 1].Text = ''
            $slotNameBoxes[$i - 1].Text = "Slot $i"
        }
    }
}

$btnLoad.Add_Click({
    $userName = $txtUser.Text.Trim()
    Write-DebugLog "Load Current Schedule clicked. Account field = '$userName'"
    if (-not $userName) {
        [System.Windows.Forms.MessageBox]::Show('Enter a Windows account name first.', 'Missing account') | Out-Null
        return
    }
    $existing = Get-CurrentLoginHours $userName
    if ($null -eq $existing) {
        Write-DebugLog "Load Current Schedule: Get-CurrentLoginHours returned NULL for '$userName' (account not found or ADSI error - see error above)"
        [System.Windows.Forms.MessageBox]::Show("Could not read a local account named '$userName' on this PC. Click 'Open Debug Log' for details.", 'Not found') | Out-Null
        return
    }
    for ($d = 0; $d -lt 7; $d++) { for ($h = 0; $h -lt 24; $h++) { $Grid[$d,$h] = $existing[$d,$h] } }
    Refresh-GridVisuals
    $script:lastOldTimes = Build-TimesString
    Write-DebugLog "Load Current Schedule: grid set to -> $($script:lastOldTimes)"
    $txtOut.Text = "Loaded current schedule for '$userName' on $($env:COMPUTERNAME). Edit the grid, then Generate."
})

$btnAllowAll.Add_Click({
    Reset-Grid $true
    Refresh-GridVisuals
})

$btnClear.Add_Click({
    Reset-Grid $false
    Refresh-GridVisuals
})

$btnGenerate.Add_Click({
    $userName = $txtUser.Text.Trim()
    if (-not $userName) {
        [System.Windows.Forms.MessageBox]::Show('Enter a Windows account name first.', 'Missing account') | Out-Null
        return
    }
    Apply-Grid-AndEmit -userName $userName -actionLabel ''
})

$btnCopyRow.Add_Click({
    # Tab-separated so pasting into Excel drops it straight into columns
    # matching the checklist: PC Name | Times Requested
    $pc = $env:COMPUTERNAME
    $row = "$pc`t$($script:lastNewTimes)"
    [System.Windows.Forms.Clipboard]::SetText($row)
    [System.Windows.Forms.MessageBox]::Show("Row copied - paste directly into the 'PC Name' cell of Student Times Checklist.xlsx (it will fill PC Name + Times Requested).", 'Copied') | Out-Null
})

function Get-PendingRelockTasks([string]$userName) {
    return @(Get-ScheduledTask -TaskName "StudentTimes-$userName-*-Relock" -ErrorAction SilentlyContinue)
}

function Remove-PendingRelockTasks([string]$userName) {
    foreach ($t in @(Get-PendingRelockTasks $userName)) {
        try { Unregister-ScheduledTask -TaskName $t.TaskName -Confirm:$false -ErrorAction Stop } catch { Write-DebugLog "Could not remove relock task $($t.TaskName): $($_.Exception.Message)" }
        $f = Join-Path (Join-Path $env:ProgramData 'StudentTimes') ($t.TaskName + '.ps1')
        Remove-Item $f -Force -ErrorAction SilentlyContinue
    }
}

function Start-UnlockPreset([int]$days) {
    # Unlock NOW (every target account), and arrange the re-lock by itself: a one-time SYSTEM task per account at
    # midnight at the end of day $days puts that account's original hours AND logout tasks back. Always applies
    # immediately (ignores the "Run it immediately" box, which is for hand edits).
    $targets = @(Get-TargetAccounts)
    if ($targets.Count -eq 0) { [System.Windows.Forms.MessageBox]::Show('Enter a Windows account name (or pick accounts) first.', 'Missing account') | Out-Null; return }

    $relockAt = (Get-Date).Date.AddDays($days)
    $plan = @()
    foreach ($acct in $targets) {
        $existing = Get-CurrentLoginHours $acct
        if ($null -eq $existing) { [System.Windows.Forms.MessageBox]::Show("Could not read a local account named '$acct' on this PC.", 'Not found') | Out-Null; return }
        $pending = @(Get-PendingRelockTasks $acct)
        $auto = Get-UserSlot -userName $acct -slotKey 'AutoBackup'
        $keepSaved = ($pending.Count -gt 0 -and $null -ne $auto)
        if ($keepSaved) {
            # already unlocked by an earlier preset: the saved schedule is the REAL one - keep it, just move the re-lock time
            $originalTimes = $auto.TimesString
        } else {
            for ($d = 0; $d -lt 7; $d++) { for ($h = 0; $h -lt 24; $h++) { $Grid[$d,$h] = $existing[$d,$h] } }
            $originalTimes = Build-TimesString
        }
        $plan += [pscustomobject]@{ Name = $acct; Original = $originalTimes; KeepSaved = $keepSaved }
    }
    $endText = $relockAt.ToString('dddd, MMMM d, yyyy') + ' 12:00 AM (midnight)'
    $listText = ($plan | ForEach-Object { "  $($_.Name): net user $($_.Name) /times:$($_.Original)" }) -join "`n"
    $c = [System.Windows.Forms.MessageBox]::Show(
        "This will, on $($env:COMPUTERNAME):`n`n1. UNLOCK $($plan.Count) account(s) right now (all hours allowed, logout tasks removed)`n2. RE-LOCK automatically at $endText - the saved schedules and logout tasks come back by themselves`n`nSaved schedules that will be restored:`n$listText`n`nContinue?",
        "Unlock for $days day(s)", 'YesNo', 'Question')
    if ($c -ne 'Yes') { return }

    $lines = @(); $problems = 0
    foreach ($p in $plan) {
        $acct = $p.Name
        $lines += "-- $acct --"
        if (-not $p.KeepSaved) { Set-UserSlot -userName $acct -slotKey 'AutoBackup' -timesValue $p.Original }
        Reset-Grid $true
        $openTimes = Build-TimesString
        $res = Invoke-NetUserTimes -userName $acct -timesValue $openTimes
        Write-DebugLog "Unlock preset ($days d) $acct : net user exit=$($res.ExitCode) $($res.Output)"
        if (-not $res.Success) { $lines += "[FAIL] The unlock FAILED (exit $($res.ExitCode)): $($res.Output) - nothing scheduled for this account."; $problems++; continue }
        $syncLines = @(Sync-LogoutTasks -userName $acct -grid $Grid)
        Remove-PendingRelockTasks $acct
        $origGrid = ConvertTo-GridFromTimesString $p.Original
        $reg = Register-FutureNetUserChange -userName $acct -timesValue $p.Original -when $relockAt -label 'Relock' -taskGrid $origGrid
        Write-DebugLog "Unlock preset $acct : relock task '$($reg.TaskName)' success=$($reg.Success) output=$($reg.Output)"
        $verified = $false; $nextRun = ''
        if ($reg.Success) {
            $t = Get-ScheduledTask -TaskName $reg.TaskName -ErrorAction SilentlyContinue
            if ($t) { $info = $t | Get-ScheduledTaskInfo; $nextRun = $info.NextRunTime; if ($nextRun -and [math]::Abs(($nextRun - $relockAt).TotalMinutes) -lt 2) { $verified = $true } }
        }
        $lines += "[OK] '$acct' is UNLOCKED now (all hours allowed)."
        $lines += $syncLines
        if ($verified) { $lines += "[OK] Automatic re-lock scheduled for $nextRun ($($reg.TaskName)); restores net user /times:$($p.Original) and the logout tasks." }
        else { $lines += "[FAIL] The automatic re-lock for '$acct' could NOT be confirmed ($(if ($reg.Success) { 'scheduled time did not match' } else { $reg.Output })). It STAYS UNLOCKED until you click 'Re-lock Now'."; $problems++ }
        Write-LogEntry -userName $acct -oldTimes $p.Original -newTimes "net user $acct /times:$openTimes" -note "Unlock $days day(s); auto re-lock $relockAt"
    }
    Refresh-GridVisuals
    $script:lastOldTimes = $plan[0].Original
    $txtSetupReport.Text = "=== Unlock for $days day(s) on $($env:COMPUTERNAME) ($(Get-Date -Format 'yyyy-MM-dd HH:mm')) ===`r`n" + ($lines -join "`r`n")
    Update-SlotLabels
    if ($problems -eq 0) { [System.Windows.Forms.MessageBox]::Show("Unlocked $($plan.Count) account(s) now. They re-lock by themselves at $($relockAt.ToString('dddd MMM d, h:mm tt')).`n`n(Details: Setup / Verify tab report.)", 'Unlocked') | Out-Null }
    else { [System.Windows.Forms.MessageBox]::Show("$problems problem(s) - see the red report on the Setup / Verify tab. Any account whose re-lock was not confirmed stays unlocked until you click 'Re-lock Now'.", 'PROBLEM', 'OK', 'Warning') | Out-Null }
}

$btnUnlockToday.Add_Click({ Start-UnlockPreset 1 })
$btnUnlock2Days.Add_Click({ Start-UnlockPreset 2 })
$btnUnlock3Days.Add_Click({ Start-UnlockPreset 3 })

$btnRestorePrevious.Add_Click({
    $targets = @(Get-TargetAccounts)
    if ($targets.Count -eq 0) { [System.Windows.Forms.MessageBox]::Show('Enter a Windows account name (or pick accounts) first.', 'Missing account') | Out-Null; return }
    $c = [System.Windows.Forms.MessageBox]::Show("Put the saved schedule back NOW (and cancel any automatic re-lock) for: $($targets -join ', ')?", 'Re-lock now', 'YesNo', 'Question')
    if ($c -ne 'Yes') { return }
    $lines = @(); $problems = 0; $first = $true
    foreach ($acct in $targets) {
        $lines += "-- $acct --"
        $auto = Get-UserSlot -userName $acct -slotKey 'AutoBackup'
        if ($null -eq $auto) { $lines += "[WARN] No saved schedule for '$acct' - skipped."; continue }
        $g = ConvertTo-GridFromTimesString $auto.TimesString
        $res = Invoke-NetUserTimes -userName $acct -timesValue $auto.TimesString
        Write-DebugLog "Re-lock Now $acct : exit=$($res.ExitCode) $($res.Output)"
        if (-not $res.Success) { $lines += "[FAIL] Could not restore (exit $($res.ExitCode)): $($res.Output)"; $problems++; continue }
        $lines += "[OK] Restored net user /times:$($auto.TimesString)"
        $syncLines = @(Sync-LogoutTasks -userName $acct -grid $g)
        $lines += $syncLines
        if (@($syncLines | Where-Object { $_ -match '^\[FAIL\]' }).Count -gt 0) { $problems++ }
        Remove-PendingRelockTasks $acct
        Write-LogEntry -userName $acct -oldTimes '' -newTimes "net user $acct /times:$($auto.TimesString)" -note 'Re-lock Now (restore previous)'
        if ($first) { Parse-TimesStringToGrid $auto.TimesString; Refresh-GridVisuals; $first = $false }
    }
    $txtSetupReport.Text = "=== Re-lock now on $($env:COMPUTERNAME) ($(Get-Date -Format 'yyyy-MM-dd HH:mm')) ===`r`n" + ($lines -join "`r`n")
    Update-SlotLabels
    if ($problems -gt 0) { [System.Windows.Forms.MessageBox]::Show("$problems problem(s) - see the red report on the Setup / Verify tab.", 'PROBLEM', 'OK', 'Warning') | Out-Null }
})

$txtUser.Add_TextChanged({
    if (@($script:selectedAccounts).Count -gt 0 -and ($script:selectedAccounts -inotcontains $txtUser.Text.Trim())) { $script:selectedAccounts = @() }
    Update-ApplyToLabel
    Update-SlotLabels
})

Update-SlotLabels
Update-ApplyToLabel

# Start with THIS PC's real current schedule in the grid (instead of an empty one), so a
# push/generate without clicking Load first sends what the account has now, not "no hours".
try {
    $startUser = $txtUser.Text.Trim()
    if ($startUser) {
        $startHours = Get-CurrentLoginHours $startUser
        if ($null -ne $startHours) {
            for ($d = 0; $d -lt 7; $d++) { for ($h = 0; $h -lt 24; $h++) { $Grid[$d,$h] = $startHours[$d,$h] } }
            Refresh-GridVisuals
            $script:lastOldTimes = Build-TimesString
            $txtOut.Text = "Started with the current schedule of '$startUser' on $($env:COMPUTERNAME) already loaded. Edit the grid, then Generate."
            Write-DebugLog "Startup auto-load: grid set to -> $($script:lastOldTimes)"
        }
    }
} catch { Write-DebugLog "Startup auto-load skipped: $($_.Exception.Message)" }

Write-DebugLog "=== Set-StudentTimes.ps1 v$ScriptVersion started on $env:COMPUTERNAME ==="
[void]$form.ShowDialog()
