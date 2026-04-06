#Requires -RunAsAdministrator
# Get-WorkingTimes.ps1  –  Windows 11 DE
# Detailed daily summary with login/lock/sleep events and work hours.

param(
    [string]$StartDate,
    [string]$EndDate,
    [int]   $DaysBack  = 7,
    [switch]$ShowDetail
)

# ── Date range ───────────────────────────────────────────────────────────────
if ($StartDate) {
    try { $since = [datetime]::ParseExact($StartDate, 'yyyy-MM-dd', $null) }
    catch { Write-Error "StartDate must be yyyy-MM-dd"; exit 1 }
} else { $since = (Get-Date).Date.AddDays(-$DaysBack) }

if ($EndDate) {
    try { $until = [datetime]::ParseExact($EndDate, 'yyyy-MM-dd', $null).AddDays(1).AddSeconds(-1) }
    catch { Write-Error "EndDate must be yyyy-MM-dd"; exit 1 }
} else { $until = Get-Date }

Write-Host "Lese Ereignisse: $($since.ToString('dd.MM.yyyy')) – $($until.ToString('dd.MM.yyyy')) ..." -ForegroundColor Cyan

# ── Collect events ───────────────────────────────────────────────────────────
function Get-SafeEvents($log, $ids) {
    try {
        Get-WinEvent -FilterHashtable @{ LogName=$log; Id=$ids; StartTime=$since; EndTime=$until } -ErrorAction Stop
    } catch {
        if ($_.Exception.Message -notmatch 'No events') { Write-Warning "'$log': $($_.Exception.Message)" }
        @()
    }
}

$secEvents = Get-SafeEvents 'Sicherheit' (4624,4634,4800,4801)
$sysEvents = Get-SafeEvents 'System'     (1,12,41,42,6005,6006)

# ── Normalize ────────────────────────────────────────────────────────────────
$allEvents = @()

foreach ($e in $secEvents) {
    $type = switch ($e.Id) {
        4624 { $lt = $e.Properties[8].Value; if ($lt -in 2,7) { 'Unlock' } else { $null } }
        4634 { 'Lock'   }
        4800 { 'Lock'   }
        4801 { 'Unlock' }
    }
    if ($type) { $allEvents += [PSCustomObject]@{ Time=$e.TimeCreated; Type=$type } }
}
foreach ($e in $sysEvents) {
    $type = switch ($e.Id) {
        1    { 'Wake'  }  12   { 'Wake'  }
        41   { 'Wake'  }  42   { 'Sleep' }
        6005 { 'Wake'  }  6006 { 'Sleep' }
    }
    $allEvents += [PSCustomObject]@{ Time=$e.TimeCreated; Type=$type }
}

$allEvents = $allEvents | Sort-Object Time

if ($allEvents.Count -eq 0) {
    Write-Warning "Keine Ereignisse gefunden. Als Administrator starten und Überwachungsrichtlinien prüfen."
    exit
}

# ── Walk timeline & collect sessions + interruptions ────────────────────────
$awake     = $false
$unlocked  = $false
$workStart = $null
$sessions  = @()   # active work intervals
$breaks    = @()   # interruption intervals (lock or sleep)
$breakStart = $null
$breakReason = $null

function Open-Work($t) {
    if (-not $script:workStart) { $script:workStart = $t }
    if ($script:breakStart) {
        $script:breaks += [PSCustomObject]@{
            Start  = $script:breakStart
            End    = $t
            Reason = $script:breakReason
            Duration = $t - $script:breakStart
        }
        $script:breakStart = $null; $script:breakReason = $null
    }
}
function Close-Work($t, $reason) {
    if ($script:workStart -and $t -gt $script:workStart) {
        $script:sessions += [PSCustomObject]@{
            Start    = $script:workStart
            End      = $t
            Duration = $t - $script:workStart
        }
    }
    $script:workStart = $null
    if (-not $script:breakStart) { $script:breakStart = $t; $script:breakReason = $reason }
}

foreach ($e in $allEvents) {
    switch ($e.Type) {
        'Wake'   { $awake = $true;   if ($awake -and $unlocked) { Open-Work  $e.Time } }
        'Sleep'  { Close-Work $e.Time 'Standby/Herunterfahren'; $awake=$false; $unlocked=$false }
        'Unlock' { $unlocked = $true;  if ($awake -and $unlocked) { Open-Work  $e.Time } }
        'Lock'   { Close-Work $e.Time 'Gesperrt/Abgemeldet';   $unlocked=$false }
    }
}
if ($awake -and $unlocked -and $workStart) { Close-Work (Get-Date) 'Aktuell aktiv' }

# ── Output helper ────────────────────────────────────────────────────────────
function Format-Dur($ts) { "{0}h {1:D2}m" -f [Math]::Floor($ts.TotalMinutes/60), ([Math]::Round($ts.TotalMinutes%60)) }
function Write-Sep  { Write-Host ('─' * 62) -ForegroundColor DarkGray }

# ── Daily output ─────────────────────────────────────────────────────────────
$allDays = ($sessions + $breaks) | Select-Object -ExpandProperty Start |
    ForEach-Object { $_.Date } | Sort-Object -Unique

$grandWork  = [timespan]::Zero
$grandBreak = [timespan]::Zero

foreach ($day in $allDays) {
    $daySessions = $sessions | Where-Object { $_.Start.Date -eq $day } | Sort-Object Start
    $dayBreaks   = $breaks   | Where-Object { $_.Start.Date -eq $day } | Sort-Object Start

    if (-not $daySessions) { continue }

    $workTotal  = $daySessions | Measure-Object -Property { $_.Duration.TotalSeconds } -Sum
    $breakTotal = $dayBreaks   | Measure-Object -Property { $_.Duration.TotalSeconds } -Sum
    $workTs     = [timespan]::FromSeconds($workTotal.Sum)
    $breakTs    = [timespan]::FromSeconds($breakTotal.Sum)
    $dayStart   = ($daySessions | Sort-Object Start | Select-Object -First 1).Start
    $dayEnd     = ($daySessions | Sort-Object End   | Select-Object -Last  1).End
    $grandWork  += $workTs
    $grandBreak += $breakTs

    Write-Sep
    Write-Host (" {0:dddd, dd. MMMM yyyy}" -f $day) -ForegroundColor White
    Write-Host (" Anwesenheit:  {0:HH:mm} – {1:HH:mm}  (gesamt {2})" -f $dayStart, $dayEnd, (Format-Dur ($dayEnd-$dayStart))) -ForegroundColor Gray
    Write-Host (" Aktive Arbeit:{0,8}   |   Unterbrechungen: {1}" -f (Format-Dur $workTs), (Format-Dur $breakTs)) -ForegroundColor Green
    Write-Host ""

    # Merge sessions & breaks into one sorted timeline for the detail view
    $timeline = @(
        $daySessions | Select-Object Start,End,Duration,@{N='Kind';E={'Arbeit'}},@{N='Reason';E={$null}}
        $dayBreaks   | Select-Object Start,End,Duration,@{N='Kind';E={'Pause'}},Reason
    ) | Sort-Object Start

    foreach ($ev in $timeline) {
        $color  = if ($ev.Kind -eq 'Arbeit') { 'Cyan' } else { 'DarkYellow' }
        $icon   = if ($ev.Kind -eq 'Arbeit') { '[+]' } else { '[-]' }
        $reason = if ($ev.Reason) { "  ($($ev.Reason))" } else { '' }
        Write-Host ("   $icon {0:HH:mm} – {1:HH:mm}  {2,8}{3}" -f $ev.Start, $ev.End, (Format-Dur $ev.Duration), $reason) -ForegroundColor $color
    }
    Write-Host ""
}

# ── Grand total ───────────────────────────────────────────────────────────────
$label = if ($StartDate -or $EndDate) {
    "$($since.ToString('dd.MM.yyyy')) – $($until.ToString('dd.MM.yyyy'))"
} else { "letzte $DaysBack Tage" }

Write-Sep
Write-Host (" Gesamt ($label)") -ForegroundColor White
Write-Host (" Aktive Arbeit:    {0}" -f (Format-Dur $grandWork))  -ForegroundColor Green
Write-Host (" Unterbrechungen:  {0}" -f (Format-Dur $grandBreak)) -ForegroundColor DarkYellow
Write-Host (" Tage ausgewertet: {0}" -f $allDays.Count) -ForegroundColor Gray
Write-Sep
