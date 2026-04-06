#
# arbeitszeiten_fixed.ps1  –  Windows 11 DE, ohne Adminrechte
#
# Logs (ohne Admin lesbar):
#   Microsoft-Windows-Winlogon/Operational   7001 = Anmeldung/Entsperren
#                                            7002 = Abmeldung/Sperren
#   System                                   1/12/41/6005 = Wake/Start
#                                            42/6006/13   = Sleep/Shutdown
#
# Genauigkeit:
#   Zwei unabhängige Gates (awake + unlocked) müssen BEIDE offen sein
#   damit Zeit als Arbeitszeit zählt. Pausen werden exakt gemessen,
#   nicht aus der Differenz von Gesamtzeit minus errechneter Summe.
#
# Usage:
#   .\arbeitszeiten_fixed.ps1 -StartDate "2026-03-01" -EndDate "2026-03-31"
#   .\arbeitszeiten_fixed.ps1 -StartDate "2026-03-01"
#   .\arbeitszeiten_fixed.ps1   (letzte 7 Tage)
#

param(
    [string]$StartDate  = "",
    [string]$EndDate    = "",
    [int]   $DaysBack   = 7,
    [string]$OutputPath = "C:\Data\Zeiterfassung\Reports.csv"
)

# ── Date range ───────────────────────────────────────────────────────────────
if ($StartDate -ne "") {
    try   { $since = [datetime]::ParseExact($StartDate, 'yyyy-MM-dd', [System.Globalization.CultureInfo]::InvariantCulture) }
    catch { Write-Error "Ungültiges StartDate-Format. Erwartet: yyyy-MM-dd"; exit 1 }
} else {
    $since = (Get-Date).Date.AddDays(-$DaysBack)
}

if ($EndDate -ne "") {
    try   { $until = [datetime]::ParseExact($EndDate, 'yyyy-MM-dd', [System.Globalization.CultureInfo]::InvariantCulture).AddDays(1).AddSeconds(-1) }
    catch { Write-Error "Ungültiges EndDate-Format. Erwartet: yyyy-MM-dd"; exit 1 }
} else {
    $until = Get-Date
}

if ($since -gt $until) { Write-Error "StartDate darf nicht nach EndDate liegen."; exit 1 }

Write-Host "`n=== Arbeitszeiten-Auswertung ===" -ForegroundColor Cyan
Write-Host "Zeitraum : $($since.ToString('dd.MM.yyyy')) – $($until.ToString('dd.MM.yyyy'))" -ForegroundColor Cyan
Write-Host "Export   : $OutputPath`n" -ForegroundColor Cyan

# ── Helpers ───────────────────────────────────────────────────────────────────
function Get-SafeEvents([string]$log, [int[]]$ids) {
    Write-Host "Lese '$log' (IDs: $($ids -join ', ')) ..." -ForegroundColor Yellow
    try {
        $ev = Get-WinEvent -FilterHashtable @{
            LogName   = $log
            Id        = $ids
            StartTime = $since
            EndTime   = $until
        } -ErrorAction Stop
        Write-Host "  $($ev.Count) Ereignisse gefunden." -ForegroundColor Green
        return $ev
    } catch {
        if ($_.Exception.Message -notmatch 'No events|keine.*Ereignisse') {
            Write-Warning "  '$log' nicht lesbar: $($_.Exception.Message)"
        } else {
            Write-Host "  Keine Ereignisse." -ForegroundColor Gray
        }
        return @()
    }
}

function Format-Dur([timespan]$ts) {
    if ($null -eq $ts -or $ts.TotalSeconds -le 0) { return '0h 00m' }
    $h = [Math]::Floor($ts.TotalMinutes / 60)
    $m = [Math]::Round($ts.TotalMinutes % 60)
    return ('{0}h {1:D2}m' -f $h, $m)
}

function Sum-Duration($items) {
    $sec = 0.0
    foreach ($i in $items) { $sec += $i.Duration.TotalSeconds }
    return [timespan]::FromSeconds($sec)
}

function Write-Sep([string]$c = '─') { Write-Host ($c * 66) -ForegroundColor DarkGray }

# ── Collect raw events ────────────────────────────────────────────────────────
$winlogonRaw = Get-SafeEvents 'Microsoft-Windows-Winlogon/Operational' (7001, 7002)
$systemRaw   = Get-SafeEvents 'System' (1, 12, 13, 41, 42, 6005, 6006)

# ── Normalize to unified list ─────────────────────────────────────────────────
# Each entry: Time, Type (Wake|Sleep|Unlock|Lock), EventId, LogName, RawDescription
$allEvents = [System.Collections.Generic.List[object]]::new()

foreach ($e in $winlogonRaw) {
    $type = switch ($e.Id) {
        7001 { 'Unlock' }
        7002 { 'Lock'   }
    }
    $desc = switch ($e.Id) {
        7001 { 'Anmeldung / Entsperren' }
        7002 { 'Abmeldung / Sperren'    }
    }
    $allEvents.Add([PSCustomObject]@{
        Time        = $e.TimeCreated
        Type        = $type
        EventId     = $e.Id
        LogName     = $e.LogName
        Description = $desc
    })
}

foreach ($e in $systemRaw) {
    $type = switch ($e.Id) {
        1    { 'Wake'  }   12   { 'Wake'  }
        41   { 'Wake'  }   6005 { 'Wake'  }
        42   { 'Sleep' }   13   { 'Sleep' }
        6006 { 'Sleep' }
    }
    $desc = switch ($e.Id) {
        1    { 'Reaktivierung aus Standby'      }
        12   { 'Systemstart'                    }
        41   { 'Unerwartete Reaktivierung'       }
        6005 { 'Systemstart (EventLog)'          }
        42   { 'Standby / Ruhezustand'           }
        13   { 'Kernel-Shutdown'                 }
        6006 { 'Herunterfahren (EventLog)'       }
    }
    $allEvents.Add([PSCustomObject]@{
        Time        = $e.TimeCreated
        Type        = $type
        EventId     = $e.Id
        LogName     = $e.LogName
        Description = $desc
    })
}

$allEvents = $allEvents | Sort-Object Time

if ($allEvents.Count -eq 0) {
    Write-Warning "Keine Ereignisse gefunden."
    Write-Host "`nDiagnose:"
    Write-Host "  Get-WinEvent -ListLog 'Microsoft-Windows-Winlogon/Operational' | Select-Object IsEnabled, RecordCount" -ForegroundColor Yellow
    exit
}

Write-Host "`nGesamt $($allEvents.Count) Ereignisse gesammelt.`n" -ForegroundColor Green

# ── Walk timeline with dual-gate logic ───────────────────────────────────────
#
# Gate 1: $awake    → false when sleep/shutdown, true when wake/boot
# Gate 2: $unlocked → false when locked/logoff, true when unlock/logon
#
# A work session starts  when BOTH gates flip to true.
# A work session ends    when EITHER gate flips to false.
# A break starts exactly when the work session ends.
# A break ends   exactly when both gates are true again.
#
# This means: no estimation, no subtraction — every second is accounted for.

$awake       = $false
$unlocked    = $false
$workStart   = $null
$breakStart  = $null
$breakReason = $null

$sessions = [System.Collections.Generic.List[object]]::new()
$breaks   = [System.Collections.Generic.List[object]]::new()
$rawLog   = [System.Collections.Generic.List[object]]::new()

function Open-Work([datetime]$t) {
    if (-not $script:workStart) {
        $script:workStart = $t
    }
    if ($script:breakStart -and $t -gt $script:breakStart) {
        $script:breaks.Add([PSCustomObject]@{
            Start    = $script:breakStart
            End      = $t
            Duration = ($t - $script:breakStart)
            Reason   = $script:breakReason
        })
        $script:breakStart  = $null
        $script:breakReason = $null
    }
}

function Close-Work([datetime]$t, [string]$reason) {
    if ($script:workStart -and $t -gt $script:workStart) {
        $script:sessions.Add([PSCustomObject]@{
            Start    = $script:workStart
            End      = $t
            Duration = ($t - $script:workStart)
        })
    }
    $script:workStart = $null
    if (-not $script:breakStart) {
        $script:breakStart  = $t
        $script:breakReason = $reason
    }
}

foreach ($e in $allEvents) {
    $wasWorking = ($awake -and $unlocked)

    switch ($e.Type) {
        'Wake'   { $awake    = $true  }
        'Sleep'  { $awake    = $false; $unlocked = $false }
        'Unlock' { $unlocked = $true  }
        'Lock'   { $unlocked = $false }
    }

    $isWorking = ($awake -and $unlocked)

    if (-not $wasWorking -and $isWorking)  { Open-Work  $e.Time }
    if ($wasWorking -and -not $isWorking)  { Close-Work $e.Time $e.Description }

    $rawLog.Add([PSCustomObject]@{
        Timestamp   = $e.Time
        Date        = $e.Time.ToString('yyyy-MM-dd')
        Time        = $e.Time.ToString('HH:mm:ss')
        Type        = $e.Type
        Description = $e.Description
        EventId     = $e.EventId
        LogName     = $e.LogName
        WorkingAfter = $isWorking
    })
}

# Close any session still open right now
if ($awake -and $unlocked -and $workStart) {
    Close-Work (Get-Date) 'Aktuell aktiv'
}

# ── Per-day output ────────────────────────────────────────────────────────────
$allDays = ($sessions.ToArray() + $breaks.ToArray()) |
    ForEach-Object { $_.Start.Date } |
    Sort-Object -Unique

$grandWork  = [timespan]::Zero
$grandBreak = [timespan]::Zero
$daySummaries = [System.Collections.Generic.List[object]]::new()

foreach ($day in $allDays) {
    $ds = @($sessions | Where-Object { $_.Start.Date -eq $day } | Sort-Object Start)
    $db = @($breaks   | Where-Object { $_.Start.Date -eq $day } | Sort-Object Start)

    if ($ds.Count -eq 0) { continue }

    $workTs  = Sum-Duration $ds
    $breakTs = Sum-Duration $db

    $dayStart = $ds[0].Start
    $dayEnd   = ($ds | Sort-Object End | Select-Object -Last 1).End
    $presence = $dayEnd - $dayStart

    $grandWork  += $workTs
    $grandBreak += $breakTs

    # Build pause detail list for CSV
    $pauseDetails = $db | ForEach-Object {
        "$($_.Start.ToString('HH:mm'))–$($_.End.ToString('HH:mm')) ($([Math]::Round($_.Duration.TotalMinutes))min, $($_.Reason))"
    }

    $daySummaries.Add([PSCustomObject]@{
        Datum           = $day.ToString('yyyy-MM-dd')
        Wochentag       = $day.ToString('dddd', [System.Globalization.CultureInfo]::GetCultureInfo('de-DE'))
        Arbeitsbeginn   = $dayStart.ToString('HH:mm')
        Arbeitsende     = $dayEnd.ToString('HH:mm')
        Anwesenheit_h   = [Math]::Round($presence.TotalHours, 2)
        Aktiv_h         = [Math]::Round($workTs.TotalHours, 2)
        Pausen_h        = [Math]::Round($breakTs.TotalHours, 2)
        Anzahl_Pausen   = $db.Count
        Pausendetails   = ($pauseDetails -join ' | ')
        Sessions        = $ds.Count
    })

    # Console output
    Write-Sep
    Write-Host (' {0:dddd, dd. MMMM yyyy}' -f $day) -ForegroundColor White
    Write-Host (' Anwesenheit  : {0:HH:mm} – {1:HH:mm}  ({2})' -f $dayStart, $dayEnd, (Format-Dur $presence)) -ForegroundColor Gray
    Write-Host (' Aktive Arbeit: {0}   |   Pausen: {1}  ({2} Unterbrechung(en))' -f (Format-Dur $workTs), (Format-Dur $breakTs), $db.Count) -ForegroundColor Green
    Write-Host ''

    $timeline = @(
        $ds | Select-Object Start, End, Duration, @{N='Kind';E={'Arbeit'}}, @{N='Reason';E={$null}}
        $db | Select-Object Start, End, Duration, @{N='Kind';E={'Pause'}},  Reason
    ) | Sort-Object Start

    foreach ($ev in $timeline) {
        $color  = if ($ev.Kind -eq 'Arbeit') { 'Cyan' } else { 'DarkYellow' }
        $icon   = if ($ev.Kind -eq 'Arbeit') { '[+]' } else { '[-]' }
        $reason = if ($ev.Reason) { "  ($($ev.Reason))" } else { '' }
        Write-Host ('   {0} {1:HH:mm} – {2:HH:mm}  {3}{4}' -f $icon, $ev.Start, $ev.End, (Format-Dur $ev.Duration), $reason) -ForegroundColor $color
    }
    Write-Host ''
}

# ── Grand total ───────────────────────────────────────────────────────────────
$label = if ($StartDate -ne '' -or $EndDate -ne '') {
    "$($since.ToString('dd.MM.yyyy')) – $($until.ToString('dd.MM.yyyy'))"
} else { "letzte $DaysBack Tage" }

Write-Sep '═'
Write-Host (' Gesamt ({0})' -f $label) -ForegroundColor White
Write-Host (' Aktive Arbeit   : {0}  ({1:N1} h)' -f (Format-Dur $grandWork),  $grandWork.TotalHours)  -ForegroundColor Green
Write-Host (' Pausen gesamt   : {0}  ({1:N1} h)' -f (Format-Dur $grandBreak), $grandBreak.TotalHours) -ForegroundColor DarkYellow
Write-Host (' Arbeitstage     : {0}'             -f $allDays.Count) -ForegroundColor Gray
if ($allDays.Count -gt 0) {
    $avg = $grandWork.TotalHours / $allDays.Count
    Write-Host (' Ø pro Tag       : {0:N1} h' -f $avg) -ForegroundColor Gray
}
Write-Sep '═'

# ── CSV export ────────────────────────────────────────────────────────────────
try {
    $outDir = Split-Path -Path $OutputPath -Parent
    if ($outDir -and -not (Test-Path $outDir)) {
        New-Item -ItemType Directory -Path $outDir -Force | Out-Null
        Write-Host "Verzeichnis erstellt: $outDir" -ForegroundColor Yellow
    }

    # Sheet 1: daily summary
    $summaryPath = $OutputPath -replace '\.csv$', '_Zusammenfassung.csv'
    $daySummaries | Export-Csv -Path $summaryPath -NoTypeInformation -Encoding UTF8 -Delimiter ';'
    Write-Host "Zusammenfassung exportiert : $summaryPath" -ForegroundColor Green

    # Sheet 2: raw event log
    $rawLog | Export-Csv -Path $OutputPath -NoTypeInformation -Encoding UTF8 -Delimiter ';'
    Write-Host "Ereignisprotokoll exportiert: $OutputPath" -ForegroundColor Green

} catch {
    Write-Warning "CSV-Export fehlgeschlagen: $($_.Exception.Message)"
}

Write-Host "`nFertig!`n" -ForegroundColor Green
