
<# 
    arbeitszeiten_fixed.ps1
    Extracts working times from Windows Event Viewer
    Tracks: Login, Logout, Lock, Unlock, Standby, Wake, Shutdown, Startup
    
    Usage:
    .\arbeitszeiten_fixed.ps1 -StartDate "2024-12-01" -EndDate "2024-12-31"
    .\arbeitszeiten_fixed.ps1 -StartDate "2024-12-01"  (EndDate defaults to today)
    .\arbeitszeiten_fixed.ps1  (defaults to last 7 days)
#>

param(
    [string]$StartDate = "",
    [string]$EndDate = "",
    [string]$OutputPath = "C:\Data\Zeiterfassung\Reports.csv"
)

# --- Parse dates ---
if ($StartDate -ne "") {
    try {
        $StartTime = [DateTime]::ParseExact($StartDate, "yyyy-MM-dd", [System.Globalization.CultureInfo]::InvariantCulture)
        Write-Host "Parsed StartDate: $StartDate -> $($StartTime.ToString('yyyy-MM-dd HH:mm:ss'))" -ForegroundColor Gray
    } catch {
        Write-Error "Invalid StartDate format. Use YYYY-MM-DD (e.g., 2024-12-01). Error: $($_.Exception.Message)"
        exit 1
    }
} else {
    $StartTime = (Get-Date).AddDays(-7)
    Write-Host "Using default StartDate (7 days ago): $($StartTime.ToString('yyyy-MM-dd HH:mm:ss'))" -ForegroundColor Gray
}

if ($EndDate -ne "") {
    try {
        $EndTime = [DateTime]::ParseExact($EndDate, "yyyy-MM-dd", [System.Globalization.CultureInfo]::InvariantCulture).AddDays(1).AddSeconds(-1)
        Write-Host "Parsed EndDate: $EndDate -> $($EndTime.ToString('yyyy-MM-dd HH:mm:ss'))" -ForegroundColor Gray
    } catch {
        Write-Error "Invalid EndDate format. Use YYYY-MM-DD (e.g., 2024-12-31). Error: $($_.Exception.Message)"
        exit 1
    }
} else {
    $EndTime = (Get-Date)
    Write-Host "Using default EndDate (today): $($EndTime.ToString('yyyy-MM-dd HH:mm:ss'))" -ForegroundColor Gray
}

# Validate date range
if ($StartTime -gt $EndTime) {
    Write-Error "StartDate cannot be after EndDate!"
    exit 1
}

$CsvPath = $OutputPath

Write-Host "`n=== Working Time Extraction ===" -ForegroundColor Cyan
Write-Host "Time Range: $($StartTime.ToString('yyyy-MM-dd HH:mm')) to $($EndTime.ToString('yyyy-MM-dd HH:mm'))" -ForegroundColor Cyan
Write-Host "Export Path: $CsvPath`n" -ForegroundColor Cyan

# --- Event IDs to track ---
$eventConfig = @{
    System = @{
        # Kernel-Power events (sleep/wake/power)
        1    = 'Wake from Sleep/Standby (Kernel-Power)' # System resumed from sleep/standby
        42   = 'Entering Sleep/Standby (Kernel-Power)'  # System entering sleep or standby
        507  = 'Standby Ended (Kernel-Power)'           # System exited standby
        105  = 'Power State Change'                     # Power state transition
        109  = 'Kernel Power (109)'                     # Power-related event
        
        # Power-Troubleshooter events
        # Event 1 is already covered above
        
        # User32 events (shutdown/restart)
        1074 = 'Shutdown/Restart Initiated (User32)'    # Planned shutdown/restart
        
        # EventLog events (system start/stop)
        6005 = 'System Startup (EventLog)'              # Event Log service started
        6006 = 'System Shutdown (EventLog)'             # Event Log service stopped
        6008 = 'Unexpected Shutdown (EventLog)'         # System crash/power loss
        6009 = 'System Start (EventLog)'                # System information at startup
        
        # Kernel-General events
        12   = 'System Boot (Kernel-General)'           # Kernel boot
        13   = 'System Shutdown (Kernel-General)'       # Kernel shutdown
        
        # Additional Kernel-Power events
        131  = 'Power Transition'                       # System power state transition
        41   = 'System Rebooted Without Clean Shutdown' # Unexpected reboot
    }
    Security = @{
        4624 = 'Logon'                               # Successful logon
        4625 = 'Logon Failed'                        # Failed logon attempt
        4634 = 'Logoff'                              # Logoff
        4647 = 'User Initiated Logoff'               # User-initiated logoff
        4648 = 'Logon Using Explicit Credentials'    # RunAs logon
        4778 = 'Session Reconnected'                 # RDP/Terminal Services reconnect
        4779 = 'Session Disconnected'                # RDP/Terminal Services disconnect
        4800 = 'Workstation Locked'                  # Screen locked
        4801 = 'Workstation Unlocked'                # Screen unlocked
        4802 = 'Screen Saver Invoked'                # Screensaver started
        4803 = 'Screen Saver Dismissed'              # Screensaver ended
    }
}

# --- Function to get events safely ---
function Get-EventsSafe {
    param(
        [string]$LogName,
        [int[]]$EventIds,
        [DateTime]$Start,
        [DateTime]$End
    )

    Write-Host "Reading $LogName events (IDs: $($EventIds -join ', '))..." -ForegroundColor Yellow
    Write-Host "  Time range: $($Start.ToString('yyyy-MM-dd HH:mm:ss')) to $($End.ToString('yyyy-MM-dd HH:mm:ss'))" -ForegroundColor Gray
    
    $results = @()
    
    foreach ($id in $EventIds) {
        try {
            $filter = @{
                LogName   = $LogName
                Id        = $id
                StartTime = $Start
                EndTime   = $End
            }
            
            $events = Get-WinEvent -FilterHashtable $filter -ErrorAction Stop
            $results += $events
            Write-Host "  Found $($events.Count) events with ID $id" -ForegroundColor Green
            
        } catch [System.Exception] {
            if ($_.Exception.Message -notlike "*No events were found*") {
                Write-Warning "  Error reading Event ID $id from $LogName`: $($_.Exception.Message)"
            } else {
                Write-Host "  No events found with ID $id" -ForegroundColor Gray
            }
        }
    }
    
    Write-Host "  Total events from $LogName`: $($results.Count)" -ForegroundColor Cyan
    return $results
}

# --- Function to extract username ---
function Get-EventUser {
    param($Event)
    
    $user = $null
    
    # Try UserId (SID)
    if ($Event.UserId) {
        try {
            $sid = $Event.UserId
            $ntAccount = (New-Object System.Security.Principal.SecurityIdentifier($sid.Value)).Translate([System.Security.Principal.NTAccount])
            $user = $ntAccount.Value
        } catch { }
    }
    
    # Try to extract from event properties
    if (-not $user -and $Event.Properties) {
        # For Security events, username is usually in Properties[5] (for 4624) or Properties[1] (for others)
        try {
            if ($Event.Id -eq 4624 -and $Event.Properties.Count -gt 5) {
                $user = $Event.Properties[5].Value
            } elseif ($Event.Properties.Count -gt 1) {
                $user = $Event.Properties[1].Value
            }
        } catch { }
    }
    
    # Clean up system accounts
    if ($user -like 'NT AUTHORITY\*' -or $user -eq 'SYSTEM' -or $user -like '*$') {
        $user = $null
    }
    
    return $user
}

# --- Function to map event to working time action ---
function Map-WorkingTimeEvent {
    param($Event, $EventTypeMap)
    
    $eventType = $EventTypeMap[$Event.Id]
    if (-not $eventType) {
        $eventType = "Event $($Event.Id)"
    }
    
    $user = Get-EventUser -Event $Event
    
    # Determine if this is a "work start" or "work end" event
    $action = switch ($Event.Id) {
        # Work START events
        { $_ -in 1, 507, 6005, 6009, 12, 4624, 4648, 4778, 4801, 4803 } { 'START' }
        # Work END events
        { $_ -in 42, 105, 109, 131, 6006, 13, 1074, 4634, 4647, 4779, 4800, 4802 } { 'END' }
        # Unexpected events
        { $_ -in 6008, 41 } { 'UNEXPECTED' }
        # Failed attempts
        { $_ -in 4625 } { 'FAILED' }
        default { 'INFO' }
    }
    
    [PSCustomObject]@{
        Timestamp    = $Event.TimeCreated
        Date         = $Event.TimeCreated.ToString('yyyy-MM-dd')
        Time         = $Event.TimeCreated.ToString('HH:mm:ss')
        Action       = $action
        EventType    = $eventType
        EventID      = $Event.Id
        LogName      = $Event.LogName
        Source       = $Event.ProviderName
        User         = $user
        Computer     = $Event.MachineName
    }
}

# --- Collect all events ---
Write-Host "`nCollecting events..." -ForegroundColor Green

$allEvents = @()

# Get System events
$systemIds = $eventConfig.System.Keys | ForEach-Object { [int]$_ }
$systemEvents = Get-EventsSafe -LogName 'System' -EventIds $systemIds -Start $StartTime -End $EndTime
$allEvents += $systemEvents | ForEach-Object { Map-WorkingTimeEvent -Event $_ -EventTypeMap $eventConfig.System }

# Get Security events
$securityIds = $eventConfig.Security.Keys | ForEach-Object { [int]$_ }
$securityEvents = Get-EventsSafe -LogName 'Security' -EventIds $securityIds -Start $StartTime -End $EndTime
$allEvents += $securityEvents | ForEach-Object { Map-WorkingTimeEvent -Event $_ -EventTypeMap $eventConfig.Security }

# Sort by timestamp
$sortedEvents = $allEvents | Sort-Object Timestamp

Write-Host "`nTotal events found: $($sortedEvents.Count)" -ForegroundColor Green

# --- Display results ---
if ($sortedEvents.Count -gt 0) {
    Write-Host "`nEvents by day:" -ForegroundColor Cyan
    
    $sortedEvents | Group-Object Date | ForEach-Object {
        Write-Host "`n--- $($_.Name) ---" -ForegroundColor Yellow
        $_.Group | Format-Table Time, Action, EventType, User -AutoSize
    }
    
    # --- Calculate working time summary with pause detection ---
    Write-Host "`nWorking Time Summary:" -ForegroundColor Cyan
    
    $workingDays = $sortedEvents | Group-Object Date | ForEach-Object {
        $dayEvents = $_.Group | Sort-Object Timestamp
        
        # Find first START event (work begins)
        $firstStart = ($dayEvents | Where-Object { $_.Action -eq 'START' } | Select-Object -First 1)
        
        # Find last END event (work ends)
        # If no END event (shutdown/logoff), use last standby or last lock as end time
        $lastEnd = ($dayEvents | Where-Object { $_.Action -eq 'END' } | Select-Object -Last 1)
        if (-not $lastEnd) {
            # No explicit END event, look for last standby or lock
            $lastEnd = ($dayEvents | Where-Object { $_.EventID -in @(42, 4800) } | Select-Object -Last 1)
        }
        
        $totalTime = $null
        $pauseTime = 0
        $activeTime = $null
        $pauseDetails = @()
        
        if ($firstStart -and $lastEnd) {
            # Calculate total time span
            $totalTime = ($lastEnd.Timestamp - $firstStart.Timestamp).TotalHours
            
            # Calculate pause times (standby, locked, logoff periods)
            $pauseStart = $null
            $pauseType = $null
            
            foreach ($evt in $dayEvents) {
                # Detect pause START (entering standby, lock, or logoff)
                if ($evt.Action -eq 'END' -and $evt.EventID -in @(42, 105, 109, 131, 4800, 4802, 4634, 4647)) {
                    if (-not $pauseStart) {
                        $pauseStart = $evt.Timestamp
                        $pauseType = $evt.EventType
                    }
                }
                # Detect pause END (wake up, unlock, login)
                elseif ($evt.Action -eq 'START' -and $evt.EventID -in @(1, 507, 4801, 4803, 4624)) {
                    if ($pauseStart) {
                        $pauseDuration = ($evt.Timestamp - $pauseStart).TotalHours
                        $pauseTime += $pauseDuration
                        $pauseDetails += [PSCustomObject]@{
                            Start    = $pauseStart.ToString('HH:mm:ss')
                            End      = $evt.Timestamp.ToString('HH:mm:ss')
                            Duration = "{0:N0} min" -f ($pauseDuration * 60)
                            Type     = $pauseType
                        }
                        $pauseStart = $null
                        $pauseType = $null
                    }
                }
            }
            
            # If there's an unclosed pause (still in standby/locked at end of day), count it as pause time
            if ($pauseStart -and $lastEnd) {
                $pauseDuration = ($lastEnd.Timestamp - $pauseStart).TotalHours
                $pauseTime += $pauseDuration
                $pauseDetails += [PSCustomObject]@{
                    Start    = $pauseStart.ToString('HH:mm:ss')
                    End      = $lastEnd.Timestamp.ToString('HH:mm:ss')
                    Duration = "{0:N0} min" -f ($pauseDuration * 60)
                    Type     = "$pauseType (unclosed)"
                }
            }
            
            # Calculate active working time
            $activeTime = $totalTime - $pauseTime
        }
        
        [PSCustomObject]@{
            Date         = $_.Name
            FirstStart   = if ($firstStart) { $firstStart.Time } else { '-' }
            LastEnd      = if ($lastEnd) { $lastEnd.Time + " (" + $lastEnd.EventType + ")" } else { '-' }
            TotalTime    = if ($totalTime) { "{0:N2} h" -f $totalTime } else { '-' }
            PauseTime    = if ($pauseTime -gt 0) { "{0:N2} h" -f $pauseTime } else { '0.00 h' }
            ActiveTime   = if ($activeTime) { "{0:N2} h" -f $activeTime } else { '-' }
            Events       = $dayEvents.Count
            Pauses       = $pauseDetails
        }
    }
    
    # Display summary table
    $workingDays | Select-Object Date, FirstStart, LastEnd, TotalTime, PauseTime, ActiveTime, Events | Format-Table -AutoSize
    
    # Display detailed pause information for each day
    Write-Host "`nDetailed Pause Information:" -ForegroundColor Cyan
    foreach ($day in $workingDays) {
        if ($day.Pauses.Count -gt 0) {
            Write-Host "`n--- $($day.Date) ---" -ForegroundColor Yellow
            $day.Pauses | Format-Table Start, End, Duration, Type -AutoSize
        }
    }
    
    # Calculate totals
    $totalActive = ($workingDays | Where-Object { $_.ActiveTime -ne '-' } | ForEach-Object { 
        [double]($_.ActiveTime -replace ' h', '') 
    } | Measure-Object -Sum).Sum
    
    $totalPause = ($workingDays | Where-Object { $_.PauseTime -ne '-' } | ForEach-Object { 
        [double]($_.PauseTime -replace ' h', '') 
    } | Measure-Object -Sum).Sum
    
    Write-Host "`nTotals:" -ForegroundColor Cyan
    Write-Host "  Total Active Working Time: $("{0:N2}" -f $totalActive) hours" -ForegroundColor Green
    Write-Host "  Total Pause Time: $("{0:N2}" -f $totalPause) hours" -ForegroundColor Yellow
    Write-Host "  Average per Day: $("{0:N2}" -f ($totalActive / $workingDays.Count)) hours" -ForegroundColor Green
    
    # --- Export to CSV ---
    try {
        # Ensure directory exists
        $outputDir = Split-Path -Path $CsvPath -Parent
        if ($outputDir -and -not (Test-Path $outputDir)) {
            Write-Host "Creating directory: $outputDir" -ForegroundColor Yellow
            New-Item -ItemType Directory -Path $outputDir -Force | Out-Null
        }
        
        $sortedEvents | Export-Csv -Path $CsvPath -NoTypeInformation -Encoding UTF8
        Write-Host "`nCSV exported successfully: $CsvPath" -ForegroundColor Green
    } catch {
        Write-Warning "CSV export failed: $($_.Exception.Message)"
    }
    
} else {
    Write-Host "`nNo events found in the specified time range." -ForegroundColor Red
    Write-Host "This could mean:" -ForegroundColor Yellow
    Write-Host "  1. The Event Logs don't go back that far" -ForegroundColor Yellow
    Write-Host "  2. The logs have been cleared" -ForegroundColor Yellow
    Write-Host "  3. You may need to run PowerShell as Administrator" -ForegroundColor Yellow
}

Write-Host "`nDone!`n" -ForegroundColor Green
