# ============================================
# 5. EVENT ANALYSIS - Comprehensive Event Processing
# ============================================
# Purpose: Analyze all events including seizures, spikes, annotations
# Exports events for clinical review

# Load configuration
. "$PSScriptRoot\0-config.ps1"

param(
    [int]$RecordIndex = 0
)

Write-Host "============================================" -ForegroundColor Cyan
Write-Host "5. EVENT ANALYSIS - Comprehensive Processing" -ForegroundColor Cyan
Write-Host "============================================" -ForegroundColor Cyan

# Initialize API
$api = Initialize-ArcApi
if ($null -eq $api) {
    Write-Error "Cannot proceed without API connection"
    exit 1
}

try {
    $records = $api.GetRecords()
    $record = $records.GetAt($RecordIndex)
    
    $openResult = $record.Open()
    if (-not $openResult.IsSuccess) {
        Write-Error "Failed to open record: $($openResult.ErrorMessage)"
        exit 1
    }
    
    $data = $record.Data
    Write-Host "Analyzing events for record: $($record.RecordKey)" -ForegroundColor Yellow

    # ========================================
    # GET ALL EVENTS
    # ========================================
    Write-Host "`n=== ALL EVENTS ===" -ForegroundColor Yellow
    $allEvents = $data.GetEvents()
    $totalEvents = $allEvents.GetCount()
    Write-Host "Total Events: $totalEvents"

    if ($totalEvents -eq 0) {
        Write-Host "No events found in this record" -ForegroundColor Yellow
    }

    # ========================================
    # CATEGORIZE BY TYPE
    # ========================================
    Write-Host "`n=== EVENT SUMMARY BY TYPE ===" -ForegroundColor Yellow
    $eventTypeCount = @{}
    $eventsByType = @{}
    
    for ($i = 0; $i -lt $totalEvents; $i++) {
        $event = $allEvents.GetAt($i)
        $type = $event.EventType
        
        if ($eventTypeCount.ContainsKey($type)) {
            $eventTypeCount[$type]++
            $eventsByType[$type] += @($event)
        } else {
            $eventTypeCount[$type] = 1
            $eventsByType[$type] = @($event)
        }
    }

    foreach ($type in $eventTypeCount.Keys | Sort-Object { $eventTypeCount[$_] } -Descending) {
        $count = $eventTypeCount[$type]
        $color = if ($type -match "Seizure|Spike") { "Red" } elseif ($type -match "Comment") { "Gray" } else { "White" }
        Write-Host "  $type : $count" -ForegroundColor $color
    }

    # ========================================
    # CLINICAL EVENTS (Seizures, Spikes)
    # ========================================
    Write-Host "`n=== CLINICAL EVENTS (Seizures & Spikes) ===" -ForegroundColor Yellow
    $clinicalTypes = @(
        "Persyst SeizureDetected", "Persyst Spike", "Persyst SpikeBurst", 
        "Persyst RhythmicBurst", "Seizure", "Spike", "Sharp Wave"
    )
    
    $clinicalEvents = @()
    try {
        $clinicalEvents = $data.GetEvents(0, [int]::MaxValue, $clinicalTypes)
        Write-Host "Clinical events found: $($clinicalEvents.GetCount())" -ForegroundColor $(if ($clinicalEvents.GetCount() -gt 0) { "Red" } else { "Green" })
    }
    catch {
        Write-Host "Could not filter clinical events: $($_.Exception.Message)" -ForegroundColor Yellow
        # Fallback: manually filter
        for ($i = 0; $i -lt $totalEvents; $i++) {
            $event = $allEvents.GetAt($i)
            if ($clinicalTypes -contains $event.EventType) {
                $clinicalEvents += $event
            }
        }
        Write-Host "Clinical events (manual filter): $($clinicalEvents.Count)"
    }

    # Display clinical events
    if ($clinicalEvents.GetCount -and $clinicalEvents.GetCount() -gt 0) {
        Write-Host "`n--- CLINICAL EVENT DETAILS ---" -ForegroundColor Red
        for ($i = 0; $i -lt [Math]::Min($clinicalEvents.GetCount(), 20); $i++) {
            $event = $clinicalEvents.GetAt($i)
            Write-Host "`nEvent $($i+1): $($event.EventType)" -ForegroundColor Red
            Write-Host "  Offset: $([Math]::Round($event.Offset.TotalSeconds, 2)) sec ($([Math]::Round($event.Offset.TotalMinutes, 2)) min)"
            Write-Host "  Duration: $([Math]::Round($event.Duration.TotalSeconds, 2)) sec"
            Write-Host "  Text: $($event.Text)"
            Write-Host "  Priority: $($event.Priority)"
        }
        if ($clinicalEvents.GetCount() -gt 20) {
            Write-Host "`n... and $($clinicalEvents.GetCount() - 20) more clinical events"
        }
    }

    # ========================================
    # EVENTS IN FIRST HOUR
    # ========================================
    Write-Host "`n=== EVENTS IN FIRST HOUR ===" -ForegroundColor Yellow
    try {
        $firstHourEvents = $data.GetEvents(0, 3600, $null)
        Write-Host "Events in first hour: $($firstHourEvents.GetCount())"
        
        for ($i = 0; $i -lt [Math]::Min($firstHourEvents.GetCount(), 10); $i++) {
            $event = $firstHourEvents.GetAt($i)
            $timeStr = [Math]::Round($event.Offset.TotalMinutes, 2)
            Write-Host "  [$timeStr min] $($event.EventType): $($event.Text)"
        }
    }
    catch {
        Write-Host "Could not retrieve first hour events: $($_.Exception.Message)" -ForegroundColor Yellow
    }

    # ========================================
    # CORTICAL STIMULATION EVENTS
    # ========================================
    Write-Host "`n=== CORTICAL STIMULATION EVENTS ===" -ForegroundColor Yellow
    try {
        $stimEvents = $data.GetCorticalStimEvents()
        Write-Host "Cortical Stim Events: $($stimEvents.GetCount())"
        
        if ($stimEvents.GetCount() -gt 0) {
            for ($i = 0; $i -lt [Math]::Min($stimEvents.GetCount(), 10); $i++) {
                $stim = $stimEvents.GetAt($i)
                Write-Host "`nStim Event $($i+1):" -ForegroundColor Cyan
                Write-Host "  Offset: $([Math]::Round($stim.Offset.TotalSeconds, 2)) sec"
                Write-Host "  Current: $($stim.CurrentDelivered) mA"
                Write-Host "  Response: $($stim.ResponseType) - $($stim.FunctionalResponse)"
                Write-Host "  Body: $($stim.BodySide) $($stim.BodyRegion)"
                Write-Host "  Task: $($stim.PatientTask)"
            }
        }
    }
    catch {
        Write-Host "Cortical stim events not available (may not be applicable)" -ForegroundColor Gray
    }

    # ========================================
    # EXPORT EVENTS
    # ========================================
    $timestamp = Get-Timestamp
    $recordKey = $record.RecordKey -replace '[^a-zA-Z0-9]', '_'

    # All events CSV
    $eventsCsv = "EventNumber,EventType,OffsetSeconds,OffsetMinutes,DurationSeconds,Text,Priority,EventID,CanDelete,IsCreatedByApi`n"
    for ($i = 0; $i -lt $totalEvents; $i++) {
        $event = $allEvents.GetAt($i)
        $text = ($event.Text -replace '"', '""') -replace "`n", " "
        $eventsCsv += "$($i+1),`"$($event.EventType)`",$($event.Offset.TotalSeconds),$([Math]::Round($event.Offset.TotalMinutes, 2)),$($event.Duration.TotalSeconds),`"$text`",$($event.Priority),$($event.EventID),$($event.CanDelete),$($event.IsCreatedByApi)`n"
    }
    Export-SafeCsv -Content $eventsCsv -FileName "events_all_${recordKey}_$timestamp.csv"

    # Clinical events CSV
    if ($clinicalEvents.GetCount -and $clinicalEvents.GetCount() -gt 0) {
        $clinicalCsv = "EventNumber,EventType,OffsetSeconds,OffsetMinutes,DurationSeconds,Text,Priority`n"
        for ($i = 0; $i -lt $clinicalEvents.GetCount(); $i++) {
            $event = $clinicalEvents.GetAt($i)
            $text = ($event.Text -replace '"', '""') -replace "`n", " "
            $clinicalCsv += "$($i+1),`"$($event.EventType)`",$($event.Offset.TotalSeconds),$([Math]::Round($event.Offset.TotalMinutes, 2)),$($event.Duration.TotalSeconds),`"$text`",$($event.Priority)`n"
        }
        Export-SafeCsv -Content $clinicalCsv -FileName "events_clinical_${recordKey}_$timestamp.csv"
    }

    # Event summary
    $summaryLines = @("Event Type,Count")
    foreach ($type in $eventTypeCount.Keys | Sort-Object) {
        $summaryLines += "$type,$($eventTypeCount[$type])"
    }
    Export-SafeCsv -Content ($summaryLines -join "`n") -FileName "events_summary_${recordKey}_$timestamp.csv"

    Write-Host "`n=== EVENT ANALYSIS COMPLETE ===" -ForegroundColor Green
    Write-Host "Total Events: $totalEvents"
    if ($clinicalEvents.GetCount) {
        Write-Host "Clinical Events: $($clinicalEvents.GetCount())" -ForegroundColor $(if ($clinicalEvents.GetCount() -gt 0) { "Red" } else { "Green" })
    }

}
catch {
    Write-Error "Error during event analysis: $($_.Exception.Message)"
    Write-Host "Stack Trace: $($_.ScriptStackTrace)" -ForegroundColor Red
}
finally {
    if ($null -ne $record -and $record.IsOpen) {
        $record.Close()
    }
    Close-ArcApi $api
}
