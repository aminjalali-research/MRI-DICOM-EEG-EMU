# ============================================
# 3. SINGLE RECORD EXPLORATION - Deep Dive
# ============================================
# Purpose: Detailed exploration of a single record
# Includes channels, time segments, montages, references

# Load configuration
. "$PSScriptRoot\0-config.ps1"

# Parameter: Record index (0-based) or can be modified to use RecordKey
param(
    [int]$RecordIndex = 0
)

Write-Host "============================================" -ForegroundColor Cyan
Write-Host "3. SINGLE RECORD EXPLORATION - Deep Dive" -ForegroundColor Cyan
Write-Host "============================================" -ForegroundColor Cyan

# Initialize API
$api = Initialize-ArcApi
if ($null -eq $api) {
    Write-Error "Cannot proceed without API connection"
    exit 1
}

try {
    # Get records
    $records = $api.GetRecords()
    $recordCount = $records.GetCount()
    
    if ($recordCount -eq 0) {
        Write-Host "No records found in system" -ForegroundColor Yellow
        exit 0
    }
    
    if ($RecordIndex -ge $recordCount) {
        Write-Error "Record index $RecordIndex is out of range (0-$($recordCount-1))"
        exit 1
    }
    
    $record = $records.GetAt($RecordIndex)
    Write-Host "Opening record $($RecordIndex + 1) of $recordCount : $($record.RecordKey)" -ForegroundColor Yellow

    # Open record to access detailed data
    $result = $record.Open()
    if (-not $result.IsSuccess) {
        Write-Error "Failed to open record: $($result.ErrorMessage)"
        exit 1
    }
    
    Write-Host "Record opened successfully" -ForegroundColor Green

    $data = $record.Data

    # === RECORDING METADATA ===
    Write-Host "`n=== RECORDING METADATA ===" -ForegroundColor Yellow
    Write-Host "Recording Start Time (UTC): $($data.RecordingStartTime)"
    Write-Host "Recording Duration: $([Math]::Round($data.RecordingDuration.TotalHours, 2)) hours"
    Write-Host "Time Zone Offset: $($data.RecordingTimeZoneOffset.TotalHours) hours"

    # === CHANNEL INFORMATION ===
    Write-Host "`n=== CHANNEL INFORMATION ===" -ForegroundColor Yellow
    $channels = $data.ChannelInformation
    $channelCount = $channels.GetCount()
    Write-Host "Total Channels: $channelCount"
    
    $channelsCsv = "ChannelNumber,ChannelName,SamplePeriodMs,SampleRateHz`n"
    
    for ($i = 0; $i -lt $channelCount; $i++) {
        $channel = $channels.GetAt($i)
        $samplePeriodMs = $channel.SamplePeriod.TotalMilliseconds
        $sampleRateHz = if ($samplePeriodMs -gt 0) { [Math]::Round(1000.0 / $samplePeriodMs, 2) } else { 0 }
        
        Write-Host "  Channel $($channel.ChannelNumber): $($channel.ChannelName) @ $sampleRateHz Hz"
        $channelsCsv += "$($channel.ChannelNumber),$($channel.ChannelName),$samplePeriodMs,$sampleRateHz`n"
    }

    # === TIME SEGMENTS (Data Continuity) ===
    Write-Host "`n=== TIME SEGMENTS (Data Continuity) ===" -ForegroundColor Yellow
    
    $segments = $data.GetTimeSegments()
    $segmentCount = $segments.GetCount()
    Write-Host "Total Segments: $segmentCount"
    Write-Host "Data Gaps: $($segmentCount - 1)"
    
    $segmentsCsv = "SegmentNumber,StartOffsetSec,EndOffsetSec,DurationSec,DurationMin`n"
    $totalDataDuration = 0
    $totalGapDuration = 0
    $continuityRate = 100
    
    # Filter segments by minimum duration
    $validSegments = @()
    $removedSegments = @()
    
    for ($i = 0; $i -lt $segmentCount; $i++) {
        $segment = $segments.GetAt($i)
        $duration = $segment.TotalDuration.TotalSeconds
        $totalDataDuration += $duration
        
        if ($duration -ge $script:MIN_SEGMENT_DURATION) {
            $validSegments += $segment
        } else {
            $removedSegments += @{
                Index = $i + 1
                Start = $segment.StartOffset.TotalSeconds
                End = $segment.EndOffset.TotalSeconds
                Duration = $duration
            }
        }
        
        $segmentsCsv += "$($i+1),$($segment.StartOffset.TotalSeconds),$($segment.EndOffset.TotalSeconds),$duration,$([Math]::Round($duration/60, 2))`n"
        
        # Calculate gap to next segment
        if ($i -lt $segmentCount - 1) {
            $nextSegment = $segments.GetAt($i + 1)
            $gap = $nextSegment.StartOffset.TotalSeconds - $segment.EndOffset.TotalSeconds
            $totalGapDuration += $gap
        }
    }
    
    Write-Host "`nValid Segments (>= $($script:MIN_SEGMENT_DURATION)s): $($validSegments.Count)"
    Write-Host "Removed Short Segments: $($removedSegments.Count)"
    
    # Display valid segments
    Write-Host "`n--- VALID CONTINUOUS DATA SEGMENTS ---" -ForegroundColor Cyan
    for ($i = 0; $i -lt $validSegments.Count; $i++) {
        $seg = $validSegments[$i]
        Write-Host "  Segment $($i+1): $([Math]::Round($seg.StartOffset.TotalMinutes, 2)) - $([Math]::Round($seg.EndOffset.TotalMinutes, 2)) min (Duration: $([Math]::Round($seg.TotalDuration.TotalMinutes, 2)) min)"
    }
    
    # Summary
    Write-Host "`n--- DATA CONTINUITY SUMMARY ---" -ForegroundColor Cyan
    Write-Host "Total Data Duration: $([Math]::Round($totalDataDuration/60, 2)) min ($([Math]::Round($totalDataDuration/3600, 2)) hours)"
    Write-Host "Total Gap Duration: $([Math]::Round($totalGapDuration/60, 2)) min"
    if (($totalDataDuration + $totalGapDuration) -gt 0) {
        $continuityRate = ($totalDataDuration / ($totalDataDuration + $totalGapDuration)) * 100
        Write-Host "Data Continuity Rate: $([Math]::Round($continuityRate, 2))%"
    }

    # === COMMON REFERENCE DATA ===
    Write-Host "`n=== COMMON REFERENCE DATA ===" -ForegroundColor Yellow
    $commonRefs = $data.GetCommonReferenceData()
    Write-Host "Reference Changes: $($commonRefs.GetCount())"
    
    for ($i = 0; $i -lt [Math]::Min($commonRefs.GetCount(), 10); $i++) {
        $ref = $commonRefs.GetAt($i)
        Write-Host "  At $([Math]::Round($ref.Offset.TotalMinutes, 2)) min: $($ref.ChannelName)"
    }
    if ($commonRefs.GetCount() -gt 10) {
        Write-Host "  ... and $($commonRefs.GetCount() - 10) more"
    }

    # === MONTAGE INFORMATION ===
    Write-Host "`n=== MONTAGE INFORMATION ===" -ForegroundColor Yellow
    $montages = $data.GetMontages($true)  # true = unique only
    Write-Host "Unique Montages: $($montages.GetCount())"
    
    for ($i = 0; $i -lt $montages.GetCount(); $i++) {
        $montage = $montages.GetAt($i)
        Write-Host "`n  Montage: $($montage.Name)" -ForegroundColor Cyan
        Write-Host "    Used at offset: $([Math]::Round($montage.AsViewed.TotalMinutes, 2)) min"
        Write-Host "    Trace Containers: $($montage.TraceContainers.GetCount())"
        
        # Show first few traces
        $containers = $montage.TraceContainers
        for ($j = 0; $j -lt [Math]::Min(3, $containers.GetCount()); $j++) {
            $container = $containers.GetAt($j)
            $traces = $container.Traces
            if ($traces.GetCount() -gt 0) {
                $trace = $traces.GetAt(0)
                Write-Host "      Trace: $($trace.Name) | Sensitivity: $($trace.Sensitivity) uV/mm | LoCut: $($trace.Locut)Hz HiCut: $($trace.Hicut)Hz"
            }
        }
    }

    # === EXPORT DATA ===
    $timestamp = Get-Timestamp
    $recordKey = $record.RecordKey -replace '[^a-zA-Z0-9]', '_'
    
    # Export channels
    Export-SafeCsv -Content $channelsCsv -FileName "record_${recordKey}_channels_$timestamp.csv"
    
    # Export segments
    Export-SafeCsv -Content $segmentsCsv -FileName "record_${recordKey}_segments_$timestamp.csv"
    
    # Export summary
    $summary = @"
Single Record Exploration Summary
Generated: $(Get-Date)
============================================

Record Key: $($record.RecordKey)
Recording Start: $($data.RecordingStartTime)
Duration: $([Math]::Round($data.RecordingDuration.TotalHours, 2)) hours

Channels: $channelCount
Time Segments: $segmentCount
Data Gaps: $($segmentCount - 1)
Montages: $($montages.GetCount())

Data Continuity Rate: $([Math]::Round($continuityRate, 2))%

Status: SUCCESS
"@
    Export-SafeCsv -Content $summary -FileName "record_${recordKey}_summary_$timestamp.txt"

    Write-Host "`n=== SINGLE RECORD EXPLORATION COMPLETE ===" -ForegroundColor Green

}
catch {
    Write-Error "Error during single record exploration: $($_.Exception.Message)"
    Write-Host "Stack Trace: $($_.ScriptStackTrace)" -ForegroundColor Red
}
finally {
    # Close record if open
    if ($null -ne $record -and $record.IsOpen) {
        $record.Close()
        Write-Host "Record closed" -ForegroundColor Gray
    }
    Close-ArcApi $api
}
