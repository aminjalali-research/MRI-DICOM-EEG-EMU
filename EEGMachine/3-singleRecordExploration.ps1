$api = New-Object -comobject Arc.Api.ArcApi
$api.Login("username", "password")

# Get specific record (or first record for example)
$records = $api.GetRecords()
$record = $records.GetAt(0)

# Open record to access detailed data
$result = $record.Open()
if (-not $result.IsSuccess) {
    Write-Host "Failed to open record: " $result.ErrorMessage
    exit
}

$data = $record.Data

Write-Host "=== DETAILED RECORD ANALYSIS ==="
Write-Host "Recording Start Time (UTC): " $data.RecordingStartTime
Write-Host "Recording Duration: " $data.RecordingDuration.TotalHours " hours"
Write-Host "Time Zone Offset: " $data.RecordingTimeZoneOffset.TotalHours " hours"

# Channel Information
Write-Host "`n=== CHANNEL INFORMATION ==="
$channels = $data.ChannelInformation
Write-Host "Total Channels: " $channels.GetCount()
for ($i = 0; $i -lt $channels.GetCount(); $i++) {
    $channel = $channels.GetAt($i)
    Write-Host "Channel $($channel.ChannelNumber): $($channel.ChannelName)"
    Write-Host "  Sample Period: $($channel.SamplePeriod.TotalMilliseconds) ms"
    Write-Host "  Sample Rate: " (1000.0 / $channel.SamplePeriod.TotalMilliseconds) " Hz"
}



# ========================================
# TIME SEGMENTS WITH FILTERING
# ========================================
Write-Host "`n=== TIME SEGMENTS (Data Continuity with Filtering) ==="

# Configuration: Minimum segment duration to keep (in seconds)
$MIN_SEGMENT_DURATION = 60  # Keep only segments >= 60 seconds (1 minute)
# Adjust this value based on your needs:
# - 10 seconds for very granular data
# - 60 seconds (1 minute) for typical clinical relevance
# - 300 seconds (5 minutes) for more significant segments only

$segments = $data.GetTimeSegments()
Write-Host "Total segments found: " $segments.GetCount()

# Filter segments
$validSegments = @()
$removedSegments = @()

for ($i = 0; $i -lt $segments.GetCount(); $i++) {
    $segment = $segments.GetAt($i)
    $duration = $segment.TotalDuration.TotalSeconds
    
    if ($duration -ge $MIN_SEGMENT_DURATION) {
        $validSegments += $segment
    } else {
        $removedSegments += @{
            Index = $i + 1
            StartOffset = $segment.StartOffset.TotalSeconds
            EndOffset = $segment.EndOffset.TotalSeconds
            Duration = $duration
        }
    }
}

Write-Host "Valid segments (>= $MIN_SEGMENT_DURATION seconds): " $validSegments.Count
Write-Host "Removed short segments: " $removedSegments.Count

# Display valid segments
Write-Host "`n--- VALID CONTINUOUS DATA SEGMENTS ---"
$totalValidDuration = 0
$totalGapDuration = 0

for ($i = 0; $i -lt $validSegments.Count; $i++) {
    $segment = $validSegments[$i]
    $durationMinutes = [Math]::Round($segment.TotalDuration.TotalMinutes, 2)
    $totalValidDuration += $segment.TotalDuration.TotalSeconds
    
    Write-Host "Segment $($i+1):"
    Write-Host "  Start Offset: $($segment.StartOffset.TotalSeconds) sec ($([Math]::Round($segment.StartOffset.TotalMinutes, 2)) min)"
    Write-Host "  End Offset: $($segment.EndOffset.TotalSeconds) sec ($([Math]::Round($segment.EndOffset.TotalMinutes, 2)) min)"
    Write-Host "  Duration: $($segment.TotalDuration.TotalSeconds) sec ($durationMinutes min)"
    
    # Calculate gap to next segment
    if ($i -lt $validSegments.Count - 1) {
        $nextSegment = $validSegments[$i + 1]
        $gapDuration = $nextSegment.StartOffset.TotalSeconds - $segment.EndOffset.TotalSeconds
        $totalGapDuration += $gapDuration
        Write-Host "  Gap to next segment: $gapDuration sec ($([Math]::Round($gapDuration/60, 2)) min)"
    }
}

# Display removed segments
if ($removedSegments.Count -gt 0) {
    Write-Host "`n--- REMOVED SHORT SEGMENTS (<$MIN_SEGMENT_DURATION seconds) ---"
    foreach ($removed in $removedSegments) {
        Write-Host "Segment $($removed.Index): $($removed.Duration) sec (at offset $($removed.StartOffset) sec)"
    }
}

# Summary statistics
Write-Host "`n--- DATA CONTINUITY SUMMARY ---"
Write-Host "Total Valid Segments: $($validSegments.Count)"
Write-Host "Total Removed Segments: $($removedSegments.Count)"
Write-Host "Total Valid Data Duration: $([Math]::Round($totalValidDuration/60, 2)) minutes ($([Math]::Round($totalValidDuration/3600, 2)) hours)"
Write-Host "Total Gap Duration: $([Math]::Round($totalGapDuration/60, 2)) minutes"
Write-Host "Data Continuity Rate: $([Math]::Round(($totalValidDuration / ($totalValidDuration + $totalGapDuration)) * 100, 2))%"

# Calculate removed data
$removedDuration = $removedSegments | ForEach-Object { $_.Duration } | Measure-Object -Sum
if ($removedDuration.Count -gt 0) {
    Write-Host "Total Removed Data Duration: $([Math]::Round($removedDuration.Sum/60, 2)) minutes"
}

# Export valid segments to CSV for further analysis
$segmentsCsv = "SegmentNumber,StartOffsetSeconds,EndOffsetSeconds,DurationSeconds,DurationMinutes`n"
for ($i = 0; $i -lt $validSegments.Count; $i++) {
    $segment = $validSegments[$i]
    $segmentsCsv += "$($i+1),$($segment.StartOffset.TotalSeconds),$($segment.EndOffset.TotalSeconds),$($segment.TotalDuration.TotalSeconds),$([Math]::Round($segment.TotalDuration.TotalMinutes, 2))`n"
}
$segmentsCsv | Out-File "C:\Output\valid_segments.csv"
Write-Host "`nValid segments exported to: C:\Output\valid_segments.csv"

# Common Reference Information
Write-Host "`n=== COMMON REFERENCE DATA ==="
$commonRefs = $data.GetCommonReferenceData()
Write-Host "Reference changes: " $commonRefs.GetCount()
for ($i = 0; $i -lt $commonRefs.GetCount(); $i++) {
    $ref = $commonRefs.GetAt($i)
    Write-Host "At offset $($ref.Offset.TotalSeconds)s ($([Math]::Round($ref.Offset.TotalMinutes, 2)) min): $($ref.ChannelName)"
}

# Montage Information
Write-Host "`n=== MONTAGE INFORMATION ==="
$montages = $data.GetMontages($true)  # true = unique only
Write-Host "Unique Montages: " $montages.GetCount()
for ($i = 0; $i -lt $montages.GetCount(); $i++) {
    $montage = $montages.GetAt($i)
    Write-Host "`nMontage: $($montage.Name)"
    Write-Host "  Used at offset: $($montage.AsViewed.TotalSeconds) sec ($([Math]::Round($montage.AsViewed.TotalMinutes, 2)) min)"
    Write-Host "  Trace Containers: " $montage.TraceContainers.GetCount()
    
    # Show first few traces
    $containers = $montage.TraceContainers
    for ($j = 0; $j -lt [Math]::Min(5, $containers.GetCount()); $j++) {
        $container = $containers.GetAt($j)
        $traces = $container.Traces
        if ($traces.GetCount() -gt 0) {
            $trace = $traces.GetAt(0)
            Write-Host "    Trace: $($trace.Name) | Active: $($trace.Active) | Ref: $($trace.Reference)"
            Write-Host "      Sensitivity: $($trace.Sensitivity) uV/mm | Filters: Lo=$($trace.Locut)Hz Hi=$($trace.Hicut)Hz | Notch: $($trace.NotchOn)"
        }
    }
}

$record.Close()
$api.Dispose()