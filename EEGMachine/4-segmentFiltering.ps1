# ========================================
# ADVANCED TIME SEGMENT FILTERING
# ========================================

$api = New-Object -comobject Arc.Api.ArcApi
$api.Login("username", "password")

$records = $api.GetRecords()
$record = $records.GetAt(0)
$record.Open()
$data = $record.Data

Write-Host "=== ADVANCED TIME SEGMENT FILTERING ==="

# ========================================
# FILTERING CONFIGURATION
# ========================================
$filterConfig = @{
    # Minimum duration (seconds) - segments shorter than this are removed
    MinDuration = 60  # 1 minute
    
    # Maximum gap tolerance (seconds) - merge segments separated by small gaps
    MaxGapToMerge = 10  # 10 seconds
    
    # Minimum start offset (seconds) - ignore segments before this time
    MinStartOffset = 0  # Start from beginning
    
    # Maximum end offset (seconds) - ignore segments after this time
    MaxEndOffset = [int]::MaxValue  # No upper limit
    
    # Remove segments at the very end if shorter than this (seconds)
    MinEndSegmentDuration = 120  # Last segment must be at least 2 minutes
}

Write-Host "Filter Configuration:"
Write-Host "  Minimum Duration: $($filterConfig.MinDuration) seconds"
Write-Host "  Max Gap to Merge: $($filterConfig.MaxGapToMerge) seconds"
Write-Host "  Time Range: $($filterConfig.MinStartOffset) - $($filterConfig.MaxEndOffset) seconds"

$segments = $data.GetTimeSegments()
Write-Host "`nOriginal segments: $($segments.GetCount())"

# ========================================
# STEP 1: Filter by time range
# ========================================
$timeFilteredSegments = @()
for ($i = 0; $i -lt $segments.GetCount(); $i++) {
    $segment = $segments.GetAt($i)
    
    # Check if segment is within time range
    if ($segment.EndOffset.TotalSeconds -ge $filterConfig.MinStartOffset -and 
        $segment.StartOffset.TotalSeconds -le $filterConfig.MaxEndOffset) {
        $timeFilteredSegments += $segment
    }
}
Write-Host "After time range filtering: $($timeFilteredSegments.Count) segments"

# ========================================
# STEP 2: Merge segments with small gaps
# ========================================
$mergedSegments = @()
$i = 0

while ($i -lt $timeFilteredSegments.Count) {
    $currentSegment = $timeFilteredSegments[$i]
    $mergedStart = $currentSegment.StartOffset.TotalSeconds
    $mergedEnd = $currentSegment.EndOffset.TotalSeconds
    
    # Look ahead to merge consecutive segments with small gaps
    $j = $i + 1
    while ($j -lt $timeFilteredSegments.Count) {
        $nextSegment = $timeFilteredSegments[$j]
        $gap = $nextSegment.StartOffset.TotalSeconds - $mergedEnd
        
        if ($gap -le $filterConfig.MaxGapToMerge) {
            # Merge this segment
            $mergedEnd = $nextSegment.EndOffset.TotalSeconds
            Write-Host "  Merging segments: gap of $gap seconds merged"
            $j++
        } else {
            break
        }
    }
    
    # Create merged segment object
    $mergedSegments += @{
        StartOffset = $mergedStart
        EndOffset = $mergedEnd
        Duration = $mergedEnd - $mergedStart
    }
    
    $i = $j
}
Write-Host "After merging small gaps: $($mergedSegments.Count) segments"

# ========================================
# STEP 3: Filter by minimum duration
# ========================================
$finalSegments = @()
$removedShortSegments = @()
$isLastSegment = $false

for ($i = 0; $i -lt $mergedSegments.Count; $i++) {
    $segment = $mergedSegments[$i]
    $isLastSegment = ($i -eq $mergedSegments.Count - 1)
    
    # Apply different criteria for last segment
    $minDuration = if ($isLastSegment) { $filterConfig.MinEndSegmentDuration } else { $filterConfig.MinDuration }
    
    if ($segment.Duration -ge $minDuration) {
        $finalSegments += $segment
    } else {
        $removedShortSegments += @{
            Index = $i + 1
            StartOffset = $segment.StartOffset
            EndOffset = $segment.EndOffset
            Duration = $segment.Duration
            Reason = if ($isLastSegment) { "Last segment too short" } else { "Duration below minimum" }
        }
    }
}

Write-Host "After duration filtering: $($finalSegments.Count) segments"
Write-Host "Removed short segments: $($removedShortSegments.Count)"

# ========================================
# DISPLAY RESULTS
# ========================================
Write-Host "`n=== FINAL VALID SEGMENTS ==="
$totalValidDuration = 0
$totalGaps = 0

for ($i = 0; $i -lt $finalSegments.Count; $i++) {
    $segment = $finalSegments[$i]
    $totalValidDuration += $segment.Duration
    
    Write-Host "`nSegment $($i+1):"
    Write-Host "  Start: $([Math]::Round($segment.StartOffset, 2)) sec ($([Math]::Round($segment.StartOffset/60, 2)) min)"
    Write-Host "  End: $([Math]::Round($segment.EndOffset, 2)) sec ($([Math]::Round($segment.EndOffset/60, 2)) min)"
    Write-Host "  Duration: $([Math]::Round($segment.Duration, 2)) sec ($([Math]::Round($segment.Duration/60, 2)) min)"
    
    # Calculate gap to next segment
    if ($i -lt $finalSegments.Count - 1) {
        $nextSegment = $finalSegments[$i + 1]
        $gapDuration = $nextSegment.StartOffset - $segment.EndOffset
        $totalGaps += $gapDuration
        Write-Host "  Gap to next: $([Math]::Round($gapDuration, 2)) sec ($([Math]::Round($gapDuration/60, 2)) min)"
    }
}

Write-Host "`n=== REMOVED SEGMENTS ==="
$totalRemovedDuration = 0
foreach ($removed in $removedShortSegments) {
    $totalRemovedDuration += $removed.Duration
    Write-Host "Segment $($removed.Index): $([Math]::Round($removed.Duration, 2)) sec at offset $([Math]::Round($removed.StartOffset, 2)) sec - Reason: $($removed.Reason)"
}

Write-Host "`n=== SUMMARY STATISTICS ==="
Write-Host "Original Segments: $($segments.GetCount())"
Write-Host "Final Valid Segments: $($finalSegments.Count)"
Write-Host "Removed Segments: $($removedShortSegments.Count)"
Write-Host ""
Write-Host "Total Valid Duration: $([Math]::Round($totalValidDuration/60, 2)) min ($([Math]::Round($totalValidDuration/3600, 2)) hours)"
Write-Host "Total Removed Duration: $([Math]::Round($totalRemovedDuration/60, 2)) min"
Write-Host "Total Gap Duration: $([Math]::Round($totalGaps/60, 2)) min"
Write-Host ""
Write-Host "Recording Efficiency: $([Math]::Round(($totalValidDuration / ($totalValidDuration + $totalGaps)) * 100, 2))%"
Write-Host "Data Retention Rate: $([Math]::Round(($totalValidDuration / ($totalValidDuration + $totalRemovedDuration)) * 100, 2))%"

# ========================================
# EXPORT RESULTS
# ========================================
Write-Host "`n=== EXPORTING RESULTS ==="

# Export valid segments
$validSegmentsCsv = "SegmentNumber,StartOffsetSec,EndOffsetSec,DurationSec,DurationMin,StartOffsetMin,EndOffsetMin`n"
for ($i = 0; $i -lt $finalSegments.Count; $i++) {
    $segment = $finalSegments[$i]
    $validSegmentsCsv += "$($i+1),$($segment.StartOffset),$($segment.EndOffset),$($segment.Duration),$([Math]::Round($segment.Duration/60, 2)),$([Math]::Round($segment.StartOffset/60, 2)),$([Math]::Round($segment.EndOffset/60, 2))`n"
}
$validSegmentsCsv | Out-File "C:\Output\valid_segments_filtered.csv"
Write-Host "Valid segments: C:\Output\valid_segments_filtered.csv"

# Export removed segments
$removedSegmentsCsv = "SegmentNumber,StartOffsetSec,EndOffsetSec,DurationSec,Reason`n"
foreach ($removed in $removedShortSegments) {
    $removedSegmentsCsv += "$($removed.Index),$($removed.StartOffset),$($removed.EndOffset),$($removed.Duration),`"$($removed.Reason)`"`n"
}
$removedSegmentsCsv | Out-File "C:\Output\removed_segments.csv"
Write-Host "Removed segments: C:\Output\removed_segments.csv"

# Export summary
$summary = @"
Time Segment Filtering Summary
==============================
Filter Configuration:
- Minimum Duration: $($filterConfig.MinDuration) seconds
- Max Gap to Merge: $($filterConfig.MaxGapToMerge) seconds
- Min End Segment Duration: $($filterConfig.MinEndSegmentDuration) seconds

Results:
- Original Segments: $($segments.GetCount())
- Final Valid Segments: $($finalSegments.Count)
- Removed Segments: $($removedShortSegments.Count)

Duration Analysis:
- Total Valid Duration: $([Math]::Round($totalValidDuration/3600, 2)) hours
- Total Removed Duration: $([Math]::Round($totalRemovedDuration/60, 2)) minutes
- Total Gap Duration: $([Math]::Round($totalGaps/60, 2)) minutes

Quality Metrics:
- Recording Efficiency: $([Math]::Round(($totalValidDuration / ($totalValidDuration + $totalGaps)) * 100, 2))%
- Data Retention Rate: $([Math]::Round(($totalValidDuration / ($totalValidDuration + $totalRemovedDuration)) * 100, 2))%
"@

$summary | Out-File "C:\Output\filtering_summary.txt"
Write-Host "Summary: C:\Output\filtering_summary.txt"

$record.Close()
$api.Dispose()

Write-Host "`n=== FILTERING COMPLETE ==="