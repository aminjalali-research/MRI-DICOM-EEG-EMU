# ============================================
# 4. SEGMENT FILTERING - Advanced Time Segment Processing
# ============================================
# Purpose: Filter, merge, and analyze time segments
# Handles data gaps and ensures quality segments

# Load configuration
. "$PSScriptRoot\0-config.ps1"

param(
    [int]$RecordIndex = 0
)

Write-Host "============================================" -ForegroundColor Cyan
Write-Host "4. SEGMENT FILTERING - Advanced Processing" -ForegroundColor Cyan
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
    Write-Host "Processing record: $($record.RecordKey)" -ForegroundColor Yellow

    # ========================================
    # FILTERING CONFIGURATION
    # ========================================
    $filterConfig = @{
        MinDuration = $script:MIN_SEGMENT_DURATION
        MaxGapToMerge = $script:MAX_GAP_TO_MERGE
        MinStartOffset = 0
        MaxEndOffset = [int]::MaxValue
        MinEndSegmentDuration = $script:MIN_END_SEGMENT_DURATION
    }

    Write-Host "`n=== FILTER CONFIGURATION ===" -ForegroundColor Yellow
    Write-Host "  Minimum Duration: $($filterConfig.MinDuration) seconds"
    Write-Host "  Max Gap to Merge: $($filterConfig.MaxGapToMerge) seconds"
    Write-Host "  Min End Segment Duration: $($filterConfig.MinEndSegmentDuration) seconds"

    $segments = $data.GetTimeSegments()
    Write-Host "`nOriginal segments: $($segments.GetCount())"

    # ========================================
    # STEP 1: Filter by time range
    # ========================================
    Write-Host "`n--- STEP 1: Time Range Filtering ---" -ForegroundColor Cyan
    $timeFilteredSegments = @()
    for ($i = 0; $i -lt $segments.GetCount(); $i++) {
        $segment = $segments.GetAt($i)
        
        if ($segment.EndOffset.TotalSeconds -ge $filterConfig.MinStartOffset -and 
            $segment.StartOffset.TotalSeconds -le $filterConfig.MaxEndOffset) {
            $timeFilteredSegments += $segment
        }
    }
    Write-Host "After time range filtering: $($timeFilteredSegments.Count) segments"

    # ========================================
    # STEP 2: Merge segments with small gaps
    # ========================================
    Write-Host "`n--- STEP 2: Gap Merging ---" -ForegroundColor Cyan
    $mergedSegments = @()
    $mergeCount = 0
    $i = 0

    while ($i -lt $timeFilteredSegments.Count) {
        $currentSegment = $timeFilteredSegments[$i]
        $mergedStart = $currentSegment.StartOffset.TotalSeconds
        $mergedEnd = $currentSegment.EndOffset.TotalSeconds
        
        $j = $i + 1
        while ($j -lt $timeFilteredSegments.Count) {
            $nextSegment = $timeFilteredSegments[$j]
            $gap = $nextSegment.StartOffset.TotalSeconds - $mergedEnd
            
            if ($gap -le $filterConfig.MaxGapToMerge) {
                $mergedEnd = $nextSegment.EndOffset.TotalSeconds
                $mergeCount++
                Write-Host "  Merged segments with $([Math]::Round($gap, 2))s gap" -ForegroundColor Gray
                $j++
            } else {
                break
            }
        }
        
        $mergedSegments += @{
            StartOffset = $mergedStart
            EndOffset = $mergedEnd
            Duration = $mergedEnd - $mergedStart
        }
        
        $i = $j
    }
    Write-Host "After merging: $($mergedSegments.Count) segments (merged $mergeCount gaps)"

    # ========================================
    # STEP 3: Filter by minimum duration
    # ========================================
    Write-Host "`n--- STEP 3: Duration Filtering ---" -ForegroundColor Cyan
    $finalSegments = @()
    $removedShortSegments = @()

    for ($i = 0; $i -lt $mergedSegments.Count; $i++) {
        $segment = $mergedSegments[$i]
        $isLastSegment = ($i -eq $mergedSegments.Count - 1)
        
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
    Write-Host "`n=== FINAL VALID SEGMENTS ===" -ForegroundColor Yellow
    $totalValidDuration = 0
    $totalGaps = 0

    for ($i = 0; $i -lt $finalSegments.Count; $i++) {
        $segment = $finalSegments[$i]
        $totalValidDuration += $segment.Duration
        
        Write-Host "`nSegment $($i+1):" -ForegroundColor Cyan
        Write-Host "  Start: $([Math]::Round($segment.StartOffset, 2)) sec ($([Math]::Round($segment.StartOffset/60, 2)) min)"
        Write-Host "  End: $([Math]::Round($segment.EndOffset, 2)) sec ($([Math]::Round($segment.EndOffset/60, 2)) min)"
        Write-Host "  Duration: $([Math]::Round($segment.Duration, 2)) sec ($([Math]::Round($segment.Duration/60, 2)) min)"
        
        if ($i -lt $finalSegments.Count - 1) {
            $nextSegment = $finalSegments[$i + 1]
            $gapDuration = $nextSegment.StartOffset - $segment.EndOffset
            $totalGaps += $gapDuration
            Write-Host "  Gap to next: $([Math]::Round($gapDuration, 2)) sec" -ForegroundColor Gray
        }
    }

    if ($removedShortSegments.Count -gt 0) {
        Write-Host "`n=== REMOVED SEGMENTS ===" -ForegroundColor Yellow
        $totalRemovedDuration = 0
        foreach ($removed in $removedShortSegments) {
            $totalRemovedDuration += $removed.Duration
            Write-Host "  Segment $($removed.Index): $([Math]::Round($removed.Duration, 2))s - $($removed.Reason)" -ForegroundColor Gray
        }
    } else {
        $totalRemovedDuration = 0
    }

    # ========================================
    # SUMMARY STATISTICS
    # ========================================
    Write-Host "`n=== SUMMARY STATISTICS ===" -ForegroundColor Yellow
    Write-Host "Original Segments: $($segments.GetCount())"
    Write-Host "Final Valid Segments: $($finalSegments.Count)"
    Write-Host "Removed Segments: $($removedShortSegments.Count)"
    Write-Host ""
    Write-Host "Total Valid Duration: $([Math]::Round($totalValidDuration/60, 2)) min ($([Math]::Round($totalValidDuration/3600, 2)) hours)"
    Write-Host "Total Removed Duration: $([Math]::Round($totalRemovedDuration/60, 2)) min"
    Write-Host "Total Gap Duration: $([Math]::Round($totalGaps/60, 2)) min"
    
    $efficiency = if (($totalValidDuration + $totalGaps) -gt 0) { 
        ($totalValidDuration / ($totalValidDuration + $totalGaps)) * 100 
    } else { 100 }
    
    $retention = if (($totalValidDuration + $totalRemovedDuration) -gt 0) { 
        ($totalValidDuration / ($totalValidDuration + $totalRemovedDuration)) * 100 
    } else { 100 }
    
    Write-Host ""
    Write-Host "Recording Efficiency: $([Math]::Round($efficiency, 2))%" -ForegroundColor $(if ($efficiency -ge 90) { "Green" } else { "Yellow" })
    Write-Host "Data Retention Rate: $([Math]::Round($retention, 2))%" -ForegroundColor $(if ($retention -ge 95) { "Green" } else { "Yellow" })

    # ========================================
    # EXPORT RESULTS
    # ========================================
    $timestamp = Get-Timestamp
    $recordKey = $record.RecordKey -replace '[^a-zA-Z0-9]', '_'

    # Valid segments CSV
    $validCsv = "SegmentNumber,StartOffsetSec,EndOffsetSec,DurationSec,DurationMin`n"
    for ($i = 0; $i -lt $finalSegments.Count; $i++) {
        $seg = $finalSegments[$i]
        $validCsv += "$($i+1),$($seg.StartOffset),$($seg.EndOffset),$($seg.Duration),$([Math]::Round($seg.Duration/60, 2))`n"
    }
    Export-SafeCsv -Content $validCsv -FileName "segments_valid_${recordKey}_$timestamp.csv"

    # Removed segments CSV
    if ($removedShortSegments.Count -gt 0) {
        $removedCsv = "SegmentNumber,StartOffsetSec,EndOffsetSec,DurationSec,Reason`n"
        foreach ($rem in $removedShortSegments) {
            $removedCsv += "$($rem.Index),$($rem.StartOffset),$($rem.EndOffset),$($rem.Duration),`"$($rem.Reason)`"`n"
        }
        Export-SafeCsv -Content $removedCsv -FileName "segments_removed_${recordKey}_$timestamp.csv"
    }

    # Summary
    $summary = @"
Segment Filtering Summary
Generated: $(Get-Date)
Record: $($record.RecordKey)
============================================

Configuration:
- Min Duration: $($filterConfig.MinDuration) seconds
- Max Gap to Merge: $($filterConfig.MaxGapToMerge) seconds
- Min End Segment: $($filterConfig.MinEndSegmentDuration) seconds

Results:
- Original Segments: $($segments.GetCount())
- Final Segments: $($finalSegments.Count)
- Removed Segments: $($removedShortSegments.Count)

Metrics:
- Valid Duration: $([Math]::Round($totalValidDuration/3600, 2)) hours
- Recording Efficiency: $([Math]::Round($efficiency, 2))%
- Data Retention: $([Math]::Round($retention, 2))%

Status: SUCCESS
"@
    Export-SafeCsv -Content $summary -FileName "segments_summary_${recordKey}_$timestamp.txt"

    Write-Host "`n=== SEGMENT FILTERING COMPLETE ===" -ForegroundColor Green

}
catch {
    Write-Error "Error during segment filtering: $($_.Exception.Message)"
    Write-Host "Stack Trace: $($_.ScriptStackTrace)" -ForegroundColor Red
}
finally {
    if ($null -ne $record -and $record.IsOpen) {
        $record.Close()
    }
    Close-ArcApi $api
}
