function Get-FilteredTimeSegments {
    param(
        [Parameter(Mandatory=$true)]
        $RecordData,
        
        [Parameter(Mandatory=$false)]
        [int]$MinDuration = 60,  # seconds
        
        [Parameter(Mandatory=$false)]
        [int]$MaxGapToMerge = 10,  # seconds
        
        [Parameter(Mandatory=$false)]
        [int]$MinEndSegmentDuration = 120  # seconds
    )
    
    $segments = $RecordData.GetTimeSegments()
    
    # Step 1: Merge segments with small gaps
    $mergedSegments = @()
    $i = 0
    
    while ($i -lt $segments.GetCount()) {
        $currentSegment = $segments.GetAt($i)
        $mergedStart = $currentSegment.StartOffset.TotalSeconds
        $mergedEnd = $currentSegment.EndOffset.TotalSeconds
        
        $j = $i + 1
        while ($j -lt $segments.GetCount()) {
            $nextSegment = $segments.GetAt($j)
            $gap = $nextSegment.StartOffset.TotalSeconds - $mergedEnd
            
            if ($gap -le $MaxGapToMerge) {
                $mergedEnd = $nextSegment.EndOffset.TotalSeconds
                $j++
            } else {
                break
            }
        }
        
        $mergedSegments += @{
            StartOffsetSeconds = $mergedStart
            EndOffsetSeconds = $mergedEnd
            DurationSeconds = $mergedEnd - $mergedStart
            DurationMinutes = ($mergedEnd - $mergedStart) / 60
        }
        
        $i = $j
    }
    
    # Step 2: Filter by duration
    $finalSegments = @()
    
    for ($i = 0; $i -lt $mergedSegments.Count; $i++) {
        $segment = $mergedSegments[$i]
        $isLastSegment = ($i -eq $mergedSegments.Count - 1)
        
        $minDuration = if ($isLastSegment) { $MinEndSegmentDuration } else { $MinDuration }
        
        if ($segment.DurationSeconds -ge $minDuration) {
            $finalSegments += $segment
        }
    }
    
    return $finalSegments
}

# Usage in main extraction pipeline:
$record.Open()
$data = $record.Data

# Get filtered segments
$validSegments = Get-FilteredTimeSegments -RecordData $data -MinDuration 60 -MaxGapToMerge 10

Write-Host "Valid segments: $($validSegments.Count)"
foreach ($segment in $validSegments) {
    Write-Host "  $([Math]::Round($segment.DurationMinutes, 2)) min starting at $([Math]::Round($segment.StartOffsetSeconds/60, 2)) min"
}
