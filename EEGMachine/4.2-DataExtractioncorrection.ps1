# ========================================
# Need to fix error with the Null record reading:
PS C:\users\nicolete\Desktop\ArcCode> .\4.1-DataExtraction.ps1

Security warning
Run only scripts that you trust. While scripts from the internet can be useful, this script can potentially harm your
computer. If you trust this script, use the Unblock-File cmdlet to allow the script to run without this warning
message. Do you want to run C:\users\nicolete\Desktop\ArcCode\4.1-DataExtraction.ps1?
[D] Do not run  [R] Run once  [S] Suspend  [?] Help (default is "D"): R
You cannot call a method on a null-valued expression.
At C:\users\nicolete\Desktop\ArcCode\4.1-DataExtraction.ps1:68 char:1
+ $record.Open()
+ ~~~~~~~~~~~~~~
    + CategoryInfo          : InvalidOperation: (:) [], RuntimeException
    + FullyQualifiedErrorId : InvokeMethodOnNull

Get-FilteredTimeSegments : Cannot bind argument to parameter 'RecordData' because it is null.
At C:\users\nicolete\Desktop\ArcCode\4.1-DataExtraction.ps1:72 char:55
+ $validSegments = Get-FilteredTimeSegments -RecordData $data -MinDurat ...
+                                                       ~~~~~
    + CategoryInfo          : InvalidData: (:) [Get-FilteredTimeSegments], ParameterBindingValidationException
    + FullyQualifiedErrorId : ParameterArgumentValidationErrorNullNotAllowed,Get-FilteredTimeSegments
# ========================================

# Function to filter time segments
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
    
    if ($null -eq $RecordData) {
        Write-Host "ERROR: RecordData is null" -ForegroundColor Red
        return @()
    }
    
    try {
        $segments = $RecordData.GetTimeSegments()
        
        if ($null -eq $segments) {
            Write-Host "WARNING: GetTimeSegments returned null" -ForegroundColor Yellow
            return @()
        }
        
        if ($segments.GetCount() -eq 0) {
            Write-Host "WARNING: No time segments found" -ForegroundColor Yellow
            return @()
        }
        
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
        
    } catch {
        Write-Host "ERROR in Get-FilteredTimeSegments: $($_.Exception.Message)" -ForegroundColor Red
        return @()
    }
}

# ========================================
# MAIN SCRIPT
# ========================================

Write-Host "=== DATA EXTRACTION WITH FILTERED TIME SEGMENTS ===" -ForegroundColor Cyan

# Initialize API
Write-Host "`nConnecting to Arc API..."
try {
    $api = New-Object -comobject Arc.Api.ArcApi
} catch {
    Write-Host "ERROR: Failed to create Arc API object: $($_.Exception.Message)" -ForegroundColor Red
    Write-Host "Make sure Arc API is installed on this machine" -ForegroundColor Red
    exit
}

# Login
Write-Host "Logging in..."
try {
    $loginResult = $api.Login("username", "password")
    if (-not $loginResult.IsSuccess) {
        Write-Host "ERROR: Login failed: $($loginResult.ErrorMessage)" -ForegroundColor Red
        $api.Dispose()
        exit
    }
    Write-Host "✓ Logged in successfully as: $($api.CurrentUserName)" -ForegroundColor Green
} catch {
    Write-Host "ERROR: Login exception: $($_.Exception.Message)" -ForegroundColor Red
    $api.Dispose()
    exit
}

# Get records
Write-Host "`nRetrieving records..."
try {
    $records = $api.GetRecords()
    
    if ($null -eq $records) {
        Write-Host "ERROR: GetRecords returned null" -ForegroundColor Red
        $api.Dispose()
        exit
    }
    
    $recordCount = $records.GetCount()
    
    if ($recordCount -eq 0) {
        Write-Host "ERROR: No records found in the database" -ForegroundColor Red
        $api.Dispose()
        exit
    }
    
    Write-Host "✓ Found $recordCount records" -ForegroundColor Green
    
} catch {
    Write-Host "ERROR: Failed to retrieve records: $($_.Exception.Message)" -ForegroundColor Red
    $api.Dispose()
    exit
}

# Process each record (or just first one for testing)
$recordsToProcess = 1  # Change to $recordCount to process all records

for ($idx = 0; $idx -lt $recordsToProcess; $idx++) {
    Write-Host "`n=========================================" -ForegroundColor Cyan
    Write-Host "Processing Record $($idx + 1) of $recordsToProcess" -ForegroundColor Cyan
    Write-Host "=========================================" -ForegroundColor Cyan
    
    try {
        # Get record
        $record = $records.GetAt($idx)
        
        if ($null -eq $record) {
            Write-Host "ERROR: Record at index $idx is null" -ForegroundColor Red
            continue
        }
        
        Write-Host "Record Key: $($record.RecordKey)"
        Write-Host "Date Recorded: $($record.DateRecorded)"
        Write-Host "Duration: $($record.Duration.TotalHours) hours"
        
        # Open record
        Write-Host "`nOpening record..."
        $openResult = $record.Open()
        
        if ($null -eq $openResult) {
            Write-Host "ERROR: Open() returned null" -ForegroundColor Red
            continue
        }
        
        if (-not $openResult.IsSuccess) {
            Write-Host "ERROR: Failed to open record: $($openResult.ErrorMessage)" -ForegroundColor Red
            continue
        }
        
        Write-Host "✓ Record opened successfully" -ForegroundColor Green
        
        # Get record data
        $data = $record.Data
        
        if ($null -eq $data) {
            Write-Host "ERROR: Record.Data is null" -ForegroundColor Red
            $record.Close()
            continue
        }
        
        Write-Host "✓ Record data accessed" -ForegroundColor Green
        Write-Host "Recording Start Time: $($data.RecordingStartTime)"
        Write-Host "Recording Duration: $($data.RecordingDuration.TotalHours) hours"
        
        # Get ALL time segments (unfiltered)
        Write-Host "`n--- ALL TIME SEGMENTS ---"
        $allSegments = $data.GetTimeSegments()
        
        if ($null -eq $allSegments) {
            Write-Host "WARNING: GetTimeSegments returned null" -ForegroundColor Yellow
        } elseif ($allSegments.GetCount() -eq 0) {
            Write-Host "WARNING: No time segments found - recording may be empty" -ForegroundColor Yellow
        } else {
            Write-Host "Total segments: $($allSegments.GetCount())"
            
            for ($i = 0; $i -lt $allSegments.GetCount(); $i++) {
                $seg = $allSegments.GetAt($i)
                Write-Host "  Segment $($i+1): $([Math]::Round($seg.StartOffset.TotalSeconds, 2))s - $([Math]::Round($seg.EndOffset.TotalSeconds, 2))s (Duration: $([Math]::Round($seg.TotalDuration.TotalSeconds, 2))s)"
            }
        }
        
        # Get FILTERED time segments
        Write-Host "`n--- FILTERED TIME SEGMENTS ---"
        $validSegments = Get-FilteredTimeSegments -RecordData $data -MinDuration 60 -MaxGapToMerge 10
        
        if ($validSegments.Count -eq 0) {
            Write-Host "WARNING: No valid segments after filtering" -ForegroundColor Yellow
        } else {
            Write-Host "Valid segments after filtering: $($validSegments.Count)" -ForegroundColor Green
            
            foreach ($segment in $validSegments) {
                Write-Host "  Duration: $([Math]::Round($segment.DurationMinutes, 2)) min starting at $([Math]::Round($segment.StartOffsetSeconds/60, 2)) min"
            }
            
            # Calculate statistics
            $totalValidDuration = ($validSegments | Measure-Object -Property DurationSeconds -Sum).Sum
            Write-Host "`nTotal valid duration: $([Math]::Round($totalValidDuration/60, 2)) minutes ($([Math]::Round($totalValidDuration/3600, 2)) hours)" -ForegroundColor Green
        }
        
        # Get channel information
        Write-Host "`n--- CHANNEL INFORMATION ---"
        $channels = $data.ChannelInformation
        
        if ($null -eq $channels) {
            Write-Host "WARNING: ChannelInformation is null" -ForegroundColor Yellow
        } else {
            Write-Host "Total channels: $($channels.GetCount())"
            
            for ($i = 0; $i -lt [Math]::Min(5, $channels.GetCount()); $i++) {
                $channel = $channels.GetAt($i)
                $samplingRate = 1000.0 / $channel.SamplePeriod.TotalMilliseconds
                Write-Host "  $($channel.ChannelName): $([Math]::Round($samplingRate, 2)) Hz"
            }
            
            if ($channels.GetCount() -gt 5) {
                Write-Host "  ... and $($channels.GetCount() - 5) more channels"
            }
        }
        
        # Get events
        Write-Host "`n--- EVENTS ---"
        $events = $data.GetEvents()
        
        if ($null -eq $events) {
            Write-Host "WARNING: GetEvents returned null" -ForegroundColor Yellow
        } else {
            Write-Host "Total events: $($events.GetCount())"
            
            # Count events by type
            $eventTypes = @{}
            for ($i = 0; $i -lt $events.GetCount(); $i++) {
                $event = $events.GetAt($i)
                $type = $event.EventType
                if ($eventTypes.ContainsKey($type)) {
                    $eventTypes[$type]++
                } else {
                    $eventTypes[$type] = 1
                }
            }
            
            foreach ($type in $eventTypes.Keys | Sort-Object) {
                Write-Host "  $type : $($eventTypes[$type])"
            }
        }
        
        # Close record
        Write-Host "`n✓ Closing record..." -ForegroundColor Green
        $record.Close()
        
    } catch {
        Write-Host "`nERROR: Exception occurred: $($_.Exception.Message)" -ForegroundColor Red
        Write-Host "Stack trace: $($_.Exception.StackTrace)" -ForegroundColor Red
        
        if ($null -ne $record -and $record.IsOpen) {
            $record.Close()
        }
    }
}

# Cleanup
Write-Host "`n=========================================" -ForegroundColor Cyan
Write-Host "Disposing API connection..." -ForegroundColor Cyan
$api.Dispose()
Write-Host "✓ Complete" -ForegroundColor Green
Write-Host "=========================================" -ForegroundColor Cyan
