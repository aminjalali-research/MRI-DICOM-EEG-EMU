# ========================================
# WAVEFORM DATA DIAGNOSTIC SCRIPT
# Use this to troubleshoot data availability
# ========================================

$api = New-Object -comobject Arc.Api.ArcApi
$api.Login("username", "password")

Write-Host "=== WAVEFORM DATA DIAGNOSTIC ==="

$records = $api.GetRecords()
$record = $records.GetAt(0)

Write-Host "`nOpening record..."
$record.Open()
$data = $record.Data

# Check basic info
Write-Host "`n--- BASIC INFORMATION ---"
Write-Host "Recording Start Time: $($data.RecordingStartTime)"
Write-Host "Recording Duration: $($data.RecordingDuration.TotalHours) hours ($($data.RecordingDuration.TotalSeconds) seconds)"
Write-Host "Time Zone Offset: $($data.RecordingTimeZoneOffset.TotalHours) hours"

# Check time segments (where data actually exists)
Write-Host "`n--- TIME SEGMENTS (Data Availability) ---"
$segments = $data.GetTimeSegments()
Write-Host "Total Segments: $($segments.GetCount())"

if ($segments.GetCount() -eq 0) {
    Write-Host "ERROR: No time segments found - no data available in this record"
} else {
    for ($i = 0; $i -lt $segments.GetCount(); $i++) {
        $segment = $segments.GetAt($i)
        Write-Host "`nSegment $($i+1):"
        Write-Host "  Start: $($segment.StartOffset.TotalSeconds) sec ($([Math]::Round($segment.StartOffset.TotalMinutes, 2)) min)"
        Write-Host "  End: $($segment.EndOffset.TotalSeconds) sec ($([Math]::Round($segment.EndOffset.TotalMinutes, 2)) min)"
        Write-Host "  Duration: $($segment.TotalDuration.TotalSeconds) sec ($([Math]::Round($segment.TotalDuration.TotalMinutes, 2)) min)"
    }
    
    # Suggest valid offset ranges
    $firstSegment = $segments.GetAt(0)
    Write-Host "`n--- SUGGESTED PARAMETERS ---"
    Write-Host "Use these values for waveform extraction:"
    Write-Host "  startOffset = $([Math]::Floor($firstSegment.StartOffset.TotalSeconds))"
    Write-Host "  duration = 60  (or less)"
    Write-Host "`nExample:"
    Write-Host '  $startOffset = ' + [Math]::Floor($firstSegment.StartOffset.TotalSeconds)
    Write-Host '  $duration = 60'
}

# Check channel information
Write-Host "`n--- CHANNEL INFORMATION ---"
$channels = $data.ChannelInformation
Write-Host "Total Channels: $($channels.GetCount())"

for ($i = 0; $i -lt [Math]::Min(5, $channels.GetCount()); $i++) {
    $channel = $channels.GetAt($i)
    Write-Host "  Channel $($i+1): $($channel.ChannelName) @ $([Math]::Round(1000.0/$channel.SamplePeriod.TotalMilliseconds, 2)) Hz"
}

# Try extracting a small sample
Write-Host "`n--- TESTING DATA EXTRACTION ---"

if ($segments.GetCount() -gt 0) {
    $testSegment = $segments.GetAt(0)
    $testOffset = [Math]::Floor($testSegment.StartOffset.TotalSeconds)
    $testDuration = 1  # Just 1 second
    
    Write-Host "Testing extraction at offset $testOffset for $testDuration second..."
    
    try {
        $testData = $data.GetDiscontinuousWaveformData($testOffset, $testDuration)
        
        if ($null -eq $testData) {
            Write-Host "✗ FAILED: GetDiscontinuousWaveformData returned null"
        } else {
            Write-Host "✓ SUCCESS: Retrieved $($testData.GetCount()) channels"
            
            if ($testData.GetCount() -gt 0) {
                $testChannel = $testData.GetAt(0)
                
                if ($null -ne $testChannel) {
                    Write-Host "  Channel: $($testChannel.ChannelName)"
                    Write-Host "  Samples: $($testChannel.SampleCount)"
                    
                    if ($null -ne $testChannel.WaveformData) {
                        Write-Host "  Waveform segments: $($testChannel.WaveformData.GetCount())"
                        
                        if ($testChannel.WaveformData.GetCount() -gt 0) {
                            $testWaveform = $testChannel.WaveformData.GetAt(0)
                            if ($null -ne $testWaveform -and $null -ne $testWaveform.Data) {
                                Write-Host "  Data points: $($testWaveform.Data.Length)"
                                Write-Host "  ✓ DATA EXTRACTION WORKING!"
                            } else {
                                Write-Host "  ✗ Waveform.Data is null"
                            }
                        } else {
                            Write-Host "  ✗ No waveform segments in WaveformData"
                        }
                    } else {
                        Write-Host "  ✗ WaveformData is null"
                    }
                } else {
                    Write-Host "  ✗ Channel object is null"
                }
            }
        }
    } catch {
        Write-Host "✗ EXCEPTION: $($_.Exception.Message)"
    }
} else {
    Write-Host "Cannot test extraction - no time segments available"
}

$record.Close()
$api.Dispose()

Write-Host "`n=== DIAGNOSTIC COMPLETE ==="
Write-Host "If you see '✓ DATA EXTRACTION WORKING!' above, use those parameters in your extraction script"
