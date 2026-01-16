# ========================================
# WAVEFORM DATA EXTRACTION (FIXED)
# With proper error handling and null checks
# ========================================

$api = New-Object -comobject Arc.Api.ArcApi
$api.Login("username", "password")

Write-Host "=== WAVEFORM DATA EXTRACTION ==="

# Get records
$records = $api.GetRecords()
if ($records.GetCount() -eq 0) {
    Write-Host "ERROR: No records found"
    $api.Dispose()
    exit
}

Write-Host "Found $($records.GetCount()) records"

# Select first record (or specify index)
$recordIndex = 0
$record = $records.GetAt($recordIndex)

Write-Host "`nOpening record $($recordIndex + 1)..."

# Open the record
$openResult = $record.Open()
if (-not $openResult.IsSuccess) {
    Write-Host "ERROR: Failed to open record: $($openResult.ErrorMessage)"
    $api.Dispose()
    exit
}

# Check if data is available
$data = $record.Data
if ($null -eq $data) {
    Write-Host "ERROR: Record data is null"
    $record.Close()
    $api.Dispose()
    exit
}

Write-Host "✓ Record opened successfully"
Write-Host "Recording Duration: $($data.RecordingDuration.TotalHours) hours"
Write-Host "Recording Start Time: $($data.RecordingStartTime)"

# ========================================
# CONFIGURE EXTRACTION PARAMETERS
# ========================================

# Start at beginning (in seconds)
$startOffset = 0

# Duration to extract (in seconds)
# Try smaller durations first (10-60 seconds)
$duration = 60

Write-Host "`nExtracting $duration seconds starting at offset $startOffset seconds..."

# ========================================
# EXTRACT WAVEFORM DATA WITH ERROR HANDLING
# ========================================

try {
    $channelData = $data.GetDiscontinuousWaveformData($startOffset, $duration)
    
    # Check if data was returned
    if ($null -eq $channelData) {
        Write-Host "ERROR: GetDiscontinuousWaveformData returned null"
        Write-Host "This may mean:"
        Write-Host "  - No data available at this time offset"
        Write-Host "  - Recording hasn't started yet at this offset"
        Write-Host "  - Data gap at this location"
        $record.Close()
        $api.Dispose()
        exit
    }
    
    $channelCount = $channelData.GetCount()
    Write-Host "✓ Channels retrieved: $channelCount"
    
    if ($channelCount -eq 0) {
        Write-Host "WARNING: No channel data available at this offset"
        $record.Close()
        $api.Dispose()
        exit
    }
    
    # ========================================
    # PROCESS EACH CHANNEL
    # ========================================
    
    for ($i = 0; $i -lt $channelCount; $i++) {
        $channel = $channelData.GetAt($i)
        
        if ($null -eq $channel) {
            Write-Host "`nWARNING: Channel $($i+1) is null, skipping..."
            continue
        }
        
        Write-Host "`n========================================="
        Write-Host "Channel $($i+1): $($channel.ChannelName)"
        Write-Host "========================================="
        Write-Host "  Total Samples: $($channel.SampleCount)"
        
        # Check for null values before accessing
        if ($null -ne $channel.StartOffset) {
            Write-Host "  Start Offset: $($channel.StartOffset.TotalSeconds) sec"
        }
        
        if ($null -ne $channel.EndOffset) {
            Write-Host "  End Offset: $($channel.EndOffset.TotalSeconds) sec"
        }
        
        if ($null -ne $channel.StartTime) {
            Write-Host "  Start Time (UTC): $($channel.StartTime)"
        }
        
        if ($null -ne $channel.EndTime) {
            Write-Host "  End Time (UTC): $($channel.EndTime)"
        }
        
        # ========================================
        # PROCESS WAVEFORM SEGMENTS
        # ========================================
        
        $waveforms = $channel.WaveformData
        
        # Check if WaveformData is null
        if ($null -eq $waveforms) {
            Write-Host "  WARNING: No waveform data available for this channel"
            continue
        }
        
        $waveformCount = $waveforms.GetCount()
        Write-Host "  Number of waveform segments: $waveformCount"
        
        if ($waveformCount -eq 0) {
            Write-Host "  WARNING: WaveformData list is empty"
            continue
        }
        
        # Process each waveform segment
        for ($j = 0; $j -lt $waveformCount; $j++) {
            $waveform = $waveforms.GetAt($j)
            
            if ($null -eq $waveform) {
                Write-Host "`n  WARNING: Waveform segment $($j+1) is null, skipping..."
                continue
            }
            
            Write-Host "`n  --- Segment $($j+1) ---"
            Write-Host "    Samples: $($waveform.SampleCount)"
            
            if ($null -ne $waveform.StartOffset) {
                Write-Host "    Start Offset: $($waveform.StartOffset.TotalSeconds) sec"
            }
            
            if ($null -ne $waveform.EndOffset) {
                Write-Host "    End Offset: $($waveform.EndOffset.TotalSeconds) sec"
            }
            
            if ($null -ne $waveform.SamplePeriod -and $waveform.SamplePeriod.TotalMilliseconds -gt 0) {
                Write-Host "    Sample Period: $($waveform.SamplePeriod.TotalMilliseconds) ms"
                $samplingRate = 1000.0 / $waveform.SamplePeriod.TotalMilliseconds
                Write-Host "    Sample Rate: $([Math]::Round($samplingRate, 2)) Hz"
            }
            
            Write-Host "    High Cut Filter: $($waveform.HighCutFilter) Hz"
            Write-Host "    Low Cut Filter: $($waveform.LowCutFilter) Hz"
            Write-Host "    Notch Filter: $($waveform.NotchFilter) Hz"
            
            # Access actual data points
            $dataPoints = $waveform.Data
            
            if ($null -eq $dataPoints) {
                Write-Host "    WARNING: Data array is null"
                continue
            }
            
            if ($dataPoints.Length -eq 0) {
                Write-Host "    WARNING: Data array is empty"
                continue
            }
            
            Write-Host "    Data array length: $($dataPoints.Length)"
            
            # Show first 10 values
            $firstTen = $dataPoints[0..[Math]::Min(9, $dataPoints.Length - 1)]
            Write-Host "    First 10 values: $($firstTen -join ', ')"
            
            # Calculate statistics
            try {
                $stats = $dataPoints | Measure-Object -Minimum -Maximum -Average
                Write-Host "    Min value: $([Math]::Round($stats.Minimum, 2)) µV"
                Write-Host "    Max value: $([Math]::Round($stats.Maximum, 2)) µV"
                Write-Host "    Mean value: $([Math]::Round($stats.Average, 2)) µV"
                Write-Host "    Peak-to-Peak: $([Math]::Round($stats.Maximum - $stats.Minimum, 2)) µV"
            } catch {
                Write-Host "    WARNING: Could not calculate statistics: $($_.Exception.Message)"
            }
        }
    }
    
    # ========================================
    # EXPORT WAVEFORM DATA TO CSV
    # ========================================
    
    Write-Host "`n`n=== EXPORTING WAVEFORM DATA TO CSV ==="
    
    $outputDir = "C:\Output"
    if (-not (Test-Path $outputDir)) {
        New-Item -ItemType Directory -Path $outputDir -Force | Out-Null
    }
    
    $csvPath = "$outputDir\waveform_data.csv"
    
    # Check if we have valid data to export
    $hasValidData = $false
    foreach ($channel in $channelData) {
        if ($null -ne $channel.WaveformData -and $channel.WaveformData.GetCount() -gt 0) {
            $hasValidData = $true
            break
        }
    }
    
    if (-not $hasValidData) {
        Write-Host "ERROR: No valid waveform data to export"
    } else {
        # Create header
        $header = "Time_Seconds,"
        for ($i = 0; $i -lt $channelData.GetCount(); $i++) {
            $channel = $channelData.GetAt($i)
            if ($null -ne $channel) {
                $header += "$($channel.ChannelName),"
            }
        }
        $header = $header.TrimEnd(',')
        
        # Initialize CSV content
        $csvContent = @($header)
        
        # Get the first valid channel to determine sample count
        $referenceChannel = $null
        $referenceSamplePeriod = 0
        
        for ($i = 0; $i -lt $channelData.GetCount(); $i++) {
            $channel = $channelData.GetAt($i)
            if ($null -ne $channel -and $null -ne $channel.WaveformData -and $channel.WaveformData.GetCount() -gt 0) {
                $waveform = $channel.WaveformData.GetAt(0)
                if ($null -ne $waveform -and $null -ne $waveform.Data -and $waveform.Data.Length -gt 0) {
                    $referenceChannel = $channel
                    if ($null -ne $waveform.SamplePeriod) {
                        $referenceSamplePeriod = $waveform.SamplePeriod.TotalSeconds
                    }
                    break
                }
            }
        }
        
        if ($null -eq $referenceChannel) {
            Write-Host "ERROR: No valid reference channel found"
        } else {
            $firstWaveform = $referenceChannel.WaveformData.GetAt(0)
            $sampleCount = [Math]::Min($firstWaveform.Data.Length, 1000)  # Limit to 1000 samples for CSV
            
            Write-Host "Exporting $sampleCount samples..."
            
            # Write data rows
            for ($sample = 0; $sample -lt $sampleCount; $sample++) {
                if ($sample % 100 -eq 0) {
                    Write-Host "  Progress: $sample / $sampleCount samples"
                }
                
                $time = $sample * $referenceSamplePeriod
                $line = "$([Math]::Round($time, 4)),"
                
                # Add data for each channel
                for ($i = 0; $i -lt $channelData.GetCount(); $i++) {
                    $channel = $channelData.GetAt($i)
                    
                    if ($null -ne $channel -and 
                        $null -ne $channel.WaveformData -and 
                        $channel.WaveformData.GetCount() -gt 0) {
                        
                        $waveform = $channel.WaveformData.GetAt(0)
                        
                        if ($null -ne $waveform -and 
                            $null -ne $waveform.Data -and 
                            $sample -lt $waveform.Data.Length) {
                            
                            $line += "$([Math]::Round($waveform.Data[$sample], 4)),"
                        } else {
                            $line += "NA,"
                        }
                    } else {
                        $line += "NA,"
                    }
                }
                
                $line = $line.TrimEnd(',')
                $csvContent += $line
            }
            
            # Write to file
            $csvContent | Out-File $csvPath -Encoding UTF8
            Write-Host "✓ CSV export complete: $csvPath"
            Write-Host "  Samples exported: $sampleCount"
            Write-Host "  Channels: $($channelData.GetCount())"
        }
    }
    
} catch {
    Write-Host "`nERROR: Exception occurred during waveform extraction"
    Write-Host "Message: $($_.Exception.Message)"
    Write-Host "Stack Trace: $($_.Exception.StackTrace)"
} finally {
    # Always close the record
    if ($record.IsOpen) {
        $record.Close()
        Write-Host "`n✓ Record closed"
    }
    
    $api.Dispose()
    Write-Host "✓ API session disposed"
}

Write-Host "`n=== WAVEFORM EXTRACTION COMPLETE ==="
