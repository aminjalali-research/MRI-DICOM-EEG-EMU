$api = New-Object -comobject Arc.Api.ArcApi
$api.Login("username", "password")

$records = $api.GetRecords()
$record = $records.GetAt(0)
$record.Open()
$data = $record.Data

Write-Host "=== WAVEFORM DATA EXTRACTION ==="

# Get discontinuous waveform data (handles gaps properly)
$startOffset = 0      # seconds
$duration = 60        # 60 seconds of data

Write-Host "Extracting $duration seconds starting at offset $startOffset..."

$channelData = $data.GetDiscontinuousWaveformData($startOffset, $duration)
Write-Host "Channels retrieved: " $channelData.GetCount()

for ($i = 0; $i -lt $channelData.GetCount(); $i++) {
    $channel = $channelData.GetAt($i)
    
    Write-Host "`nChannel: $($channel.ChannelName)"
    Write-Host "  Total Samples: $($channel.SampleCount)"
    Write-Host "  Start Offset: $($channel.StartOffset.TotalSeconds) sec"
    Write-Host "  End Offset: $($channel.EndOffset.TotalSeconds) sec"
    Write-Host "  Start Time (UTC): $($channel.StartTime)"
    Write-Host "  End Time (UTC): $($channel.EndTime)"
    Write-Host "  Number of waveform segments: " $channel.WaveformData.GetCount()
    
    # Process each waveform segment
    $waveforms = $channel.WaveformData
    for ($j = 0; $j -lt $waveforms.GetCount(); $j++) {
        $waveform = $waveforms.GetAt($j)
        
        Write-Host "`n  Segment $($j+1):"
        Write-Host "    Samples: $($waveform.SampleCount)"
        Write-Host "    Start Offset: $($waveform.StartOffset.TotalSeconds) sec"
        Write-Host "    End Offset: $($waveform.EndOffset.TotalSeconds) sec"
        Write-Host "    Sample Period: $($waveform.SamplePeriod.TotalMilliseconds) ms"
        Write-Host "    Sample Rate: " (1000.0 / $waveform.SamplePeriod.TotalMilliseconds) " Hz"
        Write-Host "    High Cut Filter: $($waveform.HighCutFilter) Hz"
        Write-Host "    Low Cut Filter: $($waveform.LowCutFilter) Hz"
        Write-Host "    Notch Filter: $($waveform.NotchFilter) Hz"
        
        # Access actual data points
        $dataPoints = $waveform.Data
        Write-Host "    Data array length: $($dataPoints.Length)"
        Write-Host "    First 10 values: $($dataPoints[0..9] -join ', ')"
        Write-Host "    Min value: $([Math]::Round(($dataPoints | Measure-Object -Minimum).Minimum, 2))"
        Write-Host "    Max value: $([Math]::Round(($dataPoints | Measure-Object -Maximum).Maximum, 2))"
        Write-Host "    Mean value: $([Math]::Round(($dataPoints | Measure-Object -Average).Average, 2))"
    }
}

# Export waveform data to CSV for analysis
Write-Host "`n`nExporting waveform data to CSV..."
$csvPath = "C:\Output\waveform_data.csv"

# Create header
$header = "Time,"
for ($i = 0; $i -lt $channelData.GetCount(); $i++) {
    $channel = $channelData.GetAt($i)
    $header += "$($channel.ChannelName),"
}
$header = $header.TrimEnd(',')
$header | Out-File $csvPath

# Write data (simplified - assumes all channels have same sampling)
$firstChannel = $channelData.GetAt(0)
$firstWaveform = $firstChannel.WaveformData.GetAt(0)
$sampleCount = $firstWaveform.SampleCount
$samplePeriod = $firstWaveform.SamplePeriod.TotalSeconds

for ($sample = 0; $sample -lt [Math]::Min($sampleCount, 1000); $sample++) {
    $line = "$($sample * $samplePeriod),"
    
    for ($i = 0; $i -lt $channelData.GetCount(); $i++) {
        $channel = $channelData.GetAt($i)
        $waveform = $channel.WaveformData.GetAt(0)
        $dataPoints = $waveform.Data
        
        if ($sample -lt $dataPoints.Length) {
            $line += "$($dataPoints[$sample]),"
        } else {
            $line += "NA,"
        }
    }
    
    $line = $line.TrimEnd(',')
    $line | Out-File $csvPath -Append
}

Write-Host "CSV export complete: $csvPath"

$record.Close()
$api.Dispose()