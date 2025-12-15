# Cadwell API scripts implementations 

## 1. Data Exploration Script
```powershell

# Initialize API
$api = New-Object -comobject Arc.Api.ArcApi
$api.Login("username", "password")

Write-Host "=== SYSTEM OVERVIEW ==="
Write-Host "API Version: " $api.Version
Write-Host "Current User: " $api.CurrentUserName
Write-Host "Total Active Records: " $api.GetRecordCount()

# Get record counts by year
$recordsByYear = $api.GetRecordCountsByYear()
Write-Host "`n=== RECORDS BY YEAR ==="
for ($i = 0; $i -lt $recordsByYear.GetCount(); $i++) {
    $yearData = $recordsByYear.GetAt($i)
    Write-Host "Year $($yearData.Year): $($yearData.RecordCount) records"
}

# Check licensed features
Write-Host "`n=== LICENSED FEATURES ==="
$eegFeatures = @(
    "Video", "Highlights", "RemoteControl", "Sentinel", 
    "Oximetry", "Wireless", "ArtifactReductionReview"
)
foreach ($feature in $eegFeatures) {
    $isLicensed = $api.IsEegFeatureLicensed([Arc.Api.LicenseFeatureEEG]::$feature)
    Write-Host "$feature : $isLicensed"
}

$api.Dispose()
```
## 2. Record Exploration
```powershell
$api = New-Object -comobject Arc.Api.ArcApi
$api.Login("username", "password")

# Get records from specific date range
$startDate = Get-Date "2024-01-01"
$endDate = Get-Date "2024-12-31"
$records = $api.GetRecords($startDate, $endDate)

Write-Host "Found $($records.GetCount()) records`n"

foreach ($record in $records) {
    Write-Host "========================================="
    Write-Host "Record Key: " $record.RecordKey
    Write-Host "Date Recorded: " $record.DateRecorded
    Write-Host "Duration (total): " $record.Duration.TotalHours " hours"
    Write-Host "Study Status: " $record.StudyStatus
    Write-Host "Study Type: " $record.StudyType.Name
    Write-Host "Data Acquisition Machine: " $record.DataAcquisitionMachine
    Write-Host "Is Exported: " $record.IsExported
    Write-Host "Is Local Active Recording: " $record.IsLocalActiveRecording
    
    # Patient Information (for exploration - anonymize before export)
    $patient = $record.Patient
    Write-Host "`n--- PATIENT INFO ---"
    Write-Host "Patient Key: " $patient.PatientKey
    
    $patientFields = $patient.Fields
    $lastNameKey = $patient.DefaultFieldDefinitionKeys.LastName
    $firstNameKey = $patient.DefaultFieldDefinitionKeys.FirstName
    $dobKey = $patient.DefaultFieldDefinitionKeys.Birthdate
    $ageKey = $patient.DefaultFieldDefinitionKeys.Age
    
    Write-Host "Last Name: " $patientFields.GetField($lastNameKey).Value.DisplayText
    Write-Host "First Name: " $patientFields.GetField($firstNameKey).Value.DisplayText
    Write-Host "Age: " $patientFields.GetField($ageKey).Value.DisplayText
    
    # Record Fields
    Write-Host "`n--- RECORD DETAILS ---"
    $recordFields = $record.Fields
    $facilityKey = $record.DefaultFieldDefinitionKeys.Facility
    $physicianKey = $record.DefaultFieldDefinitionKeys.Physician
    $medicationsKey = $record.DefaultFieldDefinitionKeys.Medications
    
    Write-Host "Facility: " $recordFields.GetField($facilityKey).Value.DisplayText
    Write-Host "Physician: " $recordFields.GetField($physicianKey).Value.DisplayText
    Write-Host "Medications: " $recordFields.GetField($medicationsKey).Value.DisplayText
    
    # Event Types Available
    Write-Host "`n--- AVAILABLE EVENT TYPES ---"
    $eventTypes = $record.EventTypes
    foreach ($eventType in $eventTypes) {
        Write-Host "  - $eventType"
    }
    
    Write-Host "`n"
}

$api.Dispose()
```

## 3. Single Record exploration
```powershell

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

# Time Segments (continuous data blocks)
Write-Host "`n=== TIME SEGMENTS (Data Continuity) ==="
$segments = $data.GetTimeSegments()
Write-Host "Number of continuous segments: " $segments.GetCount()
for ($i = 0; $i -lt $segments.GetCount(); $i++) {
    $segment = $segments.GetAt($i)
    Write-Host "Segment $($i+1):"
    Write-Host "  Start Offset: $($segment.StartOffset.TotalSeconds) sec"
    Write-Host "  End Offset: $($segment.EndOffset.TotalSeconds) sec"
    Write-Host "  Duration: $($segment.TotalDuration.TotalMinutes) min"
}

# Common Reference Information
Write-Host "`n=== COMMON REFERENCE DATA ==="
$commonRefs = $data.GetCommonReferenceData()
Write-Host "Reference changes: " $commonRefs.GetCount()
for ($i = 0; $i -lt $commonRefs.GetCount(); $i++) {
    $ref = $commonRefs.GetAt($i)
    Write-Host "At offset $($ref.Offset.TotalSeconds)s: $($ref.ChannelName)"
}

# Montage Information
Write-Host "`n=== MONTAGE INFORMATION ==="
$montages = $data.GetMontages($true)  # true = unique only
Write-Host "Unique Montages: " $montages.GetCount()
for ($i = 0; $i -lt $montages.GetCount(); $i++) {
    $montage = $montages.GetAt($i)
    Write-Host "`nMontage: $($montage.Name)"
    Write-Host "  Used at offset: $($montage.AsViewed.TotalSeconds) sec"
    Write-Host "  Trace Containers: " $montage.TraceContainers.GetCount()
    
    # Show first few traces
    $containers = $montage.TraceContainers
    for ($j = 0; $j -lt [Math]::Min(3, $containers.GetCount()); $j++) {
        $container = $containers.GetAt($j)
        $traces = $container.Traces
        if ($traces.GetCount() -gt 0) {
            $trace = $traces.GetAt(0)
            Write-Host "    Trace: $($trace.Name) | Active: $($trace.Active) | Ref: $($trace.Reference)"
            Write-Host "      Sensitivity: $($trace.Sensitivity) uV/mm | Filters: Lo=$($trace.Locut)Hz Hi=$($trace.Hicut)Hz"
        }
    }
}

$record.Close()
$api.Dispose()
```

## 4. Event Analysis and Export
```powershell

$api = New-Object -comobject Arc.Api.ArcApi
$api.Login("username", "password")

$records = $api.GetRecords()
$record = $records.GetAt(0)
$record.Open()
$data = $record.Data

Write-Host "=== EVENT ANALYSIS ==="

# Get all events
$allEvents = $data.GetEvents()
Write-Host "Total Events: " $allEvents.GetCount()

# Categorize events by type
$eventTypeCount = @{}
for ($i = 0; $i -lt $allEvents.GetCount(); $i++) {
    $event = $allEvents.GetAt($i)
    $type = $event.EventType
    if ($eventTypeCount.ContainsKey($type)) {
        $eventTypeCount[$type]++
    } else {
        $eventTypeCount[$type] = 1
    }
}

Write-Host "`nEvent Summary by Type:"
foreach ($type in $eventTypeCount.Keys | Sort-Object) {
    Write-Host "  $type : $($eventTypeCount[$type])"
}

# Get events in specific time window (e.g., first hour)
Write-Host "`n=== EVENTS IN FIRST HOUR ==="
$firstHourEvents = $data.GetEvents(0, 3600, $null)  # 0-3600 seconds
Write-Host "Events in first hour: " $firstHourEvents.GetCount()

for ($i = 0; $i -lt $firstHourEvents.GetCount(); $i++) {
    $event = $firstHourEvents.GetAt($i)
    Write-Host "`nEvent $($i+1):"
    Write-Host "  Type: $($event.EventType)"
    Write-Host "  Offset: $($event.Offset.TotalSeconds) sec"
    Write-Host "  Duration: $($event.Duration.TotalSeconds) sec"
    Write-Host "  Text: $($event.Text)"
    Write-Host "  Priority: $($event.Priority)"
    Write-Host "  Created by API: $($event.IsCreatedByApi)"
    Write-Host "  Can Delete: $($event.CanDelete)"
}

# Filter specific event types (e.g., seizures and spikes)
Write-Host "`n=== SEIZURE AND SPIKE EVENTS ==="
$clinicalEventTypes = @("Persyst SeizureDetected", "Persyst Spike", "Persyst SpikeBurst")
$clinicalEvents = $data.GetEvents(0, [int]::MaxValue, $clinicalEventTypes)
Write-Host "Clinical events found: " $clinicalEvents.GetCount()

# Export events to CSV
$eventsCsv = "EventType,OffsetSeconds,DurationSeconds,Text,Priority`n"
for ($i = 0; $i -lt $allEvents.GetCount(); $i++) {
    $event = $allEvents.GetAt($i)
    $eventsCsv += "$($event.EventType),$($event.Offset.TotalSeconds),$($event.Duration.TotalSeconds),`"$($event.Text)`",$($event.Priority)`n"
}
$eventsCsv | Out-File "C:\Output\events.csv"

# Get Cortical Stim Events (if applicable)
Write-Host "`n=== CORTICAL STIMULATION EVENTS ==="
$stimEvents = $data.GetCorticalStimEvents()
Write-Host "Cortical Stim Events: " $stimEvents.GetCount()
for ($i = 0; $i -lt $stimEvents.GetCount(); $i++) {
    $stim = $stimEvents.GetAt($i)
    Write-Host "`nStim Event $($i+1):"
    Write-Host "  Offset: $($stim.Offset.TotalSeconds) sec"
    Write-Host "  Current Delivered: $($stim.CurrentDelivered) mA"
    Write-Host "  Response Type: $($stim.ResponseType)"
    Write-Host "  Body Region: $($stim.BodyRegion)"
    Write-Host "  Body Side: $($stim.BodySide)"
    Write-Host "  Patient Task: $($stim.PatientTask)"
    Write-Host "  Functional Response: $($stim.FunctionalResponse)"
}

$record.Close()
$api.Dispose()
```

## 5. Impedance Data Analysis
```powershell

$api = New-Object -comobject Arc.Api.ArcApi
$api.Login("username", "password")

$records = $api.GetRecords()
$record = $records.GetAt(0)
$record.Open()
$data = $record.Data

Write-Host "=== IMPEDANCE ANALYSIS ==="

# Get all impedance measurements
$impedances = $data.GetImpedances(0, [int]::MaxValue)
Write-Host "Total Impedance Measurements: " $impedances.GetCount()

for ($i = 0; $i -lt $impedances.GetCount(); $i++) {
    $impedance = $impedances.GetAt($i)
    Write-Host "`nImpedance Measurement $($i+1):"
    Write-Host "  Time Offset: $($impedance.Offset.TotalSeconds) sec"
    Write-Host "  Low Boundary: $($impedance.LowBoundary) Ohms"
    Write-Host "  High Boundary: $($impedance.HighBoundary) Ohms"
    
    $items = $impedance.Items
    Write-Host "  Channel Impedances:"
    
    # Track quality statistics
    $connected = 0
    $goodImpedance = 0
    $totalChannels = $items.GetCount()
    
    for ($j = 0; $j -lt $items.GetCount(); $j++) {
        $item = $items.GetAt($j)
        $status = if ($item.IsConnected) { "Connected" } else { "Disconnected" }
        $quality = if ($item.Ohms -lt $impedance.HighBoundary) { "Good" } else { "High" }
        
        Write-Host "    $($item.Channel): $($item.Ohms) Ohms ($($item.KiloOhms) kOhms) - $status - $quality"
        
        if ($item.IsConnected) { $connected++ }
        if ($item.Ohms -lt $impedance.HighBoundary -and $item.IsConnected) { $goodImpedance++ }
    }
    
    Write-Host "`n  Summary:"
    Write-Host "    Connected: $connected / $totalChannels"
    Write-Host "    Good Impedance: $goodImpedance / $totalChannels"
    Write-Host "    Quality Rate: $([Math]::Round(($goodImpedance/$totalChannels)*100, 2))%"
}

$record.Close()
$api.Dispose()
```
## 6. Video Data Export
```powershell

$api = New-Object -comobject Arc.Api.ArcApi
$api.Login("username", "password")

# Check if video feature is licensed
$hasVideo = $api.IsEegFeatureLicensed([Arc.Api.LicenseFeatureEEG]::Video)
if (-not $hasVideo) {
    Write-Host "Video feature is not licensed"
    $api.Dispose()
    exit
}

$records = $api.GetRecords()
$record = $records.GetAt(0)
$record.Open()
$data = $record.Data

Write-Host "=== VIDEO DATA EXPORT ==="

$mediaService = $data.MediaService
$mediaChannels = $mediaService.GetMediaChannels()

Write-Host "Available Media Channels: " $mediaChannels.GetCount()

for ($i = 0; $i -lt $mediaChannels.GetCount(); $i++) {
    $channel = $mediaChannels.GetAt($i)
    Write-Host "`nChannel $($i+1):"
    Write-Host "  Video Track Number: $($channel.VideoTrackNumber)"
    Write-Host "  Audio Track Number: $($channel.AudioTrackNumber)"
    
    if ($channel.VideoTrackNumber -ge 0) {
        # Export video to file (full recording)
        $outputPath = "C:\Output\video_channel_$($i+1).mp4"
        
        Write-Host "  Exporting video to: $outputPath"
        
        # Options: Full, Shrink, Split
        $exportResult = $mediaService.ExportMp4Video(
            $channel, 
            0,                    # start offset (seconds)
            [int]::MaxValue,      # duration (use MaxValue for full video)
            [Arc.Api.Media.VideoExportOption]::Full,  # export option
            $outputPath
        )
        
        if ($exportResult.IsSuccess) {
            Write-Host "  Export successful!"
            $files = $exportResult.Files
            Write-Host "  Files created: " $files.GetCount()
            
            $meta = $exportResult.Meta
            for ($j = 0; $j -lt $meta.GetCount(); $j++) {
                $chunk = $meta.GetAt($j)
                Write-Host "    Chunk $($j+1): Offset=$($chunk.Offset.TotalSeconds)s, Duration=$($chunk.Duration.TotalSeconds)s"
            }
        } else {
            Write-Host "  Export failed: $($exportResult.ErrorMessage)"
        }
        
        # Alternative: Get video frames programmatically
        Write-Host "`n  Getting video frames for first 10 seconds..."
        $frames = $mediaService.GetVideoFrames($channel, 0, 10)
        Write-Host "  Retrieved $($frames.GetCount()) frames"
        
        if ($frames.GetCount() -gt 0) {
            $frame = $frames.GetAt(0)
            Write-Host "  First frame:"
            Write-Host "    Time: $($frame.Time)"
            Write-Host "    Width: $($frame.Width) pixels"
            Write-Host "    Height: $($frame.Height) pixels"
            Write-Host "    Bits Per Pixel: $($frame.BitsPerPixel)"
            Write-Host "    Encoding: $($frame.EncodingType)"
            Write-Host "    Buffer Size: $($frame.VideoBuffer.Length) bytes"
        }
    }
    
    if ($channel.AudioTrackNumber -ge 0) {
        Write-Host "`n  Getting audio frames for first 10 seconds..."
        $audioFrames = $mediaService.GetAudioFrames($channel, 0, 10)
        Write-Host "  Retrieved $($audioFrames.GetCount()) audio frames"
        
        if ($audioFrames.GetCount() -gt 0) {
            $audioFrame = $audioFrames.GetAt(0)
            Write-Host "  First audio frame:"
            Write-Host "    Time: $($audioFrame.Time)"
            Write-Host "    Sample Rate: $($audioFrame.SampleRate) Hz"
            Write-Host "    Encoding: $($audioFrame.EncodingType)"
            Write-Host "    Buffer Size: $($audioFrame.AudioBuffer.Length) bytes"
        }
    }
}

$record.Close()
$api.Dispose()
```

## 7. Waveform Data Extraction
```powershell

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
```

## 8. Documents Management
```powershell

$api = New-Object -comobject Arc.Api.ArcApi
$api.Login("username", "password")

$records = $api.GetRecords()
$record = $records.GetAt(0)
$record.Open()
$data = $record.Data

Write-Host "=== ASSOCIATED DOCUMENTS ==="

# Get associated documents
$documents = $data.GetAssociatedDocuments()
Write-Host "Total Associated Documents: " $documents.GetCount()

for ($i = 0; $i -lt $documents.GetCount(); $i++) {
    $doc = $documents.GetAt($i)
    
    Write-Host "`nDocument $($i+1):"
    Write-Host "  Key: $($doc.Key)"
    Write-Host "  Filename: $($doc.FileName)"
    
    # Export document
    $exportPath = "C:\Output\documents\$($doc.FileName)"
    $result = $doc.Export($exportPath)
    
    if ($result.IsSuccess) {
        Write-Host "  Exported to: $exportPath"
    } else {
        Write-Host "  Export failed: $($result.ErrorMessage)"
    }
}

# Add new associated document (example)
Write-Host "`n=== ADDING NEW DOCUMENT ==="
$newDocPath = "C:\Input\clinical_notes.pdf"
if (Test-Path $newDocPath) {
    $result = $data.AddAssociatedDocument($newDocPath)
    if ($result.IsSuccess) {
        Write-Host "Document added successfully"
    } else {
        Write-Host "Failed to add document: $($result.ErrorMessage)"
    }
}

$record.Close()
$api.Dispose()
```

## 9. Anonymization and Export Pipeline
```powershell

$api = New-Object -comobject Arc.Api.ArcApi
$api.Login("username", "password")

# Configuration
$outputBaseDir = "C:\AnonymizedEEG"
$startDate = Get-Date "2024-01-01"
$endDate = Get-Date "2024-12-31"

# Create output directories
$dirs = @(
    "$outputBaseDir\EDF",
    "$outputBaseDir\Events",
    "$outputBaseDir\Impedances",
    "$outputBaseDir\Videos",
    "$outputBaseDir\Metadata"
)
foreach ($dir in $dirs) {
    if (-not (Test-Path $dir)) {
        New-Item -ItemType Directory -Path $dir -Force | Out-Null
    }
}

Write-Host "=== COMPREHENSIVE ANONYMIZATION AND EXPORT ==="

$records = $api.GetRecords($startDate, $endDate)
Write-Host "Processing $($records.GetCount()) records...`n"

$successCount = 0
$errorCount = 0
$exportLog = @()

for ($recordIndex = 0; $recordIndex -lt $records.GetCount(); $recordIndex++) {
    $record = $records.GetAt($recordIndex)
    $anonId = "ANON_" + ($recordIndex + 1).ToString("D5")
    
    Write-Host "[$($recordIndex+1)/$($records.GetCount())] Processing Record: $anonId"
    
    try {
        # Open record
        $openResult = $record.Open()
        if (-not $openResult.IsSuccess) {
            throw "Failed to open record: $($openResult.ErrorMessage)"
        }
        
        # ===== ANONYMIZE PATIENT DATA =====
        Write-Host "  Anonymizing patient data..."
        $patient = $record.Patient
        $patientFields = $patient.Fields
        
        $fieldsToAnonymize = @{
            $patient.DefaultFieldDefinitionKeys.FirstName = "Patient"
            $patient.DefaultFieldDefinitionKeys.LastName = ($recordIndex + 1).ToString()
            $patient.DefaultFieldDefinitionKeys.PatientId = $anonId
            $patient.DefaultFieldDefinitionKeys.MiddleName = ""
            $patient.DefaultFieldDefinitionKeys.OtherId = ""
        }
        
        foreach ($key in $fieldsToAnonymize.Keys) {
            $field = $patientFields.GetField($key)
            if ($field -ne $null) {
                $field.SetValue($fieldsToAnonymize[$key])
            }
        }
        $patientFields.SaveData()
        
        # Anonymize record fields
        $recordFields = $record.Fields
        $recordFieldsToAnonymize = @{
            $record.DefaultFieldDefinitionKeys.Physician = ""
            $record.DefaultFieldDefinitionKeys.ReferringPhysician = ""
            $record.DefaultFieldDefinitionKeys.Comments = ""
            $record.DefaultFieldDefinitionKeys.History = ""
        }
        
        foreach ($key in $recordFieldsToAnonymize.Keys) {
            $field = $recordFields.GetField($key)
            if ($field -ne $null) {
                $field.SetValue($recordFieldsToAnonymize[$key])
            }
        }
        $recordFields.SaveData()
        
        # ===== EXPORT EDF =====
        Write-Host "  Exporting EDF..."
        $edfPath = "$outputBaseDir\EDF\$anonId.edf"
        $edfResult = $record.Data.ExportToEdf($edfPath)
        if ($edfResult.IsSuccess) {
            Write-Host "    EDF exported successfully"
        } else {
            Write-Host "    EDF export failed: $($edfResult.ErrorMessage)"
        }
        
        # ===== EXPORT EVENTS =====
        Write-Host "  Exporting events..."
        $events = $record.Data.GetEvents()
        $eventsCsv = "EventType,OffsetSeconds,DurationSeconds,Text,Priority`n"
        for ($i = 0; $i -lt $events.GetCount(); $i++) {
            $event = $events.GetAt($i)
            # Remove any potential identifying information from event text
            $cleanText = $event.Text -replace $patient.PatientKey.ToString(), $anonId
            $eventsCsv += "$($event.EventType),$($event.Offset.TotalSeconds),$($event.Duration.TotalSeconds),`"$cleanText`",$($event.Priority)`n"
        }
        $eventsCsv | Out-File "$outputBaseDir\Events\$anonId`_events.csv"
        Write-Host "    Events exported: $($events.GetCount())"
        
        # ===== EXPORT IMPEDANCES =====
        Write-Host "  Exporting impedances..."
        $impedances = $record.Data.GetImpedances(0, [int]::MaxValue)
        $impedanceCsv = "MeasurementNumber,OffsetSeconds,Channel,Ohms,IsConnected`n"
        for ($i = 0; $i -lt $impedances.GetCount(); $i++) {
            $impedance = $impedances.GetAt($i)
            $items = $impedance.Items
            for ($j = 0; $j -lt $items.GetCount(); $j++) {
                $item = $items.GetAt($j)
                $impedanceCsv += "$($i+1),$($impedance.Offset.TotalSeconds),$($item.Channel),$($item.Ohms),$($item.IsConnected)`n"
            }
        }
        $impedanceCsv | Out-File "$outputBaseDir\Impedances\$anonId`_impedances.csv"
        Write-Host "    Impedances exported: $($impedances.GetCount()) measurements"
        
        # ===== EXPORT VIDEO (if available) =====
        if ($api.IsEegFeatureLicensed([Arc.Api.LicenseFeatureEEG]::Video)) {
            Write-Host "  Checking for video..."
            $mediaService = $record.Data.MediaService
            $mediaChannels = $mediaService.GetMediaChannels()
            
            for ($i = 0; $i -lt $mediaChannels.GetCount(); $i++) {
                $channel = $mediaChannels.GetAt($i)
                if ($channel.VideoTrackNumber -ge 0) {
                    Write-Host "    Exporting video channel $($i+1)..."
                    $videoPath = "$outputBaseDir\Videos\$anonId`_channel$($i+1).mp4"
                    $videoResult = $mediaService.ExportMp4Video(
                        $channel, 0, [int]::MaxValue,
                        [Arc.Api.Media.VideoExportOption]::Full,
                        $videoPath
                    )
                    if ($videoResult.IsSuccess) {
                        Write-Host "      Video exported successfully"
                    }
                }
            }
        }
        
        # ===== EXPORT METADATA =====
        Write-Host "  Exporting metadata..."
        $metadata = @{
            AnonymousID = $anonId
            OriginalRecordKey = $record.RecordKey.ToString()
            DateRecorded = $record.DateRecorded.ToString("yyyy-MM-dd HH:mm:ss")
            Duration = $record.Duration.TotalHours
            StudyType = $record.StudyType.Name
            StudyStatus = $record.StudyStatus
            DataAcquisitionMachine = $record.DataAcquisitionMachine
            ChannelCount = $record.Data.ChannelInformation.GetCount()
            EventCount = $events.GetCount()
            RecordingStartTime = $record.Data.RecordingStartTime.ToString("yyyy-MM-dd HH:mm:ss")
            TimeZoneOffset = $record.Data.RecordingTimeZoneOffset.TotalHours
        }
        
        $metadata | ConvertTo-Json | Out-File "$outputBaseDir\Metadata\$anonId`_metadata.json"
        
        # Log successful export
        $exportLog += [PSCustomObject]@{
            AnonymousID = $anonId
            OriginalKey = $record.RecordKey
            Status = "Success"
            Timestamp = Get-Date
        }
        
        $successCount++
        Write-Host "  ✓ Record processed successfully`n"
        
    } catch {
        Write-Host "  ✗ Error: $($_.Exception.Message)`n" -ForegroundColor Red
        $exportLog += [PSCustomObject]@{
            AnonymousID = $anonId
            OriginalKey = $record.RecordKey
            Status = "Failed: $($_.Exception.Message)"
            Timestamp = Get-Date
        }
        $errorCount++
    } finally {
        if ($record.IsOpen) {
            $record.Close()
        }
    }
}

# Save export log
$exportLog | Export-Csv "$outputBaseDir\export_log.csv" -NoTypeInformation

Write-Host "`n========================================="
Write-Host "EXPORT COMPLETE"
Write-Host "========================================="
Write-Host "Total Records: $($records.GetCount())"
Write-Host "Successful: $successCount"
Write-Host "Failed: $errorCount"
Write-Host "Output Directory: $outputBaseDir"

$api.Dispose()
```
## 10. Quality Control and Data Validation
```powershell

$api = New-Object -comobject Arc.Api.ArcApi
$api.Login("username", "password")

Write-Host "=== DATA QUALITY CONTROL ==="

$records = $api.GetRecords()
$qualityReport = @()

foreach ($record in $records) {
    try {
        $record.Open()
        $data = $record.Data
        
        # Quality metrics
        $issues = @()
        
        # Check 1: Data duration
        if ($record.Duration.TotalMinutes -lt 1) {
            $issues += "Very short recording (<1 min)"
        }
        
        # Check 2: Number of channels
        $channelCount = $data.ChannelInformation.GetCount()
        if ($channelCount -lt 8) {
            $issues += "Low channel count ($channelCount)"
        }
        
        # Check 3: Data continuity
        $segments = $data.GetTimeSegments()
        $gapCount = $segments.GetCount() - 1
        if ($gapCount -gt 10) {
            $issues += "Many gaps in data ($gapCount gaps)"
        }
        
        # Check 4: Impedance quality
        $impedances = $data.GetImpedances(0, 60)  # First minute
        if ($impedances.GetCount() -gt 0) {
            $firstImpedance = $impedances.GetAt(0)
            $items = $firstImpedance.Items
            $connectedCount = 0
            for ($i = 0; $i -lt $items.GetCount(); $i++) {
                $item = $items.GetAt($i)
                if ($item.IsConnected) { $connectedCount++ }
            }
            $connectionRate = [Math]::Round(($connectedCount / $items.GetCount()) * 100, 1)
            if ($connectionRate -lt 90) {
                $issues += "Poor impedance connection rate ($connectionRate%)"
            }
        }
        
        # Check 5: Event presence
        $events = $data.GetEvents()
        if ($events.GetCount() -eq 0) {
            $issues += "No events recorded"
        }
        
        $qualityReport += [PSCustomObject]@{
            RecordKey = $record.RecordKey
            DateRecorded = $record.DateRecorded
            Duration = $record.Duration.TotalHours
            Channels = $channelCount
            DataGaps = $gapCount
            EventCount = $events.GetCount()
            Issues = ($issues -join "; ")
            QualityStatus = if ($issues.Count -eq 0) { "Good" } else { "Review" }
        }
        
        $record.Close()
        
    } catch {
        Write-Host "Error processing record: $($_.Exception.Message)"
    }
}

# Export quality report
$qualityReport | Export-Csv "C:\Output\quality_report.csv" -NoTypeInformation

# Display summary
Write-Host "`nQuality Summary:"
$goodRecords = ($qualityReport | Where-Object { $_.QualityStatus -eq "Good" }).Count
Write-Host "Good Quality: $goodRecords / $($qualityReport.Count)"
Write-Host "Need Review: $(($qualityReport.Count - $goodRecords))"

$api.Dispose()
```

---------
# EEG-EMU (Epilepsy Monitoring Machine)

1. Nicolet NicVue bought by Natus (short recordings). File types are:

- .e or .eeg – primary EEG data
- .erd – raw EEG binary (used by Neuroworks / Xltek systems)
- .ent – annotations/notes metadata
- .NPA – metadata wrapper used by NicVue

Can these files be read or anonymized? 
The .erd formats are compressed and proprietary, and there is no publicly documented Python library. Users convert to EDF via the Natus GUI instead.
- .erd can be read by unofficial code in Archived XltekDataReader (Python) [(Link)](https://github.com/nyuolab/XltekDataReader)

I asked Natus: " We are currently using XLTEK Neuroworks (Natus) in our EMU. I am reaching out to inquire if it is possible to export anonymised EEG data using an API."
Their response: We do not have an API, but you can use UI or EDFExport.exe. 

Natus provided a Platform Migration Utility to migrate data from legacy NicVue systems into a NeuroWorks database
[download.xltek.com](https://download.xltek.com/eeg/Software/Neuroworks/DOC-020491%20REV%2005%20-%20Platform%20Migration%20Utility%20User%20Guide.pdf#:~:text=from%20legacy%20source%20systems%20such,Database%20application%2C%20used%20with%20NeuroWorks). 

# Requested
- We are trying to batch-export via the command line using EDFExport.exe:
c:\NeuroWorks>edfexport -f "C:\Users\Nicolete\Desktop\TEST\EDFExport\Subject_1.txt" -o "C:\Users\Nicolete\Desktop\TEST\EDFExport"

- Could you please advise on how to specify a template when calling edfexport from the command line (e.g., a flag like -t or a config file path)? - Are there examples or a reference for acceptable template fields and syntax outside the UI?
- Does the command-line tool support anonymization/de-identification (e.g., removing patient name/ID, date of birth, study date/time offsets)?
- If so, what flags or template settings enable this, and can they be applied in batch mode?
- Is there a command-line user guide or man page for EDFExport.exe that covers templates/anonymization?

--------
# Natus 8.5 (long recordings)

Natus’s software suite includes a batch export tool, EDFExport.exe, for converting proprietary files to EDF/EDF+.
- https://data2bids.greydongilmore.com/run_data2bids/04_neuroworks_export
  (Create a workflow and put it here to explain)

- Users first create an export template (.exp file) within NeuroWorks: this template defines which channels to include and to de-identify patient info.
- The template is saved under the Neuroworks Settings directory and must remain there for the exporter to use it. 
- Once the template is prepared, batch conversion is done via command-line. One writes a text file listing the studies (paths to the .eeg files), then runs EDFExport with the template, for example:

```text
"C:\Neuroworks\EDFExport.exe" -f "studies_list.txt" -o "output_folder\"
```
This will output EDF/EDF+ files for each study. 
The EDFExport utility relies on the Natus software environment (not open-source).
EDFExport command-line utility:

```php
Usage:
  EDFExport -s /path-to-study-folder -t path-to-template -o path-to-output_dir
```
- template (likely JSON or XML) to specify header contents and event inclusion. With a custom template, one can remove or anonymize a patient's name, ID, DOB, etc. Without -t, EDFExport won't know how to omit or include metadata fields.

EDFExport- How to use templates with the command? Does it allow anonymisation? 

```phd
c:\NeuroWorks>edfexport -d \\10.40.15.131\public\Archive -edfplus -o "C:\Users\Nicolete\Desktop\GAVINTEST\EDFExport"

c:\NeuroWorks>edfexport -f "C:\Users\Nicolete\Desktop\GAVINTEST\EDFExport\Subject_1.txt"  -o "C:\Users\Nicolete\Desktop\GAVINTEST\EDFExport"
```
In research contexts, the usual approach is to export EEG recordings to EDF via NicVue/NeuroWorks itself, rather than parse .NPA in code. 
(Notably, the Temple University Hospital EEG Corpus was originally in Natus proprietary format and was converted to EDF using NicVue software
[par.nsf.gov](https://par.nsf.gov/servlets/purl/10199699#:~:text=,proprietary%20NicVue%20software%20tool).

----------

# Catwell Arc 3.1.534, waiting for the new App.



