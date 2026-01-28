# ============================================
# 6. WAVEFORM DATA EXTRACTION
# ============================================
# Purpose: Extract raw EEG waveform data for analysis
# Exports data to EDF format AND CSV format
#
# ⚠️ READ-ONLY: Original data is NOT modified
# This script only READS and EXPORTS copies
# ============================================

# Load configuration
. "$PSScriptRoot\0-config.ps1"

param(
    [int]$RecordIndex = 0,
    [int]$StartOffset = $script:WAVEFORM_START_OFFSET,
    [int]$Duration = $script:WAVEFORM_DURATION,
    [switch]$ExportFullRecord = $false,
    [switch]$SkipCSV = $false
)

Write-Host "============================================" -ForegroundColor Cyan
Write-Host "6. WAVEFORM DATA EXTRACTION (READ-ONLY)" -ForegroundColor Cyan
Write-Host "============================================" -ForegroundColor Cyan
Write-Host ""
Write-Host "*** Original data will NOT be modified ***" -ForegroundColor Green
Write-Host "*** Exporting COPIES in EDF and CSV format ***" -ForegroundColor Green
Write-Host ""

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
    Write-Host "Reading waveform from record: $($record.RecordKey)" -ForegroundColor Yellow
    Write-Host "(Original data remains UNCHANGED)" -ForegroundColor Gray

    $timestamp = Get-Timestamp
    $recordKey = $record.RecordKey -replace '[^a-zA-Z0-9]', '_'

    # ========================================
    # EXPORT TO EDF FORMAT (Standard EEG Format)
    # ========================================
    Write-Host "`n=== EXPORTING TO EDF FORMAT ===" -ForegroundColor Yellow
    Write-Host "EDF (European Data Format) is the standard format for EEG data"
    
    $edfFileName = "EEG_${recordKey}_$timestamp.edf"
    $edfFilePath = Join-Path $script:OUTPUT_DIR $edfFileName
    
    if ($ExportFullRecord) {
        Write-Host "Exporting FULL record to EDF..."
        $edfResult = $record.ExportToEdf($edfFilePath)
    } else {
        $endOffset = $StartOffset + $Duration
        Write-Host "Exporting segment ($StartOffset s to $endOffset s) to EDF..."
        $edfResult = $record.ExportToEdf($edfFilePath, $StartOffset, $endOffset)
    }
    
    if ($edfResult.IsSuccess) {
        Write-Host "EDF Export SUCCESS: $edfFileName" -ForegroundColor Green
        Write-Host "  Path: $edfFilePath" -ForegroundColor Cyan
    } else {
        Write-Host "EDF Export FAILED: $($edfResult.ErrorMessage)" -ForegroundColor Red
    }

    # ========================================
    # EXTRACTION PARAMETERS FOR CSV
    # ========================================
    Write-Host "`n=== EXTRACTION PARAMETERS ===" -ForegroundColor Yellow
    Write-Host "Start Offset: $StartOffset seconds"
    Write-Host "Duration: $Duration seconds"
    Write-Host "Max CSV Samples: $($script:MAX_CSV_SAMPLES)"

    if (-not $SkipCSV) {
        # ========================================
        # GET DISCONTINUOUS WAVEFORM DATA FOR CSV
        # ========================================
        Write-Host "`n=== EXTRACTING WAVEFORM DATA FOR CSV ===" -ForegroundColor Yellow
        Write-Host "This may take a moment for large extractions..."

        $channelData = $data.GetDiscontinuousWaveformData($StartOffset, $Duration)
        $channelCount = $channelData.GetCount()
        Write-Host "Channels retrieved: $channelCount" -ForegroundColor Green

        if ($channelCount -eq 0) {
            Write-Host "No waveform data found for the specified time range" -ForegroundColor Yellow
        }
    }

    # ========================================
    # PROCESS EACH CHANNEL
    # ========================================
    Write-Host "`n=== CHANNEL WAVEFORM DETAILS ===" -ForegroundColor Yellow
    
    $channelMetadata = @()
    
    for ($i = 0; $i -lt $channelCount; $i++) {
        $channel = $channelData.GetAt($i)
        
        Write-Host "`nChannel: $($channel.ChannelName)" -ForegroundColor Cyan
        Write-Host "  Total Samples: $($channel.SampleCount)"
        
        if ($null -ne $channel.StartOffset) {
            Write-Host "  Start Offset: $([Math]::Round($channel.StartOffset.TotalSeconds, 3)) sec"
        }
        if ($null -ne $channel.EndOffset) {
            Write-Host "  End Offset: $([Math]::Round($channel.EndOffset.TotalSeconds, 3)) sec"
        }
        if ($null -ne $channel.StartTime) {
            Write-Host "  Start Time (UTC): $($channel.StartTime)"
        }
        
        $waveformSegmentCount = $channel.WaveformData.GetCount()
        Write-Host "  Waveform Segments: $waveformSegmentCount"
        
        # Process waveform segments
        $waveforms = $channel.WaveformData
        for ($j = 0; $j -lt $waveformSegmentCount; $j++) {
            $waveform = $waveforms.GetAt($j)
            
            $sampleRate = if ($waveform.SamplePeriod.TotalMilliseconds -gt 0) {
                [Math]::Round(1000.0 / $waveform.SamplePeriod.TotalMilliseconds, 2)
            } else { 0 }
            
            Write-Host "`n  Segment $($j+1):" -ForegroundColor Gray
            Write-Host "    Samples: $($waveform.SampleCount)"
            Write-Host "    Sample Rate: $sampleRate Hz"
            Write-Host "    Start Offset: $([Math]::Round($waveform.StartOffset.TotalSeconds, 3)) sec"
            Write-Host "    End Offset: $([Math]::Round($waveform.EndOffset.TotalSeconds, 3)) sec"
            Write-Host "    Filters: LoCut=$($waveform.LowCutFilter)Hz HiCut=$($waveform.HighCutFilter)Hz Notch=$($waveform.NotchFilter)Hz"
            
            # Get data statistics
            $dataPoints = $waveform.Data
            if ($dataPoints.Length -gt 0) {
                $stats = $dataPoints | Measure-Object -Minimum -Maximum -Average
                Write-Host "    Data Range: $([Math]::Round($stats.Minimum, 2)) to $([Math]::Round($stats.Maximum, 2))"
                Write-Host "    Mean Value: $([Math]::Round($stats.Average, 2))"
                Write-Host "    First 5 values: $($dataPoints[0..([Math]::Min(4, $dataPoints.Length-1))] -join ', ')"
            }
            
            # Store metadata
            if ($j -eq 0) {
                $channelMetadata += @{
                    Name = $channel.ChannelName
                    SampleCount = $waveform.SampleCount
                    SampleRate = $sampleRate
                    SamplePeriod = $waveform.SamplePeriod.TotalSeconds
                    LowCutFilter = $waveform.LowCutFilter
                    HighCutFilter = $waveform.HighCutFilter
                    NotchFilter = $waveform.NotchFilter
                }
            }
        }
    }

    # ========================================
    # EXPORT WAVEFORM DATA TO CSV
    # ========================================
    Write-Host "`n=== EXPORTING WAVEFORM DATA ===" -ForegroundColor Yellow
    
    $timestamp = Get-Timestamp
    $recordKey = $record.RecordKey -replace '[^a-zA-Z0-9]', '_'
    
    # Export metadata
    $metadataCsv = "ChannelName,SampleCount,SampleRateHz,SamplePeriodSec,LowCutFilterHz,HighCutFilterHz,NotchFilterHz`n"
    foreach ($meta in $channelMetadata) {
        $metadataCsv += "$($meta.Name),$($meta.SampleCount),$($meta.SampleRate),$($meta.SamplePeriod),$($meta.LowCutFilter),$($meta.HighCutFilter),$($meta.NotchFilter)`n"
    }
    Export-SafeCsv -Content $metadataCsv -FileName "waveform_metadata_${recordKey}_$timestamp.csv"

    # Export waveform data (time-series CSV)
    Write-Host "Building waveform CSV (this may take a moment)..."
    
    # Build header
    $header = "Time_Sec"
    for ($i = 0; $i -lt $channelCount; $i++) {
        $channel = $channelData.GetAt($i)
        $header += ",$($channel.ChannelName)"
    }
    
    # Get sample count and sample period from first channel
    $firstChannel = $channelData.GetAt(0)
    $firstWaveform = $firstChannel.WaveformData.GetAt(0)
    $totalSamples = $firstWaveform.SampleCount
    $samplePeriod = $firstWaveform.SamplePeriod.TotalSeconds
    
    # Limit samples for CSV
    $samplesToExport = [Math]::Min($totalSamples, $script:MAX_CSV_SAMPLES)
    Write-Host "Exporting $samplesToExport of $totalSamples samples..."
    
    # Build CSV content
    $csvLines = New-Object System.Collections.ArrayList
    $csvLines.Add($header) | Out-Null
    
    for ($sample = 0; $sample -lt $samplesToExport; $sample++) {
        $timeValue = [Math]::Round($StartOffset + ($sample * $samplePeriod), 6)
        $line = "$timeValue"
        
        for ($ch = 0; $ch -lt $channelCount; $ch++) {
            $channel = $channelData.GetAt($ch)
            $waveform = $channel.WaveformData.GetAt(0)
            $dataPoints = $waveform.Data
            
            if ($sample -lt $dataPoints.Length) {
                $line += ",$($dataPoints[$sample])"
            } else {
                $line += ",NA"
            }
        }
        
        $csvLines.Add($line) | Out-Null
        
        # Progress indicator
        if ($sample % 1000 -eq 0 -and $sample -gt 0) {
            Write-Host "  Processed $sample samples..." -ForegroundColor Gray
        }
    }
    
    $waveformCsv = $csvLines -join "`n"
    Export-SafeCsv -Content $waveformCsv -FileName "waveform_data_${recordKey}_${StartOffset}s_${Duration}s_$timestamp.csv"
    } # End of CSV export block

    # Export summary
    $summary = @"
Waveform Extraction Summary (READ-ONLY EXPORT)
Generated: $(Get-Date)
Record: $($record.RecordKey)
============================================

*** ORIGINAL DATA WAS NOT MODIFIED ***

Extraction Parameters:
- Start Offset: $StartOffset seconds
- Duration: $Duration seconds
- Full Record Export: $ExportFullRecord

Files Generated:
- EDF File: $edfFileName (Standard EEG format)
$(if (-not $SkipCSV) { "- CSV Metadata: waveform_metadata_${recordKey}_$timestamp.csv" })
$(if (-not $SkipCSV) { "- CSV Data: waveform_data_${recordKey}_${StartOffset}s_${Duration}s_$timestamp.csv" })

EDF Export Status: $(if ($edfResult.IsSuccess) { "SUCCESS" } else { "FAILED: $($edfResult.ErrorMessage)" })

Note: EDF files can be opened with:
- EDFbrowser (free, cross-platform)
- EEGLAB (MATLAB)
- MNE-Python
- Any EDF-compatible EEG software

Status: EXPORT COMPLETE (Original unchanged)
"@
    Export-SafeCsv -Content $summary -FileName "waveform_summary_${recordKey}_$timestamp.txt"

    Write-Host "`n=== WAVEFORM EXTRACTION COMPLETE ===" -ForegroundColor Green
    Write-Host "Original data was NOT modified." -ForegroundColor Green
    Write-Host "EDF file: $edfFileName" -ForegroundColor Cyan
    Write-Host "Output directory: $($script:OUTPUT_DIR)" -ForegroundColor Cyan

}
catch {
    Write-Error "Error during waveform extraction: $($_.Exception.Message)"
    Write-Host "Stack Trace: $($_.ScriptStackTrace)" -ForegroundColor Red
}
finally {
    if ($null -ne $record -and $record.IsOpen) {
        $record.Close()
    }
    Close-ArcApi $api
}
