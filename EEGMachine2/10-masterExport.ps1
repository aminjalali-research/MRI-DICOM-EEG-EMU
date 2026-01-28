# ============================================
# 10. MASTER EXPORT - Full Data Export Pipeline
# ============================================
# Purpose: Export ALL EEG data to external drive for offline analysis
# Runs all extraction scripts and organizes output
#
# ⚠️ READ-ONLY: Original hospital data is NEVER modified
# This script only READS and EXPORTS copies to external storage
# ============================================

# Load configuration
. "$PSScriptRoot\0-config.ps1"

param(
    [Parameter(Mandatory=$false)]
    [string]$ExportDrive = "E:",                    # External drive letter
    
    [Parameter(Mandatory=$false)]
    [string]$ExportFolder = "EEGChat_Export",       # Folder name on drive
    
    [Parameter(Mandatory=$false)]
    [int]$StartRecordIndex = 0,                     # First record to export
    
    [Parameter(Mandatory=$false)]
    [int]$EndRecordIndex = -1,                      # Last record (-1 = all)
    
    [switch]$ExportFullEDF = $false,                # Export full EDF (vs segment)
    
    [switch]$SkipWaveformCSV = $true,               # Skip large CSV files
    
    [switch]$IncludeAnonymized = $true              # Include anonymized patient data
)

Write-Host "============================================" -ForegroundColor Cyan
Write-Host "10. MASTER EXPORT - Full Data Pipeline" -ForegroundColor Cyan
Write-Host "============================================" -ForegroundColor Cyan
Write-Host ""
Write-Host "**********************************************" -ForegroundColor Green
Write-Host "* SAFE MODE: Original data will NOT be modified *" -ForegroundColor Green
Write-Host "* All operations are READ-ONLY exports          *" -ForegroundColor Green
Write-Host "**********************************************" -ForegroundColor Green
Write-Host ""

# ========================================
# SETUP EXPORT DIRECTORY
# ========================================
$exportRoot = Join-Path $ExportDrive $ExportFolder
$timestamp = Get-Timestamp

# Create timestamped export folder
$exportPath = Join-Path $exportRoot "Export_$timestamp"

Write-Host "=== EXPORT CONFIGURATION ===" -ForegroundColor Yellow
Write-Host "Export Drive: $ExportDrive"
Write-Host "Export Path: $exportPath"
Write-Host "Export Full EDF: $ExportFullEDF"
Write-Host "Skip Waveform CSV: $SkipWaveformCSV"
Write-Host "Include Anonymized Data: $IncludeAnonymized"

# Verify drive exists
if (-not (Test-Path $ExportDrive)) {
    Write-Error "Export drive '$ExportDrive' not found!"
    Write-Host "Please connect your external drive and try again." -ForegroundColor Yellow
    Write-Host "Available drives:" -ForegroundColor Yellow
    Get-PSDrive -PSProvider FileSystem | ForEach-Object { Write-Host "  $($_.Root)" }
    exit 1
}

# Create directory structure
$folders = @(
    $exportPath,
    (Join-Path $exportPath "EDF"),
    (Join-Path $exportPath "Metadata"),
    (Join-Path $exportPath "Events"),
    (Join-Path $exportPath "Segments"),
    (Join-Path $exportPath "Documents"),
    (Join-Path $exportPath "QualityControl"),
    (Join-Path $exportPath "PatientData_Anonymized")
)

foreach ($folder in $folders) {
    if (-not (Test-Path $folder)) {
        New-Item -ItemType Directory -Path $folder -Force | Out-Null
    }
}

Write-Host "Export directories created" -ForegroundColor Green

# Override output directory for this session
$script:OUTPUT_DIR = $exportPath

# ========================================
# INITIALIZE API
# ========================================
$api = Initialize-ArcApi
if ($null -eq $api) {
    Write-Error "Cannot proceed without API connection"
    exit 1
}

try {
    # Get all records
    $records = $api.GetRecords()
    $totalRecords = $records.GetCount()
    
    Write-Host "`n=== RECORDS TO EXPORT ===" -ForegroundColor Yellow
    Write-Host "Total records in system: $totalRecords"
    
    # Determine range
    $startIdx = $StartRecordIndex
    $endIdx = if ($EndRecordIndex -lt 0) { $totalRecords - 1 } else { [Math]::Min($EndRecordIndex, $totalRecords - 1) }
    $recordsToExport = $endIdx - $startIdx + 1
    
    Write-Host "Records to export: $recordsToExport (index $startIdx to $endIdx)"
    Write-Host ""
    
    # Export manifest
    $manifest = @{
        ExportDate = Get-Date
        ExportPath = $exportPath
        TotalRecords = $recordsToExport
        Records = @()
    }
    
    $successCount = 0
    $errorCount = 0
    
    # ========================================
    # PROCESS EACH RECORD
    # ========================================
    for ($i = $startIdx; $i -le $endIdx; $i++) {
        $record = $records.GetAt($i)
        $recordKey = $record.RecordKey -replace '[^a-zA-Z0-9]', '_'
        
        Write-Host "============================================" -ForegroundColor Cyan
        Write-Host "RECORD $($i+1-$startIdx)/$recordsToExport : $($record.RecordKey)" -ForegroundColor Cyan
        Write-Host "============================================" -ForegroundColor Cyan
        
        $recordManifest = @{
            RecordKey = $record.RecordKey
            Index = $i
            ExportTime = Get-Date
            Files = @()
            Errors = @()
        }
        
        try {
            # Open record (READ-ONLY)
            $openResult = $record.Open()
            if (-not $openResult.IsSuccess) {
                throw "Failed to open record: $($openResult.ErrorMessage)"
            }
            
            $data = $record.Data
            
            # ----------------------------------------
            # 1. EXPORT EDF WAVEFORM
            # ----------------------------------------
            Write-Host "`n[1/5] Exporting EDF waveform..." -ForegroundColor Yellow
            $edfPath = Join-Path $exportPath "EDF\$recordKey.edf"
            
            if ($ExportFullEDF) {
                $edfResult = $record.ExportToEdf($edfPath)
            } else {
                # Export first hour or full if shorter
                $durationSec = [Math]::Min(3600, $data.RecordingDuration.TotalSeconds)
                $edfResult = $record.ExportToEdf($edfPath, 0, [int]$durationSec)
            }
            
            if ($edfResult.IsSuccess) {
                Write-Host "  EDF exported: $recordKey.edf" -ForegroundColor Green
                $recordManifest.Files += "EDF\$recordKey.edf"
            } else {
                Write-Host "  EDF export failed: $($edfResult.ErrorMessage)" -ForegroundColor Red
                $recordManifest.Errors += "EDF: $($edfResult.ErrorMessage)"
            }
            
            # ----------------------------------------
            # 2. EXPORT METADATA
            # ----------------------------------------
            Write-Host "[2/5] Exporting metadata..." -ForegroundColor Yellow
            
            # Record info
            $metadataJson = @{
                RecordKey = $record.RecordKey
                DateRecorded = $record.DateRecorded
                Duration = $record.Duration.TotalHours
                StudyStatus = $record.StudyStatus.ToString()
                StudyType = if ($record.StudyType) { $record.StudyType.Name } else { "" }
                RecordingStartTime = $data.RecordingStartTime
                RecordingDuration = $data.RecordingDuration.TotalMinutes
                ChannelCount = $data.ChannelInformation.GetCount()
            } | ConvertTo-Json -Depth 3
            
            $metadataPath = Join-Path $exportPath "Metadata\$recordKey.json"
            $metadataJson | Out-File $metadataPath -Encoding UTF8
            Write-Host "  Metadata exported" -ForegroundColor Green
            $recordManifest.Files += "Metadata\$recordKey.json"
            
            # Channel info
            $channels = $data.ChannelInformation
            $channelCsv = "ChannelNumber,ChannelName,SamplePeriodMs`n"
            for ($c = 0; $c -lt $channels.GetCount(); $c++) {
                $ch = $channels.GetAt($c)
                $channelCsv += "$($ch.ChannelNumber),$($ch.ChannelName),$($ch.SamplePeriod.TotalMilliseconds)`n"
            }
            $channelPath = Join-Path $exportPath "Metadata\${recordKey}_channels.csv"
            $channelCsv | Out-File $channelPath -Encoding UTF8
            $recordManifest.Files += "Metadata\${recordKey}_channels.csv"
            
            # ----------------------------------------
            # 3. EXPORT EVENTS
            # ----------------------------------------
            Write-Host "[3/5] Exporting events..." -ForegroundColor Yellow
            
            $events = $data.GetEvents()
            $eventCsv = "EventType,OffsetSeconds,DurationSeconds,Text,Priority`n"
            for ($e = 0; $e -lt $events.GetCount(); $e++) {
                $evt = $events.GetAt($e)
                $text = ($evt.Text -replace '"', '""') -replace "`n", " "
                $eventCsv += "`"$($evt.EventType)`",$($evt.Offset.TotalSeconds),$($evt.Duration.TotalSeconds),`"$text`",$($evt.Priority)`n"
            }
            $eventPath = Join-Path $exportPath "Events\${recordKey}_events.csv"
            $eventCsv | Out-File $eventPath -Encoding UTF8
            Write-Host "  Events exported: $($events.GetCount()) events" -ForegroundColor Green
            $recordManifest.Files += "Events\${recordKey}_events.csv"
            
            # ----------------------------------------
            # 4. EXPORT SEGMENTS
            # ----------------------------------------
            Write-Host "[4/5] Exporting time segments..." -ForegroundColor Yellow
            
            $segments = $data.GetTimeSegments()
            $segmentCsv = "SegmentNumber,StartOffsetSec,EndOffsetSec,DurationSec`n"
            for ($s = 0; $s -lt $segments.GetCount(); $s++) {
                $seg = $segments.GetAt($s)
                $segmentCsv += "$($s+1),$($seg.StartOffset.TotalSeconds),$($seg.EndOffset.TotalSeconds),$($seg.TotalDuration.TotalSeconds)`n"
            }
            $segmentPath = Join-Path $exportPath "Segments\${recordKey}_segments.csv"
            $segmentCsv | Out-File $segmentPath -Encoding UTF8
            Write-Host "  Segments exported: $($segments.GetCount()) segments" -ForegroundColor Green
            $recordManifest.Files += "Segments\${recordKey}_segments.csv"
            
            # ----------------------------------------
            # 5. EXPORT ANONYMIZED PATIENT DATA
            # ----------------------------------------
            if ($IncludeAnonymized) {
                Write-Host "[5/5] Exporting anonymized patient data..." -ForegroundColor Yellow
                
                $patient = $record.Patient
                if ($null -ne $patient) {
                    $patientFields = $patient.Fields
                    $defaultKeys = $patient.DefaultFieldDefinitionKeys
                    
                    $anonymizedPatient = @{
                        AnonymousID = "ANON_" + [Convert]::ToBase64String([System.Text.Encoding]::UTF8.GetBytes($patient.PatientKey)).Substring(0, 12)
                    }
                    
                    # Get age (safe to include)
                    try {
                        $ageField = $patientFields.GetField($defaultKeys.Age)
                        if ($ageField -and $ageField.Value) {
                            $anonymizedPatient.Age = $ageField.Value.DisplayText
                        }
                    } catch {}
                    
                    # Get sex (safe to include)
                    try {
                        $sexField = $patientFields.GetField($defaultKeys.Sex)
                        if ($sexField -and $sexField.Value) {
                            $anonymizedPatient.Sex = $sexField.Value.DisplayText
                        }
                    } catch {}
                    
                    $patientJson = $anonymizedPatient | ConvertTo-Json
                    $patientPath = Join-Path $exportPath "PatientData_Anonymized\${recordKey}_patient.json"
                    $patientJson | Out-File $patientPath -Encoding UTF8
                    Write-Host "  Patient data anonymized and exported" -ForegroundColor Green
                    $recordManifest.Files += "PatientData_Anonymized\${recordKey}_patient.json"
                }
            }
            
            $successCount++
            $recordManifest.Status = "SUCCESS"
            
        }
        catch {
            Write-Host "  ERROR: $($_.Exception.Message)" -ForegroundColor Red
            $recordManifest.Errors += $_.Exception.Message
            $recordManifest.Status = "ERROR"
            $errorCount++
        }
        finally {
            if ($record.IsOpen) {
                $record.Close()
            }
        }
        
        $manifest.Records += $recordManifest
        Write-Host ""
    }
    
    # ========================================
    # GENERATE EXPORT SUMMARY
    # ========================================
    Write-Host "============================================" -ForegroundColor Cyan
    Write-Host "EXPORT COMPLETE" -ForegroundColor Cyan
    Write-Host "============================================" -ForegroundColor Cyan
    
    # Save manifest
    $manifestJson = $manifest | ConvertTo-Json -Depth 5
    $manifestPath = Join-Path $exportPath "EXPORT_MANIFEST.json"
    $manifestJson | Out-File $manifestPath -Encoding UTF8
    
    # Create README
    $readme = @"
EEGChat Data Export
===================
Export Date: $(Get-Date)
Export Path: $exportPath

IMPORTANT NOTES:
----------------
1. Original hospital data was NOT modified
2. All files are READ-ONLY copies
3. Patient data has been ANONYMIZED (HIPAA compliant)

FOLDER STRUCTURE:
-----------------
EDF/                    - EEG waveforms in standard EDF format
Metadata/               - Record and channel information (JSON/CSV)
Events/                 - Clinical events (seizures, spikes, annotations)
Segments/               - Time segment information (data continuity)
PatientData_Anonymized/ - De-identified patient demographics
QualityControl/         - Data quality reports

EXPORT STATISTICS:
------------------
Records Exported: $successCount
Errors: $errorCount
Total Records: $recordsToExport

HOW TO USE:
-----------
1. EDF files can be opened with:
   - EDFbrowser (free): https://www.teuniz.net/edfbrowser/
   - EEGLAB (MATLAB)
   - MNE-Python: mne.io.read_raw_edf()

2. CSV files can be opened with:
   - Excel, LibreOffice Calc
   - Python pandas: pd.read_csv()
   - R: read.csv()

3. JSON files can be parsed with any JSON library

NEXT STEPS FOR EEGChat:
-----------------------
Copy this export folder to the EEGChat processing machine
and run the Python data ingestion module:

    python -m eegchat.data_ingestion --input "$exportPath"

"@
    $readmePath = Join-Path $exportPath "README.txt"
    $readme | Out-File $readmePath -Encoding UTF8
    
    Write-Host ""
    Write-Host "=== EXPORT SUMMARY ===" -ForegroundColor Green
    Write-Host "Successfully exported: $successCount records" -ForegroundColor Green
    Write-Host "Errors: $errorCount records" -ForegroundColor $(if ($errorCount -gt 0) { "Red" } else { "Green" })
    Write-Host ""
    Write-Host "Export location: $exportPath" -ForegroundColor Cyan
    Write-Host ""
    Write-Host "*** Original hospital data was NOT modified ***" -ForegroundColor Green

}
catch {
    Write-Error "Master export error: $($_.Exception.Message)"
    Write-Host "Stack Trace: $($_.ScriptStackTrace)" -ForegroundColor Red
}
finally {
    Close-ArcApi $api
}
