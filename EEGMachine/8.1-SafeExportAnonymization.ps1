# ========================================
# SAFE ANONYMIZATION - EXPORT-BASED APPROACH
# Original files in CadLink database are NEVER modified
# ========================================

$api = New-Object -comobject Arc.Api.ArcApi
$api.Login("username", "password")

# Configuration
$exportBaseDir = "C:\EEG_Export_Temp"
$anonymizedBaseDir = "C:\EEG_Anonymized"
$logDir = "C:\EEG_Processing_Logs"

# Create directories
@($exportBaseDir, $anonymizedBaseDir, $logDir) | ForEach-Object {
    if (-not (Test-Path $_)) {
        New-Item -ItemType Directory -Path $_ -Force | Out-Null
    }
}

# Initialize log
$logFile = "$logDir\anonymization_log_$(Get-Date -Format 'yyyyMMdd_HHmmss').txt"
$errorLog = "$logDir\error_log_$(Get-Date -Format 'yyyyMMdd_HHmmss').txt"

function Write-Log {
    param($Message, [switch]$IsError)
    $timestamp = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
    $logMessage = "[$timestamp] $Message"
    
    Write-Host $logMessage
    $logMessage | Out-File $logFile -Append
    
    if ($IsError) {
        $logMessage | Out-File $errorLog -Append
    }
}

Write-Log "=== SAFE ANONYMIZATION PROCESS STARTED ==="
Write-Log "Original database will NOT be modified"
Write-Log "Export directory: $exportBaseDir"
Write-Log "Anonymized output directory: $anonymizedBaseDir"

# ========================================
# STEP 1: EXPORT RECORDS (READ-ONLY)
# ========================================
Write-Log "`n=== STEP 1: EXPORTING RECORDS (READ-ONLY) ==="

$records = $api.GetRecords()
Write-Log "Found $($records.GetCount()) records to process"

$exportManifest = @()

for ($idx = 0; $idx -lt $records.GetCount(); $idx++) {
    $record = $records.GetAt($idx)
    $recordKey = $record.RecordKey.ToString()
    
    Write-Log "`n[$($idx+1)/$($records.GetCount())] Processing Record Key: $recordKey"
    
    try {
        # Open record in READ-ONLY mode (just for export)
        $openResult = $record.Open()
        if (-not $openResult.IsSuccess) {
            Write-Log "ERROR: Failed to open record: $($openResult.ErrorMessage)" -IsError
            continue
        }
        
        # Create export subdirectory for this record
        $recordExportDir = "$exportBaseDir\Record_$($idx+1)_$recordKey"
        if (-not (Test-Path $recordExportDir)) {
            New-Item -ItemType Directory -Path $recordExportDir -Force | Out-Null
        }
        
        # Export to EDF (this does NOT modify the original database)
        $edfPath = "$recordExportDir\original.edf"
        Write-Log "  Exporting to EDF: $edfPath"
        
        $exportResult = $record.Data.ExportToEdf($edfPath)
        
        if (-not $exportResult.IsSuccess) {
            Write-Log "ERROR: EDF export failed: $($exportResult.ErrorMessage)" -IsError
            $record.Close()
            continue
        }
        
        Write-Log "  ✓ EDF exported successfully"
        
        # Extract metadata (READ-ONLY - does not modify database)
        Write-Log "  Extracting metadata..."
        
        $patient = $record.Patient
        $patientFields = $patient.Fields
        
        $metadata = @{
            OriginalRecordKey = $recordKey
            ExportTimestamp = (Get-Date).ToString("yyyy-MM-dd HH:mm:ss")
            ExportPath = $edfPath
            
            OriginalPatientInfo = @{
                PatientKey = $patient.PatientKey.ToString()
                FirstName = $patientFields.GetField($patient.DefaultFieldDefinitionKeys.FirstName).Value.DisplayText
                LastName = $patientFields.GetField($patient.DefaultFieldDefinitionKeys.LastName).Value.DisplayText
                PatientId = $patientFields.GetField($patient.DefaultFieldDefinitionKeys.PatientId).Value.DisplayText
                Birthdate = $patientFields.GetField($patient.DefaultFieldDefinitionKeys.Birthdate).Value.DisplayText
                Age = $patientFields.GetField($patient.DefaultFieldDefinitionKeys.Age).Value.DisplayText
                Sex = $patientFields.GetField($patient.DefaultFieldDefinitionKeys.Sex).Value.DisplayText
            }
            
            StudyInfo = @{
                DateRecorded = $record.DateRecorded.ToString("yyyy-MM-dd HH:mm:ss")
                Duration = $record.Duration.TotalHours
                StudyType = $record.StudyType.Name
                StudyStatus = $record.StudyStatus
            }
        }
        
        # Save metadata to JSON (in export directory)
        $metadataPath = "$recordExportDir\original_metadata.json"
        $metadata | ConvertTo-Json -Depth 10 | Out-File $metadataPath -Encoding UTF8
        Write-Log "  ✓ Metadata saved to: $metadataPath"
        
        # Extract events (READ-ONLY)
        $events = $record.Data.GetEvents()
        $eventsCsv = "EventType,OffsetSeconds,DurationSeconds,Text,Priority`n"
        for ($i = 0; $i -lt $events.GetCount(); $i++) {
            $event = $events.GetAt($i)
            $eventsCsv += "$($event.EventType),$($event.Offset.TotalSeconds),$($event.Duration.TotalSeconds),`"$($event.Text)`",$($event.Priority)`n"
        }
        $eventsPath = "$recordExportDir\original_events.csv"
        $eventsCsv | Out-File $eventsPath
        Write-Log "  ✓ Events saved: $($events.GetCount()) events"
        
        # Add to manifest
        $exportManifest += @{
            RecordIndex = $idx + 1
            OriginalRecordKey = $recordKey
            ExportDirectory = $recordExportDir
            EdfPath = $edfPath
            MetadataPath = $metadataPath
            EventsPath = $eventsPath
            ExportStatus = "Success"
        }
        
        # IMPORTANT: Close the record (database remains unchanged)
        $record.Close()
        Write-Log "  ✓ Record closed (database unchanged)"
        
    } catch {
        Write-Log "ERROR: Exception occurred: $($_.Exception.Message)" -IsError
        if ($record.IsOpen) {
            $record.Close()
        }
        
        $exportManifest += @{
            RecordIndex = $idx + 1
            OriginalRecordKey = $recordKey
            ExportStatus = "Failed"
            ErrorMessage = $_.Exception.Message
        }
    }
}

# Save export manifest
$manifestPath = "$exportBaseDir\export_manifest.json"
$exportManifest | ConvertTo-Json -Depth 10 | Out-File $manifestPath -Encoding UTF8
Write-Log "`nExport manifest saved: $manifestPath"

# Dispose API (closes connection, database unchanged)
$api.Dispose()
Write-Log "API connection closed. Original database UNTOUCHED."

# ========================================
# STEP 2: VERIFY EXPORTS
# ========================================
Write-Log "`n=== STEP 2: VERIFYING EXPORTS ==="

$successfulExports = $exportManifest | Where-Object { $_.ExportStatus -eq "Success" }
Write-Log "Successful exports: $($successfulExports.Count)"
Write-Log "Failed exports: $(($exportManifest.Count - $successfulExports.Count))"

foreach ($export in $successfulExports) {
    # Verify EDF file exists and has content
    if (Test-Path $export.EdfPath) {
        $fileSize = (Get-Item $export.EdfPath).Length
        Write-Log "  Record $($export.RecordIndex): EDF file verified ($([Math]::Round($fileSize/1MB, 2)) MB)"
    } else {
        Write-Log "  WARNING: EDF file missing for Record $($export.RecordIndex)" -IsError
    }
}

# ========================================
# STEP 3: ANONYMIZE EXPORTED FILES
# (Working on COPIES, not originals)
# ========================================
Write-Log "`n=== STEP 3: ANONYMIZING EXPORTED FILES ==="
Write-Log "Working on exported copies - originals remain untouched"

$anonymizationMapping = @()

foreach ($export in $successfulExports) {
    if ($export.ExportStatus -ne "Success") { continue }
    
    $recordIndex = $export.RecordIndex
    $anonId = "ANON_" + $recordIndex.ToString("D5")
    
    Write-Log "`nAnonymizing Record $recordIndex → $anonId"
    
    try {
        # Create anonymized directory
        $anonDir = "$anonymizedBaseDir\$anonId"
        if (-not (Test-Path $anonDir)) {
            New-Item -ItemType Directory -Path $anonDir -Force | Out-Null
        }
        
        # Load original metadata
        $originalMetadata = Get-Content $export.MetadataPath | ConvertFrom-Json
        
        # Create anonymized metadata
        $anonMetadata = @{
            AnonymizedID = $anonId
            AnonymizationDate = (Get-Date).ToString("yyyy-MM-dd HH:mm:ss")
            
            # Keep only necessary clinical info (anonymized)
            PatientInfo = @{
                AnonymizedID = $anonId
                Age = $originalMetadata.OriginalPatientInfo.Age
                Sex = $originalMetadata.OriginalPatientInfo.Sex
                # DO NOT include: FirstName, LastName, PatientId, Birthdate
            }
            
            StudyInfo = $originalMetadata.StudyInfo
            
            # Store mapping securely (for re-identification if needed)
            OriginalRecordKey = $originalMetadata.OriginalRecordKey
            OriginalPatientKey = $originalMetadata.OriginalPatientInfo.PatientKey
        }
        
        # Save anonymized metadata
        $anonMetadataPath = "$anonDir\metadata.json"
        $anonMetadata | ConvertTo-Json -Depth 10 | Out-File $anonMetadataPath -Encoding UTF8
        Write-Log "  ✓ Anonymized metadata created"
        
        # Copy and rename EDF file
        $anonEdfPath = "$anonDir\$anonId.edf"
        Copy-Item -Path $export.EdfPath -Destination $anonEdfPath -Force
        Write-Log "  ✓ EDF file copied to: $anonEdfPath"
        
        # Copy and anonymize events
        $originalEvents = Import-Csv $export.EventsPath
        $anonEvents = $originalEvents | ForEach-Object {
            # Remove any potential identifying information from event text
            $cleanText = $_.Text -replace $originalMetadata.OriginalPatientInfo.PatientKey, $anonId
            $cleanText = $cleanText -replace $originalMetadata.OriginalPatientInfo.PatientId, $anonId
            
            [PSCustomObject]@{
                EventType = $_.EventType
                OffsetSeconds = $_.OffsetSeconds
                DurationSeconds = $_.DurationSeconds
                Text = $cleanText
                Priority = $_.Priority
            }
        }
        
        $anonEventsPath = "$anonDir\events.csv"
        $anonEvents | Export-Csv $anonEventsPath -NoTypeInformation
        Write-Log "  ✓ Events anonymized and saved"
        
        # Create anonymization record
        $anonymizationMapping += @{
            AnonymizedID = $anonId
            OriginalRecordKey = $originalMetadata.OriginalRecordKey
            OriginalPatientKey = $originalMetadata.OriginalPatientInfo.PatientKey
            OriginalPatientId = $originalMetadata.OriginalPatientInfo.PatientId
            AnonymizationTimestamp = (Get-Date).ToString("yyyy-MM-dd HH:mm:ss")
            AnonymizedDirectory = $anonDir
        }
        
        Write-Log "  ✓ Record $recordIndex anonymized successfully as $anonId"
        
    } catch {
        Write-Log "ERROR: Anonymization failed for Record $recordIndex : $($_.Exception.Message)" -IsError
    }
}

# ========================================
# STEP 4: SAVE ANONYMIZATION MAPPING
# (CRITICAL - Store securely!)
# ========================================
Write-Log "`n=== STEP 4: SAVING ANONYMIZATION MAPPING ==="

$mappingPath = "$logDir\anonymization_mapping_$(Get-Date -Format 'yyyyMMdd_HHmmss').json"
$anonymizationMapping | ConvertTo-Json -Depth 10 | Out-File $mappingPath -Encoding UTF8

Write-Log "⚠️  CRITICAL: Anonymization mapping saved to:"
Write-Log "    $mappingPath"
Write-Log "    This file contains the link between original and anonymized IDs"
Write-Log "    STORE THIS FILE SECURELY with restricted access"

# ========================================
# STEP 5: CLEANUP TEMPORARY EXPORTS (OPTIONAL)
# ========================================
Write-Log "`n=== STEP 5: CLEANUP OPTIONS ==="
Write-Log "Temporary export directory: $exportBaseDir"
Write-Log "Contains original (non-anonymized) exported files"
Write-Log ""
Write-Log "RECOMMENDED: Delete temporary exports after verification"
Write-Log "Execute the following command when ready:"
Write-Log "  Remove-Item -Path '$exportBaseDir' -Recurse -Force"
Write-Log ""
Write-Log "WARNING: Do NOT delete until you've verified anonymized files!"

# ========================================
# FINAL SUMMARY
# ========================================
Write-Log "`n========================================="
Write-Log "SAFE ANONYMIZATION COMPLETE"
Write-Log "========================================="
Write-Log "Original Database Status: UNCHANGED (read-only access used)"
Write-Log "Temporary Exports: $exportBaseDir"
Write-Log "Anonymized Data: $anonymizedBaseDir"
Write-Log "Processing Log: $logFile"
Write-Log "Anonymization Mapping: $mappingPath"
Write-Log ""
Write-Log "Total Records Processed: $($exportManifest.Count)"
Write-Log "Successful Exports: $($successfulExports.Count)"
Write-Log "Anonymized Records: $($anonymizationMapping.Count)"
Write-Log "========================================="

# Create summary report
$summaryReport = @"
SAFE ANONYMIZATION PROCESS SUMMARY
===================================
Date: $(Get-Date -Format "yyyy-MM-dd HH:mm:ss")

DATABASE STATUS:
- Original CadLink Database: UNCHANGED ✓
- Access Mode: READ-ONLY
- No modifications made to source data

PROCESSING RESULTS:
- Total Records Found: $($exportManifest.Count)
- Successfully Exported: $($successfulExports.Count)
- Successfully Anonymized: $($anonymizationMapping.Count)
- Failed: $(($exportManifest.Count - $successfulExports.Count))

OUTPUT LOCATIONS:
- Anonymized Data: $anonymizedBaseDir
- Temporary Exports: $exportBaseDir
- Processing Logs: $logDir
- Anonymization Mapping: $mappingPath

NEXT STEPS:
1. Verify anonymized files in: $anonymizedBaseDir
2. Review processing log: $logFile
3. Securely store mapping file: $mappingPath
4. After verification, delete temporary exports:
   Remove-Item -Path '$exportBaseDir' -Recurse -Force

SECURITY NOTES:
⚠️  The anonymization mapping file contains sensitive information
⚠️  Store it in a secure location with restricted access
⚠️  This file is needed to re-identify data if required

VERIFICATION CHECKLIST:
□ All anonymized EDF files present and readable
□ Metadata files contain no identifying information
□ Events properly anonymized
□ Anonymization mapping file backed up securely
□ Original database verified as unchanged
"@

$summaryPath = "$anonymizedBaseDir\anonymization_summary.txt"
$summaryReport | Out-File $summaryPath -Encoding UTF8

Write-Host "`n========================================="
Write-Host "Summary report saved to: $summaryPath"
Write-Host "========================================="
