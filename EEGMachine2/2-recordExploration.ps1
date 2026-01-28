# ============================================
# 2. RECORD EXPLORATION - Browse All Records
# ============================================
# Purpose: List and explore all records with patient info
# Exports record metadata for further analysis

# Load configuration
. "$PSScriptRoot\0-config.ps1"

Write-Host "============================================" -ForegroundColor Cyan
Write-Host "2. RECORD EXPLORATION - Browse All Records" -ForegroundColor Cyan
Write-Host "============================================" -ForegroundColor Cyan

# Initialize API
$api = Initialize-ArcApi
if ($null -eq $api) {
    Write-Error "Cannot proceed without API connection"
    exit 1
}

try {
    # Get records from configured date range
    Write-Host "`nRetrieving records from $($script:START_DATE.ToString('yyyy-MM-dd')) to $($script:END_DATE.ToString('yyyy-MM-dd'))..." -ForegroundColor Yellow
    
    $records = $api.GetRecords($script:START_DATE, $script:END_DATE)
    $recordCount = $records.GetCount()
    
    Write-Host "Found $recordCount records`n" -ForegroundColor Green

    if ($recordCount -eq 0) {
        Write-Host "No records found in the specified date range." -ForegroundColor Yellow
        Write-Host "Try adjusting START_DATE and END_DATE in 0-config.ps1" -ForegroundColor Yellow
        exit 0
    }

    # Prepare CSV export
    $csvHeader = "RecordKey,DateRecorded,DurationHours,StudyStatus,StudyType,Facility,Physician,PatientKey,PatientAge,EventTypes"
    $csvLines = @($csvHeader)

    # Process each record
    for ($i = 0; $i -lt $recordCount; $i++) {
        $record = $records.GetAt($i)
        
        Write-Host "=========================================" -ForegroundColor Cyan
        Write-Host "Record $($i+1) of $recordCount" -ForegroundColor Cyan
        Write-Host "=========================================" -ForegroundColor Cyan
        
        # Basic record info
        Write-Host "Record Key: $($record.RecordKey)"
        Write-Host "Date Recorded: $($record.DateRecorded)"
        
        $durationHours = 0
        if ($null -ne $record.Duration) {
            $durationHours = [Math]::Round($record.Duration.TotalHours, 2)
            Write-Host "Duration: $durationHours hours"
        }
        
        Write-Host "Study Status: $($record.StudyStatus)"
        
        $studyTypeName = ""
        if ($null -ne $record.StudyType) {
            $studyTypeName = $record.StudyType.Name
            Write-Host "Study Type: $studyTypeName"
        }
        
        Write-Host "Data Acquisition Machine: $($record.DataAcquisitionMachine)"
        Write-Host "Is Exported: $($record.IsExported)"
        Write-Host "Is Local Active Recording: $($record.IsLocalActiveRecording)"
        
        # Patient Information
        $patient = $record.Patient
        $patientKey = ""
        $patientAge = ""
        
        if ($null -ne $patient) {
            Write-Host "`n--- PATIENT INFO ---" -ForegroundColor Yellow
            $patientKey = $patient.PatientKey
            Write-Host "Patient Key: $patientKey"
            
            try {
                $patientFields = $patient.Fields
                $defaultKeys = $patient.DefaultFieldDefinitionKeys
                
                # Get patient fields safely
                if ($null -ne $defaultKeys.LastName) {
                    $lastName = $patientFields.GetField($defaultKeys.LastName)
                    if ($null -ne $lastName -and $null -ne $lastName.Value) {
                        Write-Host "Last Name: $($lastName.Value.DisplayText)"
                    }
                }
                
                if ($null -ne $defaultKeys.FirstName) {
                    $firstName = $patientFields.GetField($defaultKeys.FirstName)
                    if ($null -ne $firstName -and $null -ne $firstName.Value) {
                        Write-Host "First Name: $($firstName.Value.DisplayText)"
                    }
                }
                
                if ($null -ne $defaultKeys.Age) {
                    $age = $patientFields.GetField($defaultKeys.Age)
                    if ($null -ne $age -and $null -ne $age.Value) {
                        $patientAge = $age.Value.DisplayText
                        Write-Host "Age: $patientAge"
                    }
                }
            }
            catch {
                Write-Host "  Could not retrieve patient fields: $($_.Exception.Message)" -ForegroundColor Yellow
            }
        }
        
        # Record Fields
        Write-Host "`n--- RECORD DETAILS ---" -ForegroundColor Yellow
        $facility = ""
        $physician = ""
        
        try {
            $recordFields = $record.Fields
            $recordDefaultKeys = $record.DefaultFieldDefinitionKeys
            
            if ($null -ne $recordDefaultKeys.Facility) {
                $facilityField = $recordFields.GetField($recordDefaultKeys.Facility)
                if ($null -ne $facilityField -and $null -ne $facilityField.Value) {
                    $facility = $facilityField.Value.DisplayText
                    Write-Host "Facility: $facility"
                }
            }
            
            if ($null -ne $recordDefaultKeys.Physician) {
                $physicianField = $recordFields.GetField($recordDefaultKeys.Physician)
                if ($null -ne $physicianField -and $null -ne $physicianField.Value) {
                    $physician = $physicianField.Value.DisplayText
                    Write-Host "Physician: $physician"
                }
            }
            
            if ($null -ne $recordDefaultKeys.Medications) {
                $medicationsField = $recordFields.GetField($recordDefaultKeys.Medications)
                if ($null -ne $medicationsField -and $null -ne $medicationsField.Value) {
                    Write-Host "Medications: $($medicationsField.Value.DisplayText)"
                }
            }
        }
        catch {
            Write-Host "  Could not retrieve record fields: $($_.Exception.Message)" -ForegroundColor Yellow
        }
        
        # Event Types Available
        Write-Host "`n--- AVAILABLE EVENT TYPES ---" -ForegroundColor Yellow
        $eventTypesList = @()
        $eventTypes = $record.EventTypes
        
        if ($null -ne $eventTypes) {
            foreach ($eventType in $eventTypes) {
                Write-Host "  - $eventType"
                $eventTypesList += $eventType
            }
        }
        
        # Add to CSV
        $eventTypesStr = ($eventTypesList -join "; ") -replace ",", ";"
        $csvLines += "$($record.RecordKey),$($record.DateRecorded),$durationHours,$($record.StudyStatus),`"$studyTypeName`",`"$facility`",`"$physician`",$patientKey,$patientAge,`"$eventTypesStr`""
        
        Write-Host ""
    }

    # Export records to CSV
    $timestamp = Get-Timestamp
    $csvContent = $csvLines -join "`n"
    Export-SafeCsv -Content $csvContent -FileName "records_export_$timestamp.csv"
    
    Write-Host "`n=== RECORD EXPLORATION COMPLETE ===" -ForegroundColor Green
    Write-Host "Processed $recordCount records" -ForegroundColor Green

}
catch {
    Write-Error "Error during record exploration: $($_.Exception.Message)"
    Write-Host "Stack Trace: $($_.ScriptStackTrace)" -ForegroundColor Red
}
finally {
    Close-ArcApi $api
}
