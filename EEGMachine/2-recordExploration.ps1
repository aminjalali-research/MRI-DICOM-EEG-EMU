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