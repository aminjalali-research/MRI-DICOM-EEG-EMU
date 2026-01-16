# ========================================
# VERIFICATION SCRIPT
# Confirms original database was not modified
# ========================================

$api = New-Object -comobject Arc.Api.ArcApi
$api.Login("username", "password")

Write-Host "=== VERIFYING ORIGINAL DATABASE INTEGRITY ==="

# Load the anonymization mapping to get original record keys
$mappingFile = "C:\EEG_Processing_Logs\anonymization_mapping_[timestamp].json"  # Update with actual file
if (Test-Path $mappingFile) {
    $mapping = Get-Content $mappingFile | ConvertFrom-Json
    
    Write-Host "`nChecking $($mapping.Count) original records..."
    
    $allIntact = $true
    
    foreach ($map in $mapping) {
        $recordKey = [Guid]$map.OriginalRecordKey
        
        # Try to get the record by key
        try {
            $record = $api.GetRecord($recordKey)
            $record.Open()
            
            # Check if patient info is still original (not anonymized)
            $patient = $record.Patient
            $patientFields = $patient.Fields
            
            $firstName = $patientFields.GetField($patient.DefaultFieldDefinitionKeys.FirstName).Value.DisplayText
            $patientId = $patientFields.GetField($patient.DefaultFieldDefinitionKeys.PatientId).Value.DisplayText
            
            # Check if it's been anonymized (would start with "Patient" or "ANON_")
            if ($firstName -eq "Patient" -or $patientId.StartsWith("ANON_")) {
                Write-Host "  ✗ WARNING: Record $recordKey appears to have been modified!" -ForegroundColor Red
                $allIntact = $false
            } else {
                Write-Host "  ✓ Record $recordKey - Original data intact" -ForegroundColor Green
            }
            
            $record.Close()
            
        } catch {
            Write-Host "  ✗ ERROR: Could not access record $recordKey : $($_.Exception.Message)" -ForegroundColor Red
            $allIntact = $false
        }
    }
    
    Write-Host "`n========================================="
    if ($allIntact) {
        Write-Host "✓ VERIFICATION PASSED" -ForegroundColor Green
        Write-Host "All original records remain unchanged in database" -ForegroundColor Green
    } else {
        Write-Host "✗ VERIFICATION FAILED" -ForegroundColor Red
        Write-Host "Some records may have been modified" -ForegroundColor Red
    }
    Write-Host "========================================="
    
} else {
    Write-Host "ERROR: Mapping file not found: $mappingFile"
}

$api.Dispose()
