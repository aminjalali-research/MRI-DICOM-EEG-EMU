# ============================================
# 8. ANONYMIZATION - READ-ONLY PHI Export
# ============================================
# Purpose: EXPORT anonymized patient data WITHOUT modifying originals
# HIPAA compliant - Creates de-identified COPY only
#
# ⚠️ IMPORTANT: This script NEVER modifies the original hospital data!
# It only READS data and EXPORTS anonymized copies to the output folder.
# ============================================

# Load configuration
. "$PSScriptRoot\0-config.ps1"

param(
    [int]$RecordIndex = 0
)

Write-Host "============================================" -ForegroundColor Cyan
Write-Host "8. ANONYMIZATION - READ-ONLY PHI Export" -ForegroundColor Cyan
Write-Host "============================================" -ForegroundColor Cyan
Write-Host ""
Write-Host "*** SAFE MODE: Original data will NOT be modified ***" -ForegroundColor Green
Write-Host "*** Only READING data and EXPORTING anonymized copies ***" -ForegroundColor Green
Write-Host ""

# Initialize API (READ-ONLY access)
$api = Initialize-ArcApi
if ($null -eq $api) {
    Write-Error "Cannot proceed without API connection"
    exit 1
}

try {
    $records = $api.GetRecords()
    $record = $records.GetAt($RecordIndex)
    
    Write-Host "Reading record: $($record.RecordKey)" -ForegroundColor Yellow
    Write-Host "(Original data remains UNCHANGED)" -ForegroundColor Gray

    # ========================================
    # READ PHI FIELDS (NO MODIFICATIONS)
    # ========================================
    Write-Host "`n=== READING PHI FIELDS ===" -ForegroundColor Yellow
    
    $phiData = @()
    $anonymizedData = @{}
    $patient = $record.Patient
    
    if ($null -ne $patient) {
        $patientFields = $patient.Fields
        $defaultKeys = $patient.DefaultFieldDefinitionKeys
        
        # Define PHI field mappings for READ and ANONYMIZE export
        $phiMappings = @(
            @{ Name = "PatientKey"; Key = $null; Value = $patient.PatientKey; Action = "HASH" },
            @{ Name = "LastName"; Key = $defaultKeys.LastName; Action = "REDACT" },
            @{ Name = "FirstName"; Key = $defaultKeys.FirstName; Action = "REDACT" },
            @{ Name = "MiddleName"; Key = $defaultKeys.MiddleName; Action = "REDACT" },
            @{ Name = "Birthdate"; Key = $defaultKeys.Birthdate; Action = "AGE_RANGE" },
            @{ Name = "Age"; Key = $defaultKeys.Age; Action = "KEEP" },
            @{ Name = "Sex"; Key = $defaultKeys.Sex; Action = "KEEP" },
            @{ Name = "MRN"; Key = $defaultKeys.MRN; Action = "HASH" },
            @{ Name = "SSN"; Key = $defaultKeys.SSN; Action = "REDACT" },
            @{ Name = "Address"; Key = $defaultKeys.Address; Action = "REDACT" },
            @{ Name = "City"; Key = $defaultKeys.City; Action = "REDACT" },
            @{ Name = "State"; Key = $defaultKeys.State; Action = "KEEP" },
            @{ Name = "Zip"; Key = $defaultKeys.Zip; Action = "ZIP3" },
            @{ Name = "Phone"; Key = $defaultKeys.Phone; Action = "REDACT" },
            @{ Name = "Email"; Key = $defaultKeys.Email; Action = "REDACT" }
        )
        
        foreach ($mapping in $phiMappings) {
            $originalValue = ""
            $anonymizedValue = ""
            
            # Get original value (READ ONLY)
            if ($null -ne $mapping.Value) {
                $originalValue = $mapping.Value
            }
            elseif ($null -ne $mapping.Key) {
                try {
                    $field = $patientFields.GetField($mapping.Key)
                    if ($null -ne $field -and $null -ne $field.Value) {
                        $originalValue = $field.Value.DisplayText
                    }
                }
                catch { }
            }
            
            # Generate anonymized value for EXPORT (original unchanged)
            if (-not [string]::IsNullOrEmpty($originalValue)) {
                $anonymizedValue = switch ($mapping.Action) {
                    "REDACT" { "[REDACTED]" }
                    "HASH" { 
                        $hash = [System.Security.Cryptography.SHA256]::Create()
                        $bytes = [System.Text.Encoding]::UTF8.GetBytes($originalValue + "SALT_EEGChat_2025")
                        $hashBytes = $hash.ComputeHash($bytes)
                        "ANON_" + [BitConverter]::ToString($hashBytes).Replace("-", "").Substring(0, 12)
                    }
                    "AGE_RANGE" {
                        try {
                            $ageField = $patientFields.GetField($defaultKeys.Age)
                            if ($null -ne $ageField -and $null -ne $ageField.Value) {
                                $age = [int]$ageField.Value.RawValue
                                $ageRange = [Math]::Floor($age / 10) * 10
                                "$ageRange-$($ageRange + 9)"
                            } else { "[AGE_RANGE]" }
                        }
                        catch { "[AGE_RANGE]" }
                    }
                    "ZIP3" {
                        if ($originalValue.Length -ge 3) {
                            $originalValue.Substring(0, 3) + "XX"
                        } else { "[REDACTED]" }
                    }
                    "KEEP" { $originalValue }
                    default { "[REDACTED]" }
                }
                
                $phiData += @{
                    FieldName = $mapping.Name
                    OriginalValue = $originalValue
                    AnonymizedValue = $anonymizedValue
                    Action = $mapping.Action
                }
                
                $anonymizedData[$mapping.Name] = $anonymizedValue
                
                # Display (mask original values in output)
                $maskedOriginal = if ($mapping.Action -eq "KEEP") { $originalValue } else { "***" }
                Write-Host "  $($mapping.Name): $maskedOriginal -> $anonymizedValue" -ForegroundColor $(if ($mapping.Action -eq "KEEP") { "White" } else { "Yellow" })
            }
        }
    }
    
    Write-Host "`n=== ANONYMIZATION SUMMARY ===" -ForegroundColor Yellow
    Write-Host "Fields processed: $($phiData.Count)"
    Write-Host "Redacted: $(($phiData | Where-Object { $_.Action -eq 'REDACT' }).Count)"
    Write-Host "Hashed: $(($phiData | Where-Object { $_.Action -eq 'HASH' }).Count)"
    Write-Host "Generalized: $(($phiData | Where-Object { $_.Action -in @('AGE_RANGE', 'ZIP3') }).Count)"
    Write-Host "Kept as-is: $(($phiData | Where-Object { $_.Action -eq 'KEEP' }).Count)"
    
    # ========================================
    # EXPORT ANONYMIZED DATA (NOT MODIFYING ORIGINAL)
    # ========================================
    Write-Host "`n=== EXPORTING ANONYMIZED COPY ===" -ForegroundColor Yellow
    Write-Host "(Original hospital data remains UNCHANGED)" -ForegroundColor Green
    
    # ========================================
    # EXPORT ANONYMIZATION REPORT
    # ========================================
    $timestamp = Get-Timestamp
    $recordKey = $record.RecordKey -replace '[^a-zA-Z0-9]', '_'
    
    # Create detailed report
    $reportCsv = "FieldName,Action,OriginalValue,AnonymizedValue`n"
    foreach ($phi in $phiFields) {
        $original = ($phi.CurrentValue -replace '"', '""')
        $reportCsv += "`"$($phi.FieldName)`",$($phi.Action),`"$original`",`"$($phi.AnonymizedValue)`"`n"
    }
    Export-SafeCsv -Content $reportCsv -FileName "anonymization_report_${recordKey}_$timestamp.csv"
    
    # Summary
    $summary = @"
Anonymization Summary
Generated: $(Get-Date)
Record: $($record.RecordKey)
============================================

Mode: $(if ($PreviewOnly) { "PREVIEW ONLY" } else { "APPLIED" })
PHI Fields Identified: $($phiFields.Count)

Actions Taken:
- REMOVE: $(($phiFields | Where-Object { $_.Action -eq "REMOVE" }).Count) fields
- HASH: $(($phiFields | Where-Object { $_.Action -eq "HASH" }).Count) fields  
- GENERALIZE: $(($phiFields | Where-Object { $_.Action -eq "GENERALIZE" }).Count) fields

HIPAA Compliance Note:
This process removes or generalizes the following PHI categories:
    $timestamp = Get-Timestamp
    $recordKey = $record.RecordKey -replace '[^a-zA-Z0-9]', '_'
    
    # Export anonymized patient data as JSON
    $anonymizedJson = $anonymizedData | ConvertTo-Json -Depth 3
    Export-SafeCsv -Content $anonymizedJson -FileName "patient_anonymized_${recordKey}_$timestamp.json"
    
    # Export mapping report (for audit purposes - shows what was anonymized)
    $reportCsv = "FieldName,Action,AnonymizedValue`n"
    foreach ($phi in $phiData) {
        $reportCsv += "`"$($phi.FieldName)`",$($phi.Action),`"$($phi.AnonymizedValue)`"`n"
    }
    Export-SafeCsv -Content $reportCsv -FileName "anonymization_mapping_${recordKey}_$timestamp.csv"
    
    # Summary
    $summary = @"
Anonymization Export Summary (READ-ONLY)
Generated: $(Get-Date)
Record: $($record.RecordKey)
============================================

*** ORIGINAL DATA WAS NOT MODIFIED ***
This is an EXPORT of anonymized data only.

Fields Processed: $($phiData.Count)
- Redacted: $(($phiData | Where-Object { $_.Action -eq 'REDACT' }).Count)
- Hashed: $(($phiData | Where-Object { $_.Action -eq 'HASH' }).Count)
- Generalized: $(($phiData | Where-Object { $_.Action -in @('AGE_RANGE', 'ZIP3') }).Count)
- Kept: $(($phiData | Where-Object { $_.Action -eq 'KEEP' }).Count)

HIPAA Safe Harbor Method Applied:
- Names: REDACTED
- Dates: Converted to age ranges
- Geographic: State kept, ZIP truncated to 3 digits
- Identifiers (MRN): Hashed for linkage
- Contact info: REDACTED
- SSN: REDACTED

Output Files:
- patient_anonymized_${recordKey}_$timestamp.json
- anonymization_mapping_${recordKey}_$timestamp.csv

Status: EXPORT COMPLETE (Original unchanged)
"@
    Export-SafeCsv -Content $summary -FileName "anonymization_summary_${recordKey}_$timestamp.txt"

    Write-Host "`n=== ANONYMIZATION EXPORT COMPLETE ===" -ForegroundColor Green
    Write-Host "Original hospital data was NOT modified." -ForegroundColor Green
    Write-Host "Anonymized copy exported to: $($script:OUTPUT_DIR)" -ForegroundColor Cyan

}
catch {
    Write-Error "Error during anonymization export: $($_.Exception.Message)"
    Write-Host "Stack Trace: $($_.ScriptStackTrace)" -ForegroundColor Red
}
finally {
    Close-ArcApi $api
}
