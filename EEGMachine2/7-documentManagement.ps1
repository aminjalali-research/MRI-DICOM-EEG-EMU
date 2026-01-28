# ============================================
# 7. DOCUMENT MANAGEMENT
# ============================================
# Purpose: Manage and export associated documents
# Handles EEG reports, notes, and attachments

# Load configuration
. "$PSScriptRoot\0-config.ps1"

param(
    [int]$RecordIndex = 0,
    [string]$ExportPath = $null
)

Write-Host "============================================" -ForegroundColor Cyan
Write-Host "7. DOCUMENT MANAGEMENT" -ForegroundColor Cyan
Write-Host "============================================" -ForegroundColor Cyan

# Set export path
if ([string]::IsNullOrEmpty($ExportPath)) {
    $ExportPath = Join-Path $script:OUTPUT_DIR "Documents"
}

# Create export directory
if (-not (Test-Path $ExportPath)) {
    New-Item -ItemType Directory -Path $ExportPath -Force | Out-Null
}
Write-Host "Export directory: $ExportPath" -ForegroundColor Gray

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
    Write-Host "Managing documents for record: $($record.RecordKey)" -ForegroundColor Yellow

    # ========================================
    # GET ASSOCIATED DOCUMENTS
    # ========================================
    Write-Host "`n=== ASSOCIATED DOCUMENTS ===" -ForegroundColor Yellow
    
    $documents = $null
    try {
        $documents = $data.GetAssociatedDocuments()
        $docCount = $documents.GetCount()
        Write-Host "Found $docCount associated document(s)"
    }
    catch {
        Write-Host "Could not retrieve associated documents: $($_.Exception.Message)" -ForegroundColor Yellow
        $docCount = 0
    }

    $documentList = @()
    
    if ($docCount -gt 0) {
        Write-Host "`n--- DOCUMENT LIST ---" -ForegroundColor Cyan
        
        for ($i = 0; $i -lt $docCount; $i++) {
            $doc = $documents.GetAt($i)
            
            Write-Host "`nDocument $($i+1):" -ForegroundColor Cyan
            Write-Host "  File Name: $($doc.FileName)"
            Write-Host "  Key: $($doc.Key)"
            
            $documentList += @{
                Index = $i + 1
                FileName = $doc.FileName
                Key = $doc.Key
                Exported = $false
                ExportPath = ""
            }
            
            # Export document
            try {
                $exportFileName = "$($record.RecordKey)_$($doc.FileName)"
                $exportFilePath = Join-Path $ExportPath $exportFileName
                
                $exportResult = $doc.Export($exportFilePath)
                
                if ($exportResult.IsSuccess) {
                    Write-Host "  Exported to: $exportFilePath" -ForegroundColor Green
                    $documentList[-1].Exported = $true
                    $documentList[-1].ExportPath = $exportFilePath
                } else {
                    Write-Host "  Export failed: $($exportResult.ErrorMessage)" -ForegroundColor Red
                }
            }
            catch {
                Write-Host "  Export error: $($_.Exception.Message)" -ForegroundColor Red
            }
        }
    } else {
        Write-Host "No associated documents found for this record" -ForegroundColor Yellow
    }

    # ========================================
    # RECORD NOTES AND CLINICAL DATA
    # ========================================
    Write-Host "`n=== RECORD CLINICAL NOTES ===" -ForegroundColor Yellow
    
    $clinicalNotes = @()
    
    try {
        $recordFields = $record.Fields
        $defaultKeys = $record.DefaultFieldDefinitionKeys
        
        # Common clinical fields to extract
        $clinicalFieldKeys = @(
            @{ Name = "Medications"; Key = $defaultKeys.Medications },
            @{ Name = "Clinical History"; Key = $defaultKeys.ClinicalHistory },
            @{ Name = "Indication"; Key = $defaultKeys.Indication },
            @{ Name = "Impression"; Key = $defaultKeys.Impression },
            @{ Name = "Description"; Key = $defaultKeys.Description },
            @{ Name = "Interpretation"; Key = $defaultKeys.Interpretation }
        )
        
        foreach ($fieldInfo in $clinicalFieldKeys) {
            if ($null -ne $fieldInfo.Key) {
                try {
                    $field = $recordFields.GetField($fieldInfo.Key)
                    if ($null -ne $field -and $null -ne $field.Value -and -not [string]::IsNullOrEmpty($field.Value.DisplayText)) {
                        Write-Host "`n$($fieldInfo.Name):" -ForegroundColor Cyan
                        $text = $field.Value.DisplayText
                        Write-Host "  $text"
                        
                        $clinicalNotes += @{
                            FieldName = $fieldInfo.Name
                            Value = $text
                        }
                    }
                }
                catch {
                    # Field may not exist
                }
            }
        }
    }
    catch {
        Write-Host "Could not retrieve clinical notes: $($_.Exception.Message)" -ForegroundColor Yellow
    }

    # ========================================
    # EXPORT DOCUMENT MANIFEST
    # ========================================
    $timestamp = Get-Timestamp
    $recordKey = $record.RecordKey -replace '[^a-zA-Z0-9]', '_'
    
    # Document manifest CSV
    if ($documentList.Count -gt 0) {
        $manifestCsv = "Index,FileName,DocumentKey,Exported,ExportPath`n"
        foreach ($doc in $documentList) {
            $manifestCsv += "$($doc.Index),`"$($doc.FileName)`",$($doc.Key),$($doc.Exported),`"$($doc.ExportPath)`"`n"
        }
        Export-SafeCsv -Content $manifestCsv -FileName "documents_manifest_${recordKey}_$timestamp.csv"
    }
    
    # Clinical notes export
    if ($clinicalNotes.Count -gt 0) {
        $notesContent = "Record: $($record.RecordKey)`nDate: $(Get-Date)`n`n"
        foreach ($note in $clinicalNotes) {
            $notesContent += "=== $($note.FieldName) ===`n$($note.Value)`n`n"
        }
        Export-SafeCsv -Content $notesContent -FileName "clinical_notes_${recordKey}_$timestamp.txt"
    }
    
    # Summary
    $summary = @"
Document Management Summary
Generated: $(Get-Date)
Record: $($record.RecordKey)
============================================

Associated Documents: $docCount
Documents Exported: $(($documentList | Where-Object { $_.Exported }).Count)
Clinical Note Fields: $($clinicalNotes.Count)

Export Location: $ExportPath

Status: SUCCESS
"@
    Export-SafeCsv -Content $summary -FileName "documents_summary_${recordKey}_$timestamp.txt"

    Write-Host "`n=== DOCUMENT MANAGEMENT COMPLETE ===" -ForegroundColor Green
    Write-Host "Documents processed: $docCount"
    Write-Host "Clinical notes extracted: $($clinicalNotes.Count)"

}
catch {
    Write-Error "Error during document management: $($_.Exception.Message)"
    Write-Host "Stack Trace: $($_.ScriptStackTrace)" -ForegroundColor Red
}
finally {
    if ($null -ne $record -and $record.IsOpen) {
        $record.Close()
    }
    Close-ArcApi $api
}
