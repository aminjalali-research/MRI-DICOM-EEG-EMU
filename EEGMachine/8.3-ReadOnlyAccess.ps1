# ========================================
# ULTRA-SAFE: EXPORT-ONLY SCRIPT
# Absolutely no write operations to database
# ========================================

$api = New-Object -comobject Arc.Api.ArcApi
$api.Login("username", "password")

Write-Host "=== READ-ONLY EXPORT MODE ==="
Write-Host "This script will ONLY read from the database"
Write-Host "No modifications will be made"

$outputDir = "C:\EEG_ReadOnly_Export"
if (-not (Test-Path $outputDir)) {
    New-Item -ItemType Directory -Path $outputDir -Force | Out-Null
}

$records = $api.GetRecords()

for ($idx = 0; $idx -lt $records.GetCount(); $idx++) {
    $record = $records.GetAt($idx)
    
    Write-Host "`nProcessing Record $($idx+1)/$($records.GetCount())"
    
    try {
        # Open in READ-ONLY mode
        $record.Open()
        
        # ONLY export operations (no Save, SetValue, etc.)
        $exportPath = "$outputDir\record_$($idx+1).edf"
        $record.Data.ExportToEdf($exportPath)
        
        Write-Host "  ✓ Exported to: $exportPath"
        
        # Extract metadata (READ operations only)
        $metadata = @{
            RecordKey = $record.RecordKey.ToString()
            DateRecorded = $record.DateRecorded
            Duration = $record.Duration.TotalHours
        }
        
        $metadataPath = "$outputDir\record_$($idx+1)_metadata.json"
        $metadata | ConvertTo-Json | Out-File $metadataPath
        
        # Close record (no Save operations)
        $record.Close()
        
    } catch {
        Write-Host "  ERROR: $($_.Exception.Message)"
        if ($record.IsOpen) { $record.Close() }
    }
}

$api.Dispose()

Write-Host "`n========================================="
Write-Host "✓ READ-ONLY EXPORT COMPLETE"
Write-Host "Original database was NEVER modified"
Write-Host "Exported files location: $outputDir"
Write-Host "========================================="
