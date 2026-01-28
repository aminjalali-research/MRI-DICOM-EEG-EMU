# ============================================
# 9. QUALITY CONTROL - Data Quality Assessment
# ============================================
# Purpose: Assess data quality across all records
# Generates quality report for clinical review

# Load configuration
. "$PSScriptRoot\0-config.ps1"

Write-Host "============================================" -ForegroundColor Cyan
Write-Host "9. QUALITY CONTROL - Data Quality Assessment" -ForegroundColor Cyan
Write-Host "============================================" -ForegroundColor Cyan

# Initialize API
$api = Initialize-ArcApi
if ($null -eq $api) {
    Write-Error "Cannot proceed without API connection"
    exit 1
}

try {
    Write-Host "`n=== DATA QUALITY CONTROL ===" -ForegroundColor Yellow
    Write-Host "Quality Thresholds:"
    Write-Host "  - Minimum Duration: $($script:QC_MIN_DURATION_MINUTES) minutes"
    Write-Host "  - Minimum Channels: $($script:QC_MIN_CHANNELS)"
    Write-Host "  - Maximum Gaps: $($script:QC_MAX_GAPS)"
    Write-Host "  - Minimum Connection Rate: $($script:QC_MIN_CONNECTION_RATE)%"

    $records = $api.GetRecords()
    $totalRecords = $records.GetCount()
    Write-Host "`nProcessing $totalRecords records..."

    $qualityReport = @()
    $goodCount = 0
    $reviewCount = 0
    $poorCount = 0

    for ($i = 0; $i -lt $totalRecords; $i++) {
        $record = $records.GetAt($i)
        
        Write-Host "`n--- Record $($i+1)/$totalRecords : $($record.RecordKey) ---" -ForegroundColor Cyan
        
        $issues = @()
        $channelCount = 0
        $gapCount = 0
        $eventCount = 0
        $connectionRate = 100
        $durationMinutes = 0
        
        try {
            $openResult = $record.Open()
            if (-not $openResult.IsSuccess) {
                Write-Host "  Failed to open record" -ForegroundColor Red
                $issues += "Cannot open record"
                
                $qualityReport += [PSCustomObject]@{
                    RecordKey = $record.RecordKey
                    DateRecorded = $record.DateRecorded
                    DurationHours = 0
                    Channels = 0
                    DataGaps = 0
                    EventCount = 0
                    ConnectionRate = 0
                    Issues = "Cannot open record"
                    QualityStatus = "Error"
                }
                continue
            }
            
            $data = $record.Data
            
            # ========================================
            # CHECK 1: Data Duration
            # ========================================
            if ($null -ne $record.Duration) {
                $durationMinutes = $record.Duration.TotalMinutes
                
                if ($durationMinutes -lt $script:QC_MIN_DURATION_MINUTES) {
                    $issues += "Very short recording ($([Math]::Round($durationMinutes, 1)) min)"
                    Write-Host "  [!] Short duration: $([Math]::Round($durationMinutes, 1)) min" -ForegroundColor Yellow
                } else {
                    Write-Host "  [OK] Duration: $([Math]::Round($durationMinutes, 1)) min" -ForegroundColor Green
                }
            } else {
                $issues += "Unknown duration"
                Write-Host "  [!] Duration: Unknown" -ForegroundColor Yellow
            }
            
            # ========================================
            # CHECK 2: Number of Channels
            # ========================================
            $channels = $data.ChannelInformation
            $channelCount = $channels.GetCount()
            
            if ($channelCount -lt $script:QC_MIN_CHANNELS) {
                $issues += "Low channel count ($channelCount)"
                Write-Host "  [!] Channels: $channelCount (below minimum)" -ForegroundColor Yellow
            } else {
                Write-Host "  [OK] Channels: $channelCount" -ForegroundColor Green
            }
            
            # ========================================
            # CHECK 3: Data Continuity (Gaps)
            # ========================================
            $segments = $data.GetTimeSegments()
            $gapCount = [Math]::Max(0, $segments.GetCount() - 1)
            
            if ($gapCount -gt $script:QC_MAX_GAPS) {
                $issues += "Many gaps in data ($gapCount gaps)"
                Write-Host "  [!] Data gaps: $gapCount (above threshold)" -ForegroundColor Yellow
            } else {
                Write-Host "  [OK] Data gaps: $gapCount" -ForegroundColor Green
            }
            
            # ========================================
            # CHECK 4: Impedance Quality
            # ========================================
            try {
                $impedances = $data.GetImpedances(0, 60)  # First minute
                
                if ($impedances.GetCount() -gt 0) {
                    $firstImpedance = $impedances.GetAt(0)
                    $items = $firstImpedance.Items
                    $connectedCount = 0
                    $totalItems = $items.GetCount()
                    
                    for ($j = 0; $j -lt $totalItems; $j++) {
                        $item = $items.GetAt($j)
                        if ($item.IsConnected) { $connectedCount++ }
                    }
                    
                    if ($totalItems -gt 0) {
                        $connectionRate = [Math]::Round(($connectedCount / $totalItems) * 100, 1)
                    }
                    
                    if ($connectionRate -lt $script:QC_MIN_CONNECTION_RATE) {
                        $issues += "Poor impedance ($connectionRate% connected)"
                        Write-Host "  [!] Impedance: $connectionRate% connected" -ForegroundColor Yellow
                    } else {
                        Write-Host "  [OK] Impedance: $connectionRate% connected" -ForegroundColor Green
                    }
                } else {
                    Write-Host "  [?] Impedance: No data available" -ForegroundColor Gray
                }
            }
            catch {
                Write-Host "  [?] Impedance: Could not check" -ForegroundColor Gray
            }
            
            # ========================================
            # CHECK 5: Event Presence
            # ========================================
            try {
                $events = $data.GetEvents()
                $eventCount = $events.GetCount()
                
                if ($eventCount -eq 0) {
                    $issues += "No events recorded"
                    Write-Host "  [!] Events: None" -ForegroundColor Yellow
                } else {
                    Write-Host "  [OK] Events: $eventCount" -ForegroundColor Green
                }
            }
            catch {
                Write-Host "  [?] Events: Could not check" -ForegroundColor Gray
            }
            
            # ========================================
            # DETERMINE QUALITY STATUS
            # ========================================
            $qualityStatus = "Good"
            if ($issues.Count -eq 0) {
                $goodCount++
                Write-Host "  Status: GOOD" -ForegroundColor Green
            } elseif ($issues.Count -le 2) {
                $qualityStatus = "Review"
                $reviewCount++
                Write-Host "  Status: REVIEW NEEDED" -ForegroundColor Yellow
            } else {
                $qualityStatus = "Poor"
                $poorCount++
                Write-Host "  Status: POOR" -ForegroundColor Red
            }
            
            # Add to report
            $qualityReport += [PSCustomObject]@{
                RecordKey = $record.RecordKey
                DateRecorded = $record.DateRecorded
                DurationHours = [Math]::Round($durationMinutes / 60, 2)
                Channels = $channelCount
                DataGaps = $gapCount
                EventCount = $eventCount
                ConnectionRate = $connectionRate
                Issues = ($issues -join "; ")
                QualityStatus = $qualityStatus
            }
            
            $record.Close()
            
        }
        catch {
            Write-Host "  Error processing record: $($_.Exception.Message)" -ForegroundColor Red
            $qualityReport += [PSCustomObject]@{
                RecordKey = $record.RecordKey
                DateRecorded = $record.DateRecorded
                DurationHours = 0
                Channels = 0
                DataGaps = 0
                EventCount = 0
                ConnectionRate = 0
                Issues = "Processing error: $($_.Exception.Message)"
                QualityStatus = "Error"
            }
        }
    }

    # ========================================
    # QUALITY SUMMARY
    # ========================================
    Write-Host "`n============================================" -ForegroundColor Cyan
    Write-Host "QUALITY CONTROL SUMMARY" -ForegroundColor Cyan
    Write-Host "============================================" -ForegroundColor Cyan
    
    Write-Host "`nTotal Records Processed: $totalRecords"
    Write-Host "  Good Quality: $goodCount ($([Math]::Round($goodCount/$totalRecords*100, 1))%)" -ForegroundColor Green
    Write-Host "  Need Review: $reviewCount ($([Math]::Round($reviewCount/$totalRecords*100, 1))%)" -ForegroundColor Yellow
    Write-Host "  Poor Quality: $poorCount ($([Math]::Round($poorCount/$totalRecords*100, 1))%)" -ForegroundColor Red

    # ========================================
    # EXPORT QUALITY REPORT
    # ========================================
    $timestamp = Get-Timestamp
    
    # CSV Report
    $csvPath = Join-Path $script:OUTPUT_DIR "quality_report_$timestamp.csv"
    $qualityReport | Export-Csv -Path $csvPath -NoTypeInformation
    Write-Host "`nQuality report exported: $csvPath" -ForegroundColor Cyan
    
    # Summary Report
    $summaryReport = @"
EEG Data Quality Control Report
Generated: $(Get-Date)
============================================

SUMMARY
-------
Total Records: $totalRecords
Good Quality: $goodCount ($([Math]::Round($goodCount/$totalRecords*100, 1))%)
Need Review: $reviewCount ($([Math]::Round($reviewCount/$totalRecords*100, 1))%)
Poor Quality: $poorCount ($([Math]::Round($poorCount/$totalRecords*100, 1))%)

QUALITY THRESHOLDS USED
-----------------------
- Minimum Duration: $($script:QC_MIN_DURATION_MINUTES) minutes
- Minimum Channels: $($script:QC_MIN_CHANNELS)
- Maximum Data Gaps: $($script:QC_MAX_GAPS)
- Minimum Connection Rate: $($script:QC_MIN_CONNECTION_RATE)%

RECORDS NEEDING ATTENTION
-------------------------
$(($qualityReport | Where-Object { $_.QualityStatus -ne "Good" } | ForEach-Object { "- $($_.RecordKey): $($_.Issues)" }) -join "`n")

Status: COMPLETE
"@
    Export-SafeCsv -Content $summaryReport -FileName "quality_summary_$timestamp.txt"

    Write-Host "`n=== QUALITY CONTROL COMPLETE ===" -ForegroundColor Green

}
catch {
    Write-Error "Error during quality control: $($_.Exception.Message)"
    Write-Host "Stack Trace: $($_.ScriptStackTrace)" -ForegroundColor Red
}
finally {
    Close-ArcApi $api
}
