# ============================================
# 1. DATA EXPLORATION - System Overview
# ============================================
# Purpose: Get system overview, record counts, and licensed features
# Run this first to verify API connectivity

# Load configuration
. "$PSScriptRoot\0-config.ps1"

Write-Host "============================================" -ForegroundColor Cyan
Write-Host "1. DATA EXPLORATION - System Overview" -ForegroundColor Cyan
Write-Host "============================================" -ForegroundColor Cyan

# Initialize API
$api = Initialize-ArcApi
if ($null -eq $api) {
    Write-Error "Cannot proceed without API connection"
    exit 1
}

try {
    Write-Host "`n=== SYSTEM OVERVIEW ===" -ForegroundColor Yellow
    Write-Host "API Version: $($api.Version)"
    Write-Host "Current User: $($api.CurrentUserName)"
    Write-Host "User Key: $($api.CurrentUserKey)"
    Write-Host "Total Active Records: $($api.GetRecordCount())"
    
    # Check if any records are open locally
    Write-Host "Records Open Locally: $($api.AreRecordsOpenLocally())"

    # Get record counts by year
    $recordsByYear = $api.GetRecordCountsByYear()
    Write-Host "`n=== RECORDS BY YEAR ===" -ForegroundColor Yellow
    
    if ($recordsByYear.GetCount() -gt 0) {
        for ($i = 0; $i -lt $recordsByYear.GetCount(); $i++) {
            $yearData = $recordsByYear.GetAt($i)
            Write-Host "  Year $($yearData.Year): $($yearData.RecordCount) records"
        }
    } else {
        Write-Host "  No records found in system" -ForegroundColor Yellow
    }

    # Check licensed EEG features
    Write-Host "`n=== LICENSED EEG FEATURES ===" -ForegroundColor Yellow
    $eegFeatures = @(
        "Video", 
        "Highlights", 
        "RemoteControl", 
        "Sentinel", 
        "Oximetry", 
        "Wireless", 
        "ArtifactReductionReview",
        "ArtifactReductionAcquisition",
        "Easy3Amplifier",
        "RoomAutomation",
        "SatelliteView"
    )
    
    foreach ($feature in $eegFeatures) {
        try {
            $isLicensed = $api.IsEegFeatureLicensed([Arc.Api.LicenseFeatureEEG]::$feature)
            $status = if ($isLicensed) { "[YES]" } else { "[NO]" }
            $color = if ($isLicensed) { "Green" } else { "Gray" }
            Write-Host "  $status $feature" -ForegroundColor $color
        }
        catch {
            Write-Host "  [?] $feature - Unable to check" -ForegroundColor Yellow
        }
    }

    # Check licensed Synopsis features
    Write-Host "`n=== LICENSED SYNOPSIS FEATURES ===" -ForegroundColor Yellow
    $synopsisFeatures = @(
        "AmplitudeIntegratedEEGAnalyzer",
        "BandPowerAnalyzer",
        "EnvelopeAsymmetryAnalyzer",
        "EnvelopeEEGAnalyzer",
        "EventDetectionFeature",
        "PowerRatioAnalyzer",
        "SpectralEntropyAnalyzer",
        "SpectrogramAnalyzer",
        "TrendingFeature"
    )
    
    foreach ($feature in $synopsisFeatures) {
        try {
            $isLicensed = $api.IsSynopsisFeatureLicensed([Arc.Api.LicenseFeatureSynopsis]::$feature)
            $status = if ($isLicensed) { "[YES]" } else { "[NO]" }
            $color = if ($isLicensed) { "Green" } else { "Gray" }
            Write-Host "  $status $feature" -ForegroundColor $color
        }
        catch {
            Write-Host "  [?] $feature - Unable to check" -ForegroundColor Yellow
        }
    }

    # Get available montages
    Write-Host "`n=== AVAILABLE MONTAGES ===" -ForegroundColor Yellow
    $montages = $api.GetMontages()
    if ($montages.GetCount() -gt 0) {
        for ($i = 0; $i -lt $montages.GetCount(); $i++) {
            $montage = $montages.GetAt($i)
            Write-Host "  - $($montage.Name)"
        }
    } else {
        Write-Host "  No montages available" -ForegroundColor Yellow
    }

    # Export system overview
    $timestamp = Get-Timestamp
    $overviewReport = @"
EEG System Overview Report
Generated: $(Get-Date)
============================================

API Information:
- Version: $($api.Version)
- User: $($api.CurrentUserName)
- Total Records: $($api.GetRecordCount())

Records by Year:
$(for ($i = 0; $i -lt $recordsByYear.GetCount(); $i++) {
    $yearData = $recordsByYear.GetAt($i)
    "  $($yearData.Year): $($yearData.RecordCount) records"
})

Status: SUCCESS
"@

    Export-SafeCsv -Content $overviewReport -FileName "system_overview_$timestamp.txt"
    
    Write-Host "`n=== DATA EXPLORATION COMPLETE ===" -ForegroundColor Green

}
catch {
    Write-Error "Error during data exploration: $($_.Exception.Message)"
    Write-Host "Stack Trace: $($_.ScriptStackTrace)" -ForegroundColor Red
}
finally {
    Close-ArcApi $api
}
