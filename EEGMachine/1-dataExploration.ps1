# Initialize API
$api = New-Object -comobject Arc.Api.ArcApi
$api.Login("username", "password")

Write-Host "=== SYSTEM OVERVIEW ==="
Write-Host "API Version: " $api.Version
Write-Host "Current User: " $api.CurrentUserName
Write-Host "Total Active Records: " $api.GetRecordCount()

# Get record counts by year
$recordsByYear = $api.GetRecordCountsByYear()
Write-Host "`n=== RECORDS BY YEAR ==="
for ($i = 0; $i -lt $recordsByYear.GetCount(); $i++) {
    $yearData = $recordsByYear.GetAt($i)
    Write-Host "Year $($yearData.Year): $($yearData.RecordCount) records"
}

# Check licensed features
Write-Host "`n=== LICENSED FEATURES ==="
$eegFeatures = @(
    "Video", "Highlights", "RemoteControl", "Sentinel", 
    "Oximetry", "Wireless", "ArtifactReductionReview"
)
foreach ($feature in $eegFeatures) {
    $isLicensed = $api.IsEegFeatureLicensed([Arc.Api.LicenseFeatureEEG]::$feature)
    Write-Host "$feature : $isLicensed"
}

$api.Dispose()