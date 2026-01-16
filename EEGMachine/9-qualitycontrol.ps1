$api = New-Object -comobject Arc.Api.ArcApi
$api.Login("username", "password")

Write-Host "=== DATA QUALITY CONTROL ==="

$records = $api.GetRecords()
$qualityReport = @()

foreach ($record in $records) {
    try {
        $record.Open()
        $data = $record.Data
        
        # Quality metrics
        $issues = @()
        
        # Check 1: Data duration
        if ($record.Duration.TotalMinutes -lt 1) {
            $issues += "Very short recording (<1 min)"
        }
        
        # Check 2: Number of channels
        $channelCount = $data.ChannelInformation.GetCount()
        if ($channelCount -lt 8) {
            $issues += "Low channel count ($channelCount)"
        }
        
        # Check 3: Data continuity
        $segments = $data.GetTimeSegments()
        $gapCount = $segments.GetCount() - 1
        if ($gapCount -gt 10) {
            $issues += "Many gaps in data ($gapCount gaps)"
        }
        
        # Check 4: Impedance quality
        $impedances = $data.GetImpedances(0, 60)  # First minute
        if ($impedances.GetCount() -gt 0) {
            $firstImpedance = $impedances.GetAt(0)
            $items = $firstImpedance.Items
            $connectedCount = 0
            for ($i = 0; $i -lt $items.GetCount(); $i++) {
                $item = $items.GetAt($i)
                if ($item.IsConnected) { $connectedCount++ }
            }
            $connectionRate = [Math]::Round(($connectedCount / $items.GetCount()) * 100, 1)
            if ($connectionRate -lt 90) {
                $issues += "Poor impedance connection rate ($connectionRate%)"
            }
        }
        
        # Check 5: Event presence
        $events = $data.GetEvents()
        if ($events.GetCount() -eq 0) {
            $issues += "No events recorded"
        }
        
        $qualityReport += [PSCustomObject]@{
            RecordKey = $record.RecordKey
            DateRecorded = $record.DateRecorded
            Duration = $record.Duration.TotalHours
            Channels = $channelCount
            DataGaps = $gapCount
            EventCount = $events.GetCount()
            Issues = ($issues -join "; ")
            QualityStatus = if ($issues.Count -eq 0) { "Good" } else { "Review" }
        }
        
        $record.Close()
        
    } catch {
        Write-Host "Error processing record: $($_.Exception.Message)"
    }
}

# Export quality report
$qualityReport | Export-Csv "C:\Output\quality_report.csv" -NoTypeInformation

# Display summary
Write-Host "`nQuality Summary:"
$goodRecords = ($qualityReport | Where-Object { $_.QualityStatus -eq "Good" }).Count
Write-Host "Good Quality: $goodRecords / $($qualityReport.Count)"
Write-Host "Need Review: $(($qualityReport.Count - $goodRecords))"

$api.Dispose()