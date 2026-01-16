$api = New-Object -comobject Arc.Api.ArcApi
$api.Login("username", "password")

$records = $api.GetRecords()
$record = $records.GetAt(0)
$record.Open()
$data = $record.Data

Write-Host "=== EVENT ANALYSIS ==="

# Get all events
$allEvents = $data.GetEvents()
Write-Host "Total Events: " $allEvents.GetCount()

# Categorize events by type
$eventTypeCount = @{}
for ($i = 0; $i -lt $allEvents.GetCount(); $i++) {
    $event = $allEvents.GetAt($i)
    $type = $event.EventType
    if ($eventTypeCount.ContainsKey($type)) {
        $eventTypeCount[$type]++
    } else {
        $eventTypeCount[$type] = 1
    }
}

Write-Host "`nEvent Summary by Type:"
foreach ($type in $eventTypeCount.Keys | Sort-Object) {
    Write-Host "  $type : $($eventTypeCount[$type])"
}

# Get events in specific time window (e.g., first hour)
Write-Host "`n=== EVENTS IN FIRST HOUR ==="
$firstHourEvents = $data.GetEvents(0, 3600, $null)  # 0-3600 seconds
Write-Host "Events in first hour: " $firstHourEvents.GetCount()

for ($i = 0; $i -lt $firstHourEvents.GetCount(); $i++) {
    $event = $firstHourEvents.GetAt($i)
    Write-Host "`nEvent $($i+1):"
    Write-Host "  Type: $($event.EventType)"
    Write-Host "  Offset: $($event.Offset.TotalSeconds) sec"
    Write-Host "  Duration: $($event.Duration.TotalSeconds) sec"
    Write-Host "  Text: $($event.Text)"
    Write-Host "  Priority: $($event.Priority)"
    Write-Host "  Created by API: $($event.IsCreatedByApi)"
    Write-Host "  Can Delete: $($event.CanDelete)"
}

# Filter specific event types (e.g., seizures and spikes)
Write-Host "`n=== SEIZURE AND SPIKE EVENTS ==="
$clinicalEventTypes = @("Persyst SeizureDetected", "Persyst Spike", "Persyst SpikeBurst")
$clinicalEvents = $data.GetEvents(0, [int]::MaxValue, $clinicalEventTypes)
Write-Host "Clinical events found: " $clinicalEvents.GetCount()

# Export events to CSV
$eventsCsv = "EventType,OffsetSeconds,DurationSeconds,Text,Priority`n"
for ($i = 0; $i -lt $allEvents.GetCount(); $i++) {
    $event = $allEvents.GetAt($i)
    $eventsCsv += "$($event.EventType),$($event.Offset.TotalSeconds),$($event.Duration.TotalSeconds),`"$($event.Text)`",$($event.Priority)`n"
}
$eventsCsv | Out-File "C:\Output\events.csv"

# Get Cortical Stim Events (if applicable)
Write-Host "`n=== CORTICAL STIMULATION EVENTS ==="
$stimEvents = $data.GetCorticalStimEvents()
Write-Host "Cortical Stim Events: " $stimEvents.GetCount()
for ($i = 0; $i -lt $stimEvents.GetCount(); $i++) {
    $stim = $stimEvents.GetAt($i)
    Write-Host "`nStim Event $($i+1):"
    Write-Host "  Offset: $($stim.Offset.TotalSeconds) sec"
    Write-Host "  Current Delivered: $($stim.CurrentDelivered) mA"
    Write-Host "  Response Type: $($stim.ResponseType)"
    Write-Host "  Body Region: $($stim.BodyRegion)"
    Write-Host "  Body Side: $($stim.BodySide)"
    Write-Host "  Patient Task: $($stim.PatientTask)"
    Write-Host "  Functional Response: $($stim.FunctionalResponse)"
}

$record.Close()
$api.Dispose()