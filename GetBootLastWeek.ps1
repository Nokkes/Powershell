$now = Get-Date
$lastWeek = $now.AddDays(-7)
$ts = New-TimeSpan -Start $now -End $lastWeek
for ($lastWeek -lt $now) {
    $dayStart = $lastWeek.AddMinutes(-1)
    Write-Host $dayStart
    $lastWeek.AddDays(1)
    
}