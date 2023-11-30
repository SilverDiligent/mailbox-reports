$est = [System.TimeZoneInfo]::FindSystemTimeZoneById('Eastern Standard Time')
$utcNow = [System.DateTime]::UtcNow
$estNow = [System.TimeZoneInfo]::ConvertTimeFromUtc($utcNow, $est)
$firstDayOfLastMonth = $estNow.AddMonths(-1).AddDays(-$estNow.Day + 1).ToString("yyyy-MM-ddTHH:mm:ss")
$lastDayOfLastMonth = $estNow.AddDays(-$estNow.Day).AddSeconds(-1).ToString("yyyy-MM-ddTHH:mm:ss")

# Output the results, noting that they are in Eastern Time
Write-Host "First day of last month (Eastern Time): $firstDayOfLastMonth"
Write-Host "Last day of last month (Eastern Time): $lastDayOfLastMonth"
