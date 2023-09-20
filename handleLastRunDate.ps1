# Function to get the last run date
function Get-LastRunDate {
  if (Test-Path -Path "lastRunDate.txt") {
    $lastRunDate = Get-Content -Path "lastRunDate.txt"
    if ($null -eq $lastRunDate -or $lastRunDate -eq '') {
      # Fallback logic if the file is empty or unreadable
      return (Get-Date).AddDays(-7).ToString("yyyy-MM-ddTHH:mm:ss") + "Z"
    }
    else {
      return [datetime]::ParseExact($lastRunDate, "yyyy-MM-ddTHH:mm:ss", $null).ToString("yyyy-MM-ddTHH:mm:ss") + "Z"
    }
  }
  else {
    # Fallback logic if the file does not exist
    return (Get-Date).AddDays(-7).ToString("yyyy-MM-ddTHH:mm:ss") + "Z"
  }
}


# Function to set the last run date
function Set-LastRunDate {
  (Get-Date).ToString("yyyy-MM-ddTHH:mm:ss") | Out-File -FilePath ".\lastRunDate.txt"
}
