# Import the sendinfo.ps1 file
Write-Host "Debug: About to source sendinfo.ps1"
. .\sendinfo.ps1
Write-Host "Debug: Finished sourcing sendinfo.ps1"

# Import the handleLastRunDate.ps1 file
. .\handleLastRunDate.ps1

# Get the last run date
$startDate = Get-LastRunDate

# Function definitions
Function Get-LastMonthString {
  param (
    [dateTime]$date
  )
  $previousMonthDate = $date.AddMonths(-1)
  return $previousMonthDate.ToString("MMMM")
}



# Load the config.json data file
$configData = Get-Content -Path '.\config.json' | ConvertFrom-Json
$mailboxData = Get-Content -Path '.\mailboxMap.json' | ConvertFrom-Json
$emailConfigData = Get-Content -Path '.\emailConfig.json' | ConvertFrom-Json
Write-Host "Debug: Content of emailConfig: $($emailConfigData | Out-String)"


# Convert the imported JSON to a hash table
$mailboxMap = @{}
$mailboxData.PSObject.Properties | ForEach-Object { $mailboxMap[$_.Name] = $_.Value }

# Debugging: Verify the type and content of $mailboxMap
Write-Host "Type of mailboxMap: $($mailboxMap.GetType().FullName)"
Write-Host "Debug: Content of mailboxMap"
Write-Host ($mailboxMap | Out-String)

# Set the configuration parameters
$tenantId = $configData.tenantId
$appId = $configData.appId
$appSecret = $configData.clientSecretString
# $mailboxReports = $mailboxData.mailboxReports
$scope1 = $configData.scope1
$scope2 = $configData.scope2
$clientSecret = ConvertTo-SecureString -String $appSecret -AsPlainText -Force

Write-Host "Scope1: $scope1"
# Acquire the access token



# Acquire token for second service (Mail.Send)
$token_2 = Get-MsalToken -ClientId $appId -TenantId $tenantId -ClientSecret $clientSecret -Scopes $scope2

# Include the access token in the headers
$mailApiHeaders = @{
  'Authorization' = "Bearer $($token_1.AccessToken)"
  'Accept'        = 'application/json'
  'Content-Type'  = 'application/json'
}

# New code for setting headers for the Mail.Send service
$mailSendHeaders = @{
  'Authorization' = "Bearer $($token_2.AccessToken)"
  'Accept'        = 'application/json'
  'Content-Type'  = 'application/json'
}

# Current date in UTC
Write-Host "Debug: Current Date UTC: $currentDateUTC"
$currentDateUTC = [System.DateTime]::UtcNow

# Define Arizona time zone
$arizonaTimeZone = [System.TimeZoneInfo]::FindSystemTimeZoneById("US Mountain Standard Time")
$currentDateArizona = [System.TimeZoneInfo]::ConvertTimeFromUtc($currentDateUTC, $arizonaTimeZone)
# Output the Arizona time
Write-Host "Arizona time: $currentDateArizona"
# Check if today is the first day of the month
if ($currentDateArizona.Day -eq 3) {
  Write-Host "Debug: It's the 3rd day of the month, preparing to send last month's reports."

  # Get last month's string
  $previousMonth = Get-LastMonthString -date $currentDateArizona

  # Define the date range for the entire previous month in Arizona time
  $startDate = $currentDateArizona.AddMonths(-1).Date.AddDays(1)   # Start from the 1st day of the previous month
  $endDate = $currentDateArizona.Date.AddDays(-1)                # End on the last day of the previous month

  # Loop through each mailbox to check if the report exists and send the email
  foreach ($key in $mailboxMap.Keys) {
    # Create last month's CSV file name
    $lastMonthCsvFileName = "${key}_${previousMonth}_Report.csv"

    # Send last month's report
    $recipientEmail = $mailboxMap[$key]
    if (Test-Path -Path $lastMonthCsvFileName) {
      Write-host "Last month's report for $recipientEmail exists. Preparing to send email."
      Send-Email -recipientEmail $recipientEmail -accessToken $token_2.AccessToken -individualCsvFileName $lastMonthCsvFileName -fromEmail $emailConfigData.fromEmail
    }
    else {
      Write-Host "Last month's report for $recipientEmail does not exist. Skipping email send."
    }
  }
}
else {
  Write-Host "Debug: It's not the 3rd day of the month. No reports will be sent."
}

Write-Host "Appending data to existing CSV file for the month: $csvFileName"

$csvFileName = "$($currentDateArizona.ToString('MMMM'))_Report.csv"

# Start date is (n) days before the current date
$startDate = $currentDateUTC.AddDays(-7).ToString("yyyy-MM-ddTHH:mm:ss") + "Z"
# End date is the 2 days before the current date
$endDate = $currentDateUTC.AddDays(-2).ToString("yyyy-MM-ddTHH:mm:ss") + "Z"


Write-host "Debug: Entering loop"
# Loop through each email address in the hash table
foreach ($key in $mailboxMap.Keys) {
   
  # Debug lines for key and current date
  Write-Host "Debug: About to process mailbox: $key"
  Write-Host "Debug: Current key=$key"
  Write-Host "Debug: Current currentDateUTC=$currentDateUTC"

  $convertedTime = [System.TimeZoneInfo]::ConvertTimeFromUtc($currentDateUTC, $arizonaTimeZone)
  Write-Host "Debug: convertedTime is $convertedTime"

  $monthString = $convertedTime.ToString('MMMM')

  # Debug line for month part
  Write-Host "Debug: Month= $monthString"
  $mailbox = $key
  $recipientEmail = $mailboxMap[$key]
  Write-Host "Debug: About to set individualCsvFileName"
  $individualCsvFileName = "${key}_${monthString}_Report.csv"
   

  Write-Host "Debug: individualCsvFileName is $individualCsvFileName"

  # Fields you want to select
  $selectFields = "Received,Subject,RecipientAddress,SenderAddress,Status"

  # Define the URL for the message trace endpoint
  $messageTraceUrl = "https://reports.office365.com/ecp/reportingwebservice/reporting.svc/MessageTrace/?`$select=$selectFields&`$filter=StartDate eq datetime'$startDate' and EndDate eq datetime'$endDate' and RecipientAddress eq '$mailbox'"

  try {

    # Invoke the REST API
    $response = Invoke-RestMethod -Uri $messageTraceUrl -Method Get -Headers $mailApiHeaders
  }
  catch {
    write-host "Error: $($_.Exception.message)"
    write-host "Error Details: $($_.Exception.Response.Content.ReadAsStringAsync().Result)"
  }
  # Select only the required properties
  $reportData = $response.value | Select-Object -Property Received, Subject, RecipientAddress, SenderAddress, Status

  Write-Host "Debug: About to write CSV for $key"
  if ($reportData -ne $null -and $reportData.Count -gt 0) {
    $reportData | Export-Csv -Path $individualCsvFileName -NoTypeInformation -Append
    Write-Host "Debug: Finished writing CSV for $key"
  }
  else {
    Write-Host "Debug: No data to write for $key"
  }
   
  # Check if the file exists before attempting to send the email
  $fileExists = Test-Path -Path $individualCsvFileName  # Assign the result to $fileExists
  if ($fileExists -and $currentDateArizona.Day -eq 3) {
    Write-Host "Debug: File exists and it's the 3rd of the month. About to send email."
    Write-Host "Debug: Inside Send-Email, csvFilePath is $individualcsvFileName"
    Send-Email -recipientEmail $recipientEmail -accessToken $token_2.AccessToken -individualCsvFileName $individualCsvFileName -fromEmail $emailConfigData.fromEmail
  }
  else {
    Write-Host "Debug: File does not exist. Skipping email send."
    Get-ChildItem -Path . -Filter "*.csv"
  }

}
Write-host "Debug: Exiting loop"

# Check if it's the 3rd day of the new month
if ($currentDateArizona.Day -eq 3) {
  Write-Host "Debug: It's the 3rd day of the month, preparing to send last month's reports."

  # Get last month's string
  $previousMonth = Get-LastMonthString -date $currentDateArizona
     

  # Loop through each mail ID to check if the report exists and send the email
  foreach ($key in $mailboxMap.Keys) {
    # Create last month's CSV file name
    $lastMonthCsvFileName = "${key}_${previousMonth}_Report.csv"

    # Create this month's CSV file name
    $individualCsvFileName = "${key}_${monthString}_Report.csv"

    # Create a new CSV for this month
    $header = "Received,Subject,RecipientAddress,SenderAddress,Status"
    $header | Out-File -FilePath $individualCsvFileName -Force
    Write-Host "Created new CSV file for the month and mailbox: $individualCsvFileName"

    # Send last month's report
    $recipientEmail = $mailboxMap[$key]
    if (Test-Path -Path $lastMonthCsvFileName) {
      Write-host "Last month's report for $recipientEmail exists. Preparing to send email."
      Send-Email -recipientEmail $recipientEmail -accessToken $token_2.AccessToken -individualCsvFileName $lastMonthCsvFileName -fromEmail $emailConfigData.fromEmail
    }
    else {
      Write-Host "Last month's report for $recipientEmail does not exist. Skipping email send."
    }
  }
}

$currentDateUTC.ToString("yyyy-MM-ddTHH:mm:ss") | Out-File -FilePath ".\lastRunDate.txt"

# Set the last run date
Set-LastRunDate
