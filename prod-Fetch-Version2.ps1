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
$scope1 = $configData.scope1
$scope2 = $configData.scope2
$clientSecret = ConvertTo-SecureString -String $appSecret -AsPlainText -Force

Write-Host "Scope1: $scope1"

# Acquire the access token
$token_1 = Get-MsalToken -ClientId $appId -TenantId $tenantId -ClientSecret $clientSecret -Scopes $scope1

# Acquire token for second service (Mail.Send)
$token_2 = Get-MsalToken -ClientId $appId -TenantId $tenantId -ClientSecret $clientSecret -Scopes $scope2

# Include the access token in the headers
$mailApiHeaders = @{
  'Authorization' = "Bearer $($token_1.AccessToken)"
  'Accept'        = 'application/json'
  'Content-Type'  = 'application/json'
}

# Define Arizona time zone
$arizonaTimeZone = [System.TimeZoneInfo]::FindSystemTimeZoneById("US Mountain Standard Time")

# Current date in Arizona time
$currentDateArizona = [System.TimeZoneInfo]::ConvertTimeFromUtc([System.DateTime]::UtcNow, $arizonaTimeZone)

# Define the date range for the entire previous month in Arizona time
$previousMonthStartDate = $currentDateArizona.AddMonths(-1).Date.AddDays(1)   # Start from the 1st day of the previous month
$previousMonthEndDate = $currentDateArizona.Date.AddDays(-1)                # End on the last day of the previous month

# Create a single CSV file for the entire previous month
$consolidatedCsvFileName = "PreviousMonth_Report.csv"
$header = "Received,Subject,RecipientAddress,SenderAddress,Status"
$header | Out-File -FilePath $consolidatedCsvFileName -Force

# Loop through each mailbox to collect data for the previous month
foreach ($key in $mailboxMap.Keys) {
  $mailbox = $key
  $recipientEmail = $mailboxMap[$key]
    
  # Define the URL for the message trace endpoint with the updated date range
  $messageTraceUrl = "https://reports.office365.com/ecp/reportingwebservice/reporting.svc/MessageTrace/?`$select=$header&`$filter=StartDate ge datetime'$previousMonthStartDate' and EndDate le datetime'$previousMonthEndDate' and RecipientAddress eq '$mailbox'"
    
  try {
    # Invoke the REST API
    $response = Invoke-RestMethod -Uri $messageTraceUrl -Method Get -Headers $mailApiHeaders
  }
  catch {
    Write-Host "Error: $($_.Exception.Message)"
  }
    
  # Select only the required properties
  $reportData = $response.value | Select-Object -Property $header
    
  # Append the data to the consolidated CSV file
  if ($reportData -ne $null -and $reportData.Count -gt 0) {
    $reportData | Export-Csv -Path $consolidatedCsvFileName -NoTypeInformation -Append
  }
}

# Send a single email with the consolidated CSV file
Send-Email -recipientEmail $recipientEmail -accessToken $token_2.AccessToken -individualCsvFileName $consolidatedCsvFileName -fromEmail $emailConfigData.fromEmail

# Update the last run date
$currentDateArizona.ToString("yyyy-MM-ddTHH:mm:ss") | Out-File -FilePath ".\lastRunDate.txt"
Set-LastRunDate
