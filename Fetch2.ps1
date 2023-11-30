# Import the sendinfo.ps1 file
Write-Host "Debug: About to source sendinfo.ps1"
. .\sendinfo.ps1
Write-Host "Debug: Finished sourcing sendinfo.ps1"


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
$mailboxReports = $mailboxData.mailboxReports
$scope = $configData.scope
$clientSecret = ConvertTo-SecureString -String $appSecret -AsPlainText -Force

# Acquire the access token
$token = Get-MsalToken -ClientId $appId -TenantId $tenantId -ClientSecret $clientSecret -Scopes $scope

# Include the access token in the headers
$mailApiHeaders = @{
  'Authorization' = "Bearer $($token.AccessToken)"
  'Accept'        = 'application/json'
  'Content-Type'  = 'application/json'
}

# Current date in UTC
$currentDateUTC = [System.DateTime]::UtcNow
# $currentDateUTC = Get-Date "2023-10-02 00:00:00Z" # Uncomment this for dry-run
$currentDateUTC = [DateTime]::SpecifyKind($currentDateUTC, [System.DateTimeKind]::Utc)  # # Set DateTimeKind to Utc

# Convert to Eastern Time (Miami
$easternZone = [System.TimeZoneInfo]::FindSystemTimeZoneById("Eastern Standard Time")
$currentDateEastern = [System.TimeZoneInfo]::ConvertTimeFromUtc($currentDateUTC, $easternZone)
$csvFileName = "$($currentDateEastern.ToString('MMMM'))_Report.csv"

# Start date is (n) days before the current date
$startDate = $currentDateUTC.AddDays(-7).ToString("yyyy-MM-ddTHH:mm:ss") + "Z"
# End date is the 2 days before the current date
$endDate = $currentDateUTC.AddDays(-2).ToString("yyyy-MM-ddTHH:mm:ss") + "Z"
# $csvFileName = $currentDateEastern.ToString("MMMM") + "_Report.csv"
$currentDateEastern = [System.TimeZoneInfo]::ConvertTimeFromUtc($currentDateEastern, $easternZone)


Write-host "Debug: Entering loop"
# Loop through each email address in the hash table
foreach ($key in $mailboxMap.Keys) {
    
  # Debug lines for key and current date
  Write-Host "Debug: About to process mailbox: $key"
  Write-Host "Debug: Current key=$key"
  Write-Host "Debug: Current currentDateUTC=$currentDateUTC"

  $convertedTime = [System.TimeZoneInfo]::ConvertTimeFromUtc($currentDateUTC, $easternZone)
  $monthString = $convertedTime.ToString('MMMM')

  # Debug line for month part
  Write-Host "Debug: Month= $monthString"
  $mailbox = $key
  $recipientEmail = $mailboxMap[$key]
  Write-Host "Debug: About to set individualCsvFileName"
  # $individualCsvFileName_old = "${key}_${monthString}_Report.csv"
  $individualCsvFileName = ".\alexisc@alexislab.com_September_Report.csv"

  Write-Host "Debug: individualCsvFileName is $individualCsvFileName"

  # Fields you want to select
  $selectFields = "Received,Subject,RecipientAddress,SenderAddress,Status"

  # Define the URL for the message trace endpoint
  $messageTraceUrl = "https://reports.office365.com/ecp/reportingwebservice/reporting.svc/MessageTrace/?`$select=$selectFields&`$filter=StartDate eq datetime'$startDate' and EndDate eq datetime'$endDate' and RecipientAddress eq '$mailbox'"

  # Invoke the REST API
  $response = Invoke-RestMethod -Uri $messageTraceUrl -Method Get -Headers $mailApiHeaders

  # Select only the required properties
  $reportData = $response.value | Select-Object -Property Received, Subject, RecipientAddress, SenderAddress, Status

  $reportData | Export-Csv -Path $individualCsvFileName -NoTypeInformation -Append
}

if ($currentDateEastern.Day -eq 6) {
  Write-Host "Debug: Sending Emails via GRAPH on the 4th day of the month"
      
  # Loop to send emails to all addresses
  foreach ($recipientEmail in $mailboxMap.Values) {

    Write-Host "Debug: About to call Send-Email, individualCsvFileName is $individualCsvFileName"
    Send-Email -recipientEmail $recipientEmail -accessToken $token.AccessToken -individualCsvFileName $individualCsvFileName -fromEmail $emailConfigData.fromEmail
  }
}
  