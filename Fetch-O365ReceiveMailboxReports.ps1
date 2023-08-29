# Import the sendinfo.ps1 file
.\sendinfo.ps1

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

# Convert to Eastern Time (Miami
$easternZone = [System.TimeZoneInfo]::FindSystemTimeZoneById("Eastern Standard Time")
$currentDateEastern = [System.TimeZoneInfo]::ConvertTimeFromUtc($currentDateUTC, $easternZone)
$csvFileName = "$($currentDateEastern.ToString('MMMM'))_Report.csv"

# Start date is (n) days before the current date
$startDate = $currentDateUTC.AddDays(-3).ToString("yyyy-MM-ddTHH:mm:ss") + "Z"
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
    $individualCsvFileName = "${key}_${monthString}_Report.csv"

    Write-Host "Debug: Full Filename=$individualCsvFileName"

    # Fields you want to select
    $selectFields = "Received,Subject,RecipientAddress,SenderAddress,Status"

    # Define the URL for the message trace endpoint
    $messageTraceUrl = "https://reports.office365.com/ecp/reportingwebservice/reporting.svc/MessageTrace/?`$select=$selectFields&`$filter=StartDate eq datetime'$startDate' and EndDate eq datetime'$endDate' and RecipientAddress eq '$mailbox'"

    # Invoke the REST API
    $response = Invoke-RestMethod -Uri $messageTraceUrl -Method Get -Headers $mailApiHeaders

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
    if (Test-Path -Path $individualCsvFileName) {
        Write-Host "Debug: File exists. Attempting to send email."
        Send-Email -recipientEmail $recipientEmail -accessToken $token.AccessToken -csvFilePath $csvFilePath -fromEmail $emailConfigData.fromEmail
    }
    else {
        Write-Host "Debug: File does not exist. Getting child items."
        Get-ChildItem -Path . -Filter "*.csv"
    
    }
}
Write-host "Debug: Exiting loop"

# Check if it's the end of the month
if ($currentDateEastern.Day -eq [DateTime]::DaysInMonth($currentDateEastern.Year, $currentDateEastern.Month)) {
    foreach ($key in $mailboxMap.Keys) {
        Write-host "Debug: Iterating over key: $key"
        Write-Host "Debug: About to write CSV for $key to $individualCsvFileName"
        
        $recipientEmail = $mailboxMap[$key]
        Write-Host "Debug: Recipient email is $recipientEmail"

        
    }
}

# Check if the file exists before attempting to send the email
Write-Host "Debug: Checking for file $individualCsvFileName"  # Debugging line added here
    

# Check if it's a new month and if the CSV file doesn't exist
if ($currentDateEastern.Day -eq 1 -and !(Test-Path $csvFileName)) {
    # Create a new CSV file with the header, you can adjust the header based on your data
    @("Received", "Subject", "RecipientAddress", "SenderAddress", "Status") | Out-File -Path $csvFileName
}


