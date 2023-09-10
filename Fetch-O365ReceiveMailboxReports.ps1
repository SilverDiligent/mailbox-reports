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

# New code for setting headers for the Mail.Send service
$mailSendHeaders = @{
    'Authorization' = "Bearer $($token_2.AccessToken)"
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
    if ($fileExists) {
        Write-Host "Debug: File exists. About to send email."
        Write-Host "Debug: Inside Send-Email, csvFilePath is $individualcsvFileName"
        Send-Email -recipientEmail $recipientEmail -accessToken $token.AccessToken -individualCsvFileName $individualCsvFileName -fromEmail $emailConfigData.fromEmail
    }
    else {
        Write-Host "Debug: File does not exist. Skipping email send."
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

Function Get-LastMonthString {
    param (
        [dateTime]$date
    )
    $previousMonthDate = $date.AddMonths(-1)
    return $previousMonthDate.ToString("MMMM")
}

# Check if the file exists before attempting to send the email
Write-Host "Debug: Checking for file $individualCsvFileName"  # Debugging line added here
    

# Check if it's the first day of the new month
if ($currentDateEastern.Day -eq 9) {
    Write-Host "Debug: It's the SECOND! day of the month, preparing to send last month's reports."

    # Get last month's string
    $previousMonth = Get-LastMonthString -date $currentDateEastern
        

    # Loop through each mail ID to check if the report exists and send the email
    foreach ($key in $mailboxMap.Keys) {
        $fromEmail = $key
        $recipientEmail = $mailboxMap[$key]

        Write-Host "About to send email to $recipientEmail from $fromEmail."

        # Check if the file exists before attempting to send emails.
        if (Test-Path -Path $individualCsvFileName) {
            Write-host "Debug: Last month's report for $recipientEmail exists. Preparing to send emails."
            write-Host "Debug: About to call Send-Email with CSV: $lastMonthCsvFileName"
            Send-Email -recipientEmail $recipientEmail -accessToken $token_2.AccessToken -individualCsvFileName $individualCsvFileName -fromEmail $emailConfigData.fromEmail
        }
        else {
            Write-Host "Debug: Last month's report for $recipientEmail does not exist. Skipping email send."
        }
    }
}