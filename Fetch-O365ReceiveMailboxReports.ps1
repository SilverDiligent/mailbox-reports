# Load the config.json data file
$configData = Get-Content -Path '.\config.json' | ConvertFrom-Json
$mailboxData = Get-Content -Path '.\mailboxMap.json' | ConvertFrom-Json

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
    'Accept' = 'application/json'
    'Content-Type' = 'application/json'
}

# Current date in UTC
$currentDateUTC = [System.DateTime]::UtcNow

# Start date is (n) days before the current date
$startDate = $currentDateUTC.AddDays(-3).ToString("yyyy-MM-ddTHH:mm:ss") + "Z"

# End date is the 2 days before the current date
$endDate = $currentDateUTC.AddDays(-2).ToString("yyyy-MM-ddTHH:mm:ss") + "Z"

# Define CSV file name based on the current month
$csvFileName = $currentDateUTC.ToString("MMMM") + "_Report.csv"

# Check if it's a new month and if the CSV file doesn't exist
if ((Get-Date).Day -eq 1 -and !(Test-Path $csvFileName)) {
    # Create a new CSV file with the header, you can adjust the header based on your data
    @("Received", "Subject", "RecipientAddress", "SenderAddress", "Status") | Out-File -Path $csvFileName
}

# Loop through each email address in the hash table
foreach ($key in $mailboxMap.Keys) {
    write-host "Debug: Current key is $key"
    $mailbox = $key
    $recipientEmail = $mailboxMap[$key]

    # Fields you want to select
    $selectFields = "Received,Subject,RecipientAddress,SenderAddress,Status"

    # Define the URL for the message trace endpoint
    $messageTraceUrl = "https://reports.office365.com/ecp/reportingwebservice/reporting.svc/MessageTrace/?`$select=$selectFields&`$filter=StartDate eq datetime'$startDate' and EndDate eq datetime'$endDate' and RecipientAddress eq '$mailbox'"

    # Invoke the REST API
    $response = Invoke-RestMethod -Uri $messageTraceUrl -Method Get -Headers $mailApiHeaders

    # Select only the required properties
$reportData = $response.value | Select-Object -Property Received, Subject, RecipientAddress, SenderAddress, Status

# Export to a CSV file
$individualCsvFileName = "${key}_${currentDateUTC.ToString("MMMM")}_Report.csv"
$reportData | Export-Csv -Path $individualCsvFileName -NoTypeInformation -Append

# Check if it's the end of the month
if ((Get-Date).Day -eq [DateTime]::DaysInMonth((Get-Date).Year, (Get-Date).Month)) {
        foreach ($key in $mailboxMap.Keys) {
            $individualCsvFileName = "${key}_${currentDateUTC.ToString("MMMM")}_Report.csv"
            $recipientEmail = $mailboxMap[$key]
            Send-GraphEmail -recipientEmail $recipientEmail -subject "Monthly Report" -body "Here is your monthly report." -attachmentPath $individualCsvFileName -accessToken $token.AccessToken
    }
}
}