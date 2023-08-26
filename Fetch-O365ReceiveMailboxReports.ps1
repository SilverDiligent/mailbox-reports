# Load the config.json data file
$configData = Get-Content -Path '.\config.json' | ConvertFrom-Json
$mailboxData = Get-Content -Path '.\mailboxIds.json' | ConvertFrom-Json

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

# Start date is 9 days before the current date
$startDate = $currentDateUTC.AddDays(-8).ToString("yyyy-MM-ddTHH:mm:ss") + "Z"

# End date is the 2 days before the current date
$endDate = $currentDateUTC.AddDays(-7).ToString("yyyy-MM-ddTHH:mm:ss") + "Z"

# Define CSV file name based on the current month
$csvFileName = $currentDateUTC.ToString("MMMM") + "_Report.csv"

# Check if it's a new month and if the CSV file doesn't exist
if ((Get-Date).Day -eq 1 -and !(Test-Path $csvFileName)) {
    # Create a new CSV file with the header, you can adjust the header based on your data
    @("Received", "Subject", "RecipientAddress", "SenderAddress", "Status") | Out-File -Path $csvFileName
}


# Iterate through the mailbox reports
foreach ($report in $mailboxData.mailboxReports) {
    $senderEmail = $report.mailboxId
    $recipientEmail = $report.recipient

    # Fields you want to select
    $selectFields = "Received,Subject,RecipientAddress,SenderAddress,Status"

    # Define the URL for the message trace endpoint
    $messageTraceUrl = "https://reports.office365.com/ecp/reportingwebservice/reporting.svc/MessageTrace/?`$select=$selectFields&`$filter=StartDate eq datetime'$startDate' and EndDate eq datetime'$endDate' and RecipientAddress eq 'alexisc@alexislab.com'"

    # Invoke the REST API
    $response = Invoke-RestMethod -Uri $messageTraceUrl -Method Get -Headers $mailApiHeaders

    # Select only the required properties
$reportData = $response.value | Select-Object -Property Received, Subject, RecipientAddress, SenderAddress, Status

# Export to a CSV file
$reportData | Export-Csv -Path 'report.csv' -NoTypeInformation -Append

    # Loop through each message
    $reportContent = "Report for $senderEmail`n`n"
    foreach ($message in $response) {
        $reportContent += "Sender: $($message.SenderAddress)`nRecipient: $($message.RecipientAddress)`nSubject: $($message.Subject)`nReceived: $($message.Received)`nStatus: $($message.Status)`n`n"
}

    # Send the report to the recipient (e.g., via email)
    Send-Report -Recipient $recipientEmail -Content $reportContent
}

# Function to send the report (you can implement this based on your specific requirements)
function Send-Report {
    param (
        [string]$Recipient,
        [string]$Content
    )

    # Logic to send the report (e.g., via email) to the recipient
    # ...
} Process the response as needed
