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

$est = [System.TimeZoneInfo]::FindSystemTimeZoneById('Eastern Standard Time')
$utcNow = [System.DateTime]::UtcNow
$estNow = [System.TimeZoneInfo]::ConvertTimeFromUtc($utcNow, $est)
$firstDayOfLastMonth = $estNow.AddMonths(-1).AddDays(-$estNow.Day + 1).ToString("yyyy-MM-ddTHH:mm:ss") + "Z"
$lastDayOfLastMonth = $estNow.AddDays(-$estNow.Day).AddSeconds(-1).ToString("yyyy-MM-ddTHH:mm:ss") + "Z"

foreach ($report in $mailboxReports) {
  $mailboxId = $report.mailboxId
  $reportRecipient = $report.recipient
  $mailApiUrl = "https://graph.microsoft.com/v1.0/users/$mailboxId/mailFolders/inbox/messages?`$filter=receivedDateTime ge $firstDayOfLastMonth and receivedDateTime le $lastDayOfLastMonth"
  $csvPath = "./MailboxData_$mailboxId.csv"
  $mailData = @()

  do {
    try {
      $mailResponse = Invoke-RestMethod -Uri $mailApiUrl -Method GET -Headers $mailApiHeaders
      $mailData += $mailResponse.value | Select subject, receivedDateTime, @{name='from';expression={$_.from.emailAddress.address}}, @{name='toRecipients';expression={$_.toRecipients.emailAddress.address}}, sentDateTime
      $mailApiUrl = $null
      if ($mailResponse.'@odata.nextLink') {
          $mailApiUrl = $mailResponse.'@odata.nextLink'
      }
    }
    catch {
      if ($_.Exception.Response.StatusCode -eq 'Unauthorized') {
          $token = Get-MsalToken -ClientId $appId -TenantId $tenantId -ClientSecret $clientSecret -Scopes $scope
          $mailApiHeaders = @{
              'Authorization' = "Bearer $($token.AccessToken)"
              'Accept' = 'application/json'
          }
      }
      else {
          throw $_
      }
    }
  } while ($mailApiUrl)

  $mailData | Export-Csv -Path $csvPath -NoTypeInformation

  $csvContent = [System.IO.File]::ReadAllBytes($csvPath)
  $csvBase64 = [System.Convert]::ToBase64String($csvContent)

  $emailJsonPayload = @{
    'message' = @{
        'subject' = "Mailbox Data"
        'body' = @{
            'contentType' = "Text"
            'content' = "Attached is your mailbox data for last month."
        }
        'from' = @{
            'emailAddress' = @{
                'address' = $mailboxId
            }
        }
        'toRecipients' = @(
            @{
                'emailAddress' = @{
                    'address' = $reportRecipient
                }
            }
        )
        'attachments' = @(
            @{
                '@odata.type' = "#microsoft.graph.fileAttachment"
                'name' = "MailboxData_$mailboxId.csv"
                'contentType' = "text/csv"
                'contentBytes' = $csvBase64
            }
        )
    }
    'saveToSentItems' = "true"
  } | ConvertTo-Json -Depth 4

  $sendMailUrl = "https://graph.microsoft.com/v1.0/users/$mailboxId/sendMail"
  $sendMailResponse = Invoke-RestMethod -Uri $sendMailUrl -Method POST -Headers $mailApiHeaders -Body $emailJsonPayload -ContentType 'application/json'

  Remove-Item $csvPath
}
