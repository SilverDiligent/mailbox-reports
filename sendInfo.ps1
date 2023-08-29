function Send-Email ($recipientEmail, $accessToken, $csvFilePath, $fromEmail) {
  Write-Host "Debug: Inside Send-Email, fromEmail is $fromEmail"
  # Check if the path is null or empty
  if ([string]::IsNullOrEmpty($csvFilePath)) {
    write-host "The path for the CSV file is null or empty."
    return
  }
  # Check if the the file exists
  if (-Not (Test-Path -Path $csvFilePath)) {
    write-host "The specified CSV file does not exist: $csvFilePath"
    return
  }

  write-host "Reading file from $csvFilePath"

  $uri = "https://graph.microsoft.com/v1.0/me/sendMail"

  $emailContent = @{
    message = @{
      subject      = "Monthly Report"
      from         = @{
        emailAddress = @{
          address = $fromEmail
        }
      }
      toRecipients = @(
        @{
          emailAddress = @{
            address = $recipientEmail
          }
        }
      )
      attachments  = @(
        @{
          "@odata.type" = "#microsoft.graph.fileAttachment"
          name          = [System.IO.Path]::GetFileName($csvFilePath)
          contentBytes  = [System.Convert]::ToBase64String([System.IO.File]::ReadAllBytes($csvFilePath))
        }
      )
    }
  } | ConvertTo-Json

  $headers = @{
    Authorization  = "Bearer $($accessToken)"
    "Content-Type" = "application/json"
  }

  $response = Invoke-RestMethod -Uri $uri -Headers $headers -Method Post -Body $emailContent

  
}