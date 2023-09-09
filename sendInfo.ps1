function Send-Email ($recipientEmail, $accessToken, $individualCsvFileName, $fromEmail) {
  Write-Host "Debug: Inside Send-Email, individualCsvFileName is $individualCsvFileName"
  Write-Host "Debug: Inside Send-Email, fromEmail is $fromEmail"
  # Check if the path is null or empty
  if ([string]::IsNullOrEmpty($individualCsvFileName)) {
    write-host "The path for the CSV file is null or empty."
    return
  }
  # Check if the the file exists
  if (-Not (Test-Path -Path $individualCsvFileName)) {
    write-host "The specified CSV file does not exist: $individualCsvFileName"
    return
  }

  write-host "Reading file from $individualCsvFileName"
  Write-Host "Debug: URI is $uri"
  Write-Host "Debug: Token is $token_2"
  
  $uri = "https://graph.microsoft.com/v1.0/users/$fromEmail/sendMail"

  $emailContent = @{
    message         = @{
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
          name          = [System.IO.Path]::GetFileName($individualCsvFileName)
          contentBytes  = [System.Convert]::ToBase64String([System.IO.File]::ReadAllBytes($individualCsvFileName))
        }
      )
    }
    saveToSentItems = $false
  } | ConvertTo-Json -Depth 4

  Write-Host "Debug: URI is $uri"
  Write-Host "Debug: Token is $AccessToken"

  $headers = @{
    Authorization  = "Bearer $($token_2.AccessToken)"
    "Content-Type" = "application/json"

  }
  # Add this line before sending the email
  Write-Host "About to send email to $recipientEmail from $fromEmail."
  $response = Invoke-RestMethod -Uri $uri -Headers $headers -Method Post -Body $emailContent -ContentType "application/json"

  # Add this line after sending the email
  Write-Host "Email sent to $recipientEmail from $fromEmail."
  return $response

  
}

