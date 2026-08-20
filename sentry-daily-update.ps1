<#
  .SYNOPSIS
  A Powershell script to download user data from a pre-prepared Alma Analytics report and convert it to CSV format suitable for importing into Sentry-Isis via its "Daily Update" function

  .DESCRIPTION
  The Sentry-Isis building access management system can import data from a number of different sources as long as the format adheres to the program's requirements. CSV is the format chosen here.
  This script is designed to be run as a scheduled task shortly before Sentry's Daily Update is scheduled to run, ensuring that Sentry imports a fresh set of data from a CSV file, courtesy of this script.

  .PARAMETER ApiRegion
    Specifies the API region code. Available codes are:
      Asia Pacific = ap
      Canada = ca
      China = cn
      Europe = eu (default)
      North America = na

      Note that when this parameter is specified, the BaseUrl will be automatically modified to include the ApiRegion specified.
      By default this is set to eu [String]

  .PARAMETER BaseUrl
    This sets the Alma API base URL. You will probably not need to explicitly set this at runtime. By default this is set to https://api-eu.hosted.exlibrisgroup.com [String] [Optional]

  .PARAMETER AnalyticsApiBasePath
    This sets the Alma Analytics API base Path. You will probably not need to explicitly set this at runtime. By default this is set to /almaws/v1/analytics/reports [String] [Optional]

  .PARAMETER AlmaApiKeyIdentifier
    The identifier for the API key to use when authenticating to the Alma API. [string]

  .PARAMETER EmailRecipient
    This sets the destination email address for failure related emails to go to. You need to explicitly set this at runtime. There is no default. You can set multiple email addresses by comma-separating them. [String] [Mandatory]

  .PARAMETER EmailSender
    This sets the sender email address for failure related emails to come from. You need to explicitly set this at runtime. There is no default. [String] [Mandatory]

  .PARAMETER EmailSmtp
    This sets the SMTP server address to be used for sending failure related emails. You need to explicitly set this at runtime. There is no default. [String] [Mandatory]

  .PARAMETER EmailSubjectPrefix
    This sets the first part of the subject line for any failure related emails. You will probably not need to explicitly set this at runtime. By default this is set to "Sentry Daily Update:". [String] [Optional]

  .PARAMETER EnableEmail
    This switch parameter enables emailing error messages. When enabled, the parameters EmailRecipient, EmailSender and EmailSmtp become mandatory. [Switch] [Optional]

  .PARAMETER LogFilePath
    The full file path to a log file to write to. If the file does not already exist it will be automatically created.
    If the specified file is used across multiple runs, the existing file content will be preserved, with new log entries appended to the file. [string]

  .PARAMETER OutputFilename
    This sets the filename of the output file. You will probably not need to explicitly set this at runtime. By default this is set to daily_update.csv. [String] [Optional]

  .PARAMETER OutputFileDirectoryPath
    This sets the destination directory path for the output file. The path can be relative to this script file or absolute. You might need to explicitly set this at runtime. By default this is set to the same directory as the scipt file. [String] [Optional]

  .PARAMETER ProblemRowCount
    This sets the number of minimum number of rows that should be returned before the output is considered "good". This is a sanity check because sometimes the API might consider the download to be complete prematurely. [String] [Optional]
    By default this is set to 20000. [Int]

  .PARAMETER ReportPath
    This sets the Alma Analytics report path. The path should be wrapped in quotes. Spaces in the folder names are expected and the required percent encoding is handled automatically. You need to explicitly set this at runtime. There is no default. [String] [Mandatory]

  .PARAMETER RetryAttempts
    This sets the number of attempts that the script will retry if there is a problem. By default this is set to 5 [Int] [Optional]

  .PARAMETER RowLimit
    This sets the number of rows to gather per request. There will likely be multiple requests per scrip-run. You will probably not need to explicitly set this at runtime. By default this is set to 1000. [Int] [Optional]

  .EXAMPLE
  PS> ./sentry-daily-update.ps1 -EmailSender "do-not-reply@example.org" -EnableEmail -EmailRecipient "john.smith@example.org" -EmailSmtp "smtp.example.org" -ReportPath "/shared/Example University/Reports/Sentry/Sentry user export" -AlmaServerApiKeyIdentifier "AlmaSentryKey"
#>
#Requires -Modules TUN.CredentialManager
#Requires -Version 5.1
[CmdletBinding(DefaultParameterSetName = 'logonly')]
param (
  [Parameter(Mandatory)]
  [ValidateNotNullorEmpty()]
  [string]$AlmaServerApiKeyIdentifier,

  [ValidateSet('ap', 'ca', 'cn', 'eu', 'na')]
  [ValidateNotNullOrEmpty()]
  [string]$ApiRegion = 'eu',

  [ValidateNotNullOrEmpty()]
  [string]$AnalyticsApiBasePath = '/almaws/v1/analytics/reports',

  [ValidateNotNullOrEmpty()]
  [string]$BaseUrl = 'https://api-eu.hosted.exlibrisgroup.com',

  [Parameter(Mandatory, ParameterSetName = 'email')]
  [ValidateNotNullOrEmpty()]
  [string]$EmailRecipient,

  [Parameter(Mandatory, ParameterSetName = 'email')]
  [ValidateNotNullOrEmpty()]
  [string]$EmailSender,

  [Parameter(Mandatory, ParameterSetName = 'email')]
  [ValidateNotNullOrEmpty()]
  [string]$EmailSmtp,

  [Parameter(ParameterSetName = 'email')]
  [ValidateNotNullOrEmpty()]
  [string]$EmailSubjectPrefix = 'Sentry Daily Update:',

  [Parameter(Mandatory, ParameterSetName = 'email')]
  [switch]$EnableEmail,

  [Parameter(ParameterSetName = 'logonly')]
  [Parameter(ParameterSetName = 'email')]
  [ValidateScript({
      Try {
        $resolved = $ExecutionContext.SessionState.Path.GetUnresolvedProviderPathFromPSPath($_)
      }
      Catch {
        Throw 'Syntactically invalid log file path'
      }

      # Block existing folder paths
      If (Test-Path -Path $resolved -PathType Container) {
        Throw 'The path points to an existing folder, but a file path is required.'
      }

      # Check that the parent folder of the path specified exists
      $parent = Split-Path -Path $resolved -Parent
      If (-not (Test-Path -Path $parent -PathType Container)) {
        Throw 'Specified log file parent folder {0} does not exist' -f $parent
      }

      $True
    })]
  [string]$LogFilePath = '.\sentry-daily-update.log',

  [Parameter(ParameterSetName = 'logonly')]
  [Parameter(ParameterSetName = 'email')]
  [ValidateScript({
      # Check that the specified path is syntactically correct
      If (-not $(Test-Path -Path $_ -IsValid)) {
        Throw 'Syntactically invalid output filename'
      }

      $True
    })]
  [string]$OutputFilename = 'daily_update.csv',

  [Parameter(ParameterSetName = 'logonly')]
  [Parameter(ParameterSetName = 'email')]
  [ValidateScript({
      # Check that the specified path is syntactically correct
      If (-not $(Test-Path -Path $_ -IsValid)) {
        Throw 'Syntactically invalid output folder'
      }

      # Check that the folder exists
      If (-not (Test-Path -Path $_ -PathType Container)) {
        Throw 'Specified output folder does not exist'
      }

      $True
    })]
  [string]$OutputFileDirectoryPath = $PSScriptRoot,

  [Parameter(ParameterSetName = 'logonly')]
  [Parameter(ParameterSetName = 'email')]
  [int]$ProblemRowCount = 20000,

  [Parameter(Mandatory, ParameterSetName = 'logonly')]
  [Parameter(Mandatory, ParameterSetName = 'email')]
  [string]$ReportPath,

  [Parameter(ParameterSetName = 'logonly')]
  [Parameter(ParameterSetName = 'email')]
  [int]$RetryAttempts = 5,

  [Parameter(ParameterSetName = 'logonly')]
  [Parameter(ParameterSetName = 'email')]
  [int]$RowLimit = 1000
)

$ErrorActionPreference = 'Stop'

function Send-Email {
  param (
    [string]$EmailSender,
    [string]$EmailRecipient,
    [string]$EmailSmtp,
    [string]$EmailSubject,
    [string]$EmailBody
  )
  Send-MailMessage -From $EmailSender -To $($EmailRecipient.Split(',') | ForEach-Object { $_.Trim() }) -SmtpServer $EmailSmtp -Subject $EmailSubject -Body $EmailBody
}

<#
    .PARAMETER ApiKeyIdentifier
    The identifier for the API key to use when authenticating to the Alma API.
#>
function Get-WindowsCredentialManagerApiKey {
  param (
    [Parameter(Mandatory)]
    [ValidateNotNullorEmpty()]
    [string]$ApiKeyIdentifier
  )

  Try {
    $credentialObject = Get-StoredCredential -Target $ApiKeyIdentifier -AsCredentialObject -IncludeSecurePassword
    $BSTR = [System.Runtime.InteropServices.Marshal]::SecureStringToBSTR($credentialObject.SecurePassword)
    $plainApikey = [System.Runtime.InteropServices.Marshal]::PtrToStringAuto($BSTR)
    Write-Output $plainApikey
  }
  Catch {
    Throw 'Failed to retrieve API key for {0}' -f $ApiKeyIdentifier
  }
  Finally {
    Remove-Variable -Name credentialObject, BSTR, plainApikey -ErrorAction SilentlyContinue
  }
}

function Get-Data {
  param (
    [string]$StrUrl,
    [string]$OutputFileDirectoryPath,
    [string]$OutputFilename,
    [int]$RowLimit,
    [int]$RetryAttempts,
    [string]$EmailSender,
    [string]$EmailRecipient,
    [string]$EmailSmtp,
    [string]$EmailSubjectPrefix,
    [switch]$EnableEmail,
    [string]$LogFilePath,
    [int]$ProblemRowCount
  )
  [int]$retryCount = 0
  [int]$rowCount = 0
  [bool]$complete = $false
  [bool]$success = $true
  $tmpCsvFile = New-TemporaryFile
  if ($EnableEmail) {
    [string]$emailFooter = "`n`nThis email was generated by $(Split-Path -Path $PSCommandPath -Leaf) at $((Get-Date).ToString('dd/MM/yyyy HH:mm:ss')) on $(([System.Net.Dns]::GetHostEntry(($env:computerName))).Hostname)." + `
      "`nFor more information on this email, see https://uoy.atlassian.net/wiki/spaces/ittechdocs/pages/43389721/FMSYS+Library+-+Sentry+turnstile+software"
  }

  $plainApikey = Get-WindowsCredentialManagerApiKey -ApiKeyIdentifier $AlmaServerApiKeyIdentifier

  do {
    try {
      $objRestReq = Invoke-WebRequest -Uri $strUrl -Method Get -Headers @{'Authorization' = "apikey ${plainApikey}" } -TimeoutSec 60 -UseBasicParsing
      if ($objRestReq.StatusCode -eq 200) {
        $restXml = [xml]$objRestReq.Content
        $objRows = $restXml.report.QueryResult.ResultXml.rowset.GetElementsByTagName('Row')
        $objToken = $restXml.report.QueryResult.ResumptionToken
        $objFin = $restXml.report.QueryResult.IsFinished
        $objRemoteError = $restXml.web_service_result.errorsExist
        foreach ($objRow in $objRows) {
          $rowCount++
          $csvLine = '"' + ($objRow.Column0, $objRow.Column1, $objRow.Column2, $objRow.Column3 -join '","') + '"'
          Add-Content -Path $tmpCsvFile -Value $csvLine
        }
        if ($objToken) {
          $strUrl = '{0}{1}?path={2}&limit={3}&token={4}' -f $BaseUrl, $AnalyticsApiBasePath, [System.Uri]::EscapeDataString($ReportPath), $RowLimit, $objToken
        }
        if ($objFin -and $objFin -eq 'true') {
          $complete = $true
          Copy-Item -Path $tmpCsvFile -Destination "$OutputFileDirectoryPath\$OutputFilename"
          Remove-Item -Path $tmpCsvFile
        }
        elseif ($objRemoteError) {
          if ($EnableEmail) {
            Send-Email -EmailSender $EmailSender -EmailRecipient $EmailRecipient -EmailSmtp $EmailSmtp -EmailSubject "${EmailSubjectPrefix} XML response error" -EmailBody "Total rows written: ${rowCount}`nError description: ${objRemoteError}${emailFooter}"
          }
          '{0:yyyy-MM-dd HH:mm:ss}: XML response error - Total rows written: {1} Error description: {2}' -f $(Get-Date), $rowCount, $objRemoteError | Tee-Object -FilePath $LogFilePath -Append
          $success = $false
          break
        }
      }
      else {
        if ($EnableEmail) {
          Send-Email -EmailSender $EmailSender -EmailRecipient $EmailRecipient -EmailSmtp $EmailSmtp -EmailSubject "${EmailSubjectPrefix} Unexpected HTTP response code" -EmailBody "Total rows written: ${rowCount}`nError description: $(objRestReq.StatusText)${emailFooter}"
        }
        '{0:yyyy-MM-dd HH:mm:ss}: Unexpected HTTP response code - Total rows written: {1} Error description: {2}' -f $(Get-Date), $rowCount, $(objRestReq.StatusText) | Tee-Object -FilePath $LogFilePath -Append
        $success = $false
        break
      }
    }
    catch {
      $retryCount++
      if ($retryCount -eq $RetryAttempts) {
        if ($EnableEmail) {
          Send-Email -EmailSender $EmailSender -EmailRecipient $EmailRecipient -EmailSmtp $EmailSmtp -EmailSubject "${EmailSubjectPrefix} $($_.Exception.GetType().Name)" -EmailBody "Total rows written: ${rowCount}`nError description: $($_.Exception.Message)${emailFooter}"
        }
        '{0:yyyy-MM-dd HH:mm:ss}: {1} - Total rows written: {2} Error description: {3}' -f $(Get-Date), $($_.Exception.GetType().Name), $rowCount, $($_.Exception.Message) | Tee-Object -FilePath $LogFilePath -Append
        $success = $false
      }
    }
    # Here we snooze for 2 seconds between resumptions. As advised at https://developers.exlibrisgroup.com/discussions#!/forum/posts/list/63.page
    Start-Sleep -Seconds 2
  } until ($complete -or $retryCount -eq $RetryAttempts)
  Remove-Variable -Name plainApikey -ErrorAction SilentlyContinue
  if ($rowCount -lt $ProblemRowCount -and $success -eq $true) {
    if ($EnableEmail) {
      Send-Email -EmailSender $EmailSender -EmailRecipient $EmailRecipient -EmailSmtp $EmailSmtp -EmailSubject "${EmailSubjectPrefix} Rows written report" -EmailBody "Total rows written: ${rowCount}${emailFooter}"
    }
    '{0:yyyy-MM-dd HH:mm:ss}: Rows written report - Total rows written: {1}' -f $(Get-Date), $rowCount | Tee-Object -FilePath $LogFilePath -Append
  }
}

$resolvedLogFilePath = $ExecutionContext.SessionState.Path.GetUnresolvedProviderPathFromPSPath($Script:LogFilePath)
$resolvedOutputFileDirectoryPath = $ExecutionContext.SessionState.Path.GetUnresolvedProviderPathFromPSPath($Script:OutputFileDirectoryPath)

$BaseUrl = $BaseUrl -replace 'api-[^\\.]+', "api-$ApiRegion"
$strUrl = '{0}{1}?path={2}&limit={3}' -f $BaseUrl, $AnalyticsApiBasePath, [System.Uri]::EscapeDataString($ReportPath), $RowLimit
Get-Data -StrUrl $strUrl `
  -OutputFileDirectoryPath $resolvedOutputFileDirectoryPath `
  -OutputFilename $OutputFilename `
  -RowLimit $RowLimit `
  -RetryAttempts $RetryAttempts `
  -EmailSender $EmailSender `
  -EmailRecipient $EmailRecipient `
  -EmailSmtp $EmailSmtp `
  -EmailSubjectPrefix $EmailSubjectPrefix `
  -EnableEmail $EnableEmail `
  -LogFilePath $resolvedLogFilePath `
  -ProblemRowCount $ProblemRowCount
