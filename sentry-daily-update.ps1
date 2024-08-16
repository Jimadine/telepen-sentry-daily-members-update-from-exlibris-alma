#Requires -Version 5.1
<#
  .SYNOPSIS
  A Powershell script to download user data from a pre-prepared Alma Analytics report and convert it to CSV format suitable for importing into Sentry-Isis via its "Daily Update" function

  .DESCRIPTION
  The Sentry-Isis building access management system can import data from a number of different sources as long as the format adheres to the program's requirements. CSV is the format chosen here.
  This script is designed to be run as a scheduled task shortly before Sentry's Daily Update is scheduled to run, ensuring that Sentry imports a fresh set of data from a CSV file, courtesy of this script.

  .PARAMETER ApiKeysDirectoryPath
    This script stores and reads the Alma API key in and from a file. This parameter sets the path to that file. By default this is set to a subdirectory called 'auth' in the same directory as this script. [String]

  .PARAMETER ApiRegion
  Specifies the API region code. Available codes are:
    Asia Pacific = ap  
    Canada = ca
    China = cn
    Europe = eu (default)
    North America = na

    Note that when this paramater is specified, the BaseUrl will be automatically modified to include the ApiRegion specified. 
    By default this is set to eu [String]

  .PARAMETER BaseUrl
    This sets the Alma API base URL. You will probably not need to explicitly set this at runtime. By default this is set to https://api-eu.hosted.exlibrisgroup.com [String] [Optional]

  .PARAMETER BasePath
    This sets the Alma API base Path. You will probably not need to explicitly set this at runtime. By default this is set to /almaws/v1/analytics/reports [String] [Optional]

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
  PS> ./sentry-daily-update.ps1 -EmailSender "do-not-reply@example.org" -EnableEmail -EmailRecipient "john.smith@example.org" -EmailSmtp "smtp.example.org" -ReportPath "/shared/Example University/Reports/Sentry/Sentry user export"
#>
[CmdletBinding(DefaultParameterSetName='logonly')]
param (
    [Parameter(ParameterSetName='logonly')]
    [string]$ApiKeysDirectoryPath = "$PSScriptRoot\auth",

    [Parameter(ParameterSetName='logonly')]
    [ValidateSet("ap", "ca", "cn", "eu", "na")]
    [string]$ApiRegion = "eu",

    [Parameter(ParameterSetName='logonly')]
    [string]$BasePath = '/almaws/v1/analytics/reports',

    [Parameter(ParameterSetName='logonly')]
    [string]$BaseUrl = 'https://api-eu.hosted.exlibrisgroup.com',

    [Parameter(ParameterSetName='email', Mandatory=$true)]
    [string]$EmailRecipient,

    [Parameter(ParameterSetName='email', Mandatory=$true)]
    [string]$EmailSender,

    [Parameter(ParameterSetName='email', Mandatory=$true)]
    [string]$EmailSmtp,

    [Parameter(ParameterSetName='logonly')]
    [string]$EmailSubjectPrefix = "Sentry Daily Update:",

    [Parameter(ParameterSetName='email')]
    [switch]$EnableEmail,

    [Parameter(ParameterSetName='logonly')]
    [string]$LogFilePath = '.\sentry-daily-update.log',

    [Parameter(ParameterSetName='logonly')]
    [string]$OutputFilename = "daily_update.csv",

    [Parameter(ParameterSetName='logonly')]
    [string]$OutputFileDirectoryPath = $PSScriptRoot,

    [Parameter(ParameterSetName='logonly')]
    [int]$ProblemRowCount = 20000,

    [Parameter(ParameterSetName='email', Mandatory=$true)]
    [Parameter(ParameterSetName='logonly', Mandatory=$true)]
    [string]$ReportPath,

    [Parameter(ParameterSetName='logonly')]
    [int]$RetryAttempts = 5,

    [Parameter(ParameterSetName='logonly')]
    [int]$RowLimit = 1000
)

$ErrorActionPreference = "Stop"

If ((Test-Path -Path (Split-Path $LogFilePath)) -ne $true) {
  Write-Warning "$(Split-Path $LogFilePath) path doesn't exist"
  exit 1
}

If ((Test-Path -Path $ApiKeysDirectoryPath) -ne $true) {
    Write-Warning "${ApiKeysDirectoryPath} path doesn't exist - creating missing subfolder"
    $null = New-Item -Type 'directory' -Path "$ApiKeysDirectoryPath" -Force
}

if (-not (Test-Path "$ApiKeysDirectoryPath\apikey.xml")) {
    Write-Warning 'The apikey.xml file doesn''t exist'
    $apikey = Read-Host 'Enter the Ex Libris Alma API key'
    $apikey | Export-Clixml -Path "$ApiKeysDirectoryPath\apikey.xml" -Force
}

try {
    $strApiKey = Import-Clixml -Path "$ApiKeysDirectoryPath\apikey.xml"    
}
catch {
    $($Error[0].Message)
} finally {
    if ([string]::IsNullOrEmpty($strApiKey)) {
        Write-Warning 'API key import issue'
      }      
}

if (-not (Test-Path -Path $OutputFileDirectoryPath)) {
    Write-Host "Output path does not exist"
    exit 1
} 

[string]$BaseUrl = $BaseUrl -replace 'api-[^\.]+', "api-$ApiRegion"
[string]$strUrl = '{0}{1}?path={2}&limit={3}' -f $BaseUrl, $BasePath, [System.Uri]::EscapeDataString($ReportPath), $RowLimit
[int]$retryCount = 0
[int]$rowCount = 0
[bool]$complete = $false
[bool]$success = $true
$tmpCsvFile = New-TemporaryFile

do {
    try {
        $objRestReq = Invoke-WebRequest -Uri $strUrl -Method Get -Headers @{'Authorization' = "apikey ${strApiKey}"} -TimeoutSec 60
        if ($objRestReq.StatusCode -eq 200) {
            $restXml = [xml]$objRestReq.Content
            $objRows = $restXml.report.QueryResult.ResultXml.rowset.GetElementsByTagName("Row")
            $objToken = $restXml.report.QueryResult.ResumptionToken
            $objFin = $restXml.report.QueryResult.IsFinished
            $objRemoteError = $restXml.web_service_result.errorsExist
            foreach ($objRow in $objRows) {
                $rowCount++
                $csvLine = '"' + ($objRow.Column0, $objRow.Column1, $objRow.Column2, $objRow.Column3 -join '","') + '"'
                Add-Content -Path $tmpCsvFile -Value $csvLine
            }
            if ($objToken) {
                $strUrl = '{0}{1}?path={2}&limit={3}&token={4}' -f $BaseUrl, $BasePath, [System.Uri]::EscapeDataString($ReportPath), $RowLimit, $objToken
            }
            if ($objFin -and $objFin -eq "true") {
                $complete = $true
                Copy-Item -Path $tmpCsvFile -Destination "$OutputFileDirectoryPath\$OutputFilename"
                Remove-Item -Path $tmpCsvFile
            } elseif ($objRemoteError) {
                if ($EnableEmail) {
                  Send-MailMessage -From $EmailSender -To $EmailRecipient -SmtpServer $EmailSmtp -Subject "${EmailSubjectPrefix} XML response error" -Body "Total rows written: ${rowCount}`nError description: ${objRemoteError}"
                }
                "{0:yyyy-MM-dd HH:mm:ss}: XML response error - Total rows written: {1} Error description: {2}" -f $(Get-Date), $rowCount, $objRemoteError | Tee-Object -FilePath $LogFilePath -Append
                $success = $false
                break
            }
        } else {
            if ($EnableEmail) {
              Send-MailMessage -From $EmailSender -To $EmailRecipient -SmtpServer $EmailSmtp -Subject "${EmailSubjectPrefix} Unexpected HTTP response code" -Body "Total rows written: ${rowCount}`nError description: $(objRestReq.StatusText)"
            }
            "{0:yyyy-MM-dd HH:mm:ss}: Unexpected HTTP response code - Total rows written: {1} Error description: {2}" -f $(Get-Date), $rowCount, $(objRestReq.StatusText) | Tee-Object -FilePath $LogFilePath -Append
            $success = $false
            break
        }
    } catch {
        $retryCount++
        if ($retryCount -eq $RetryAttempts) {
            if ($EnableEmail) {
              Send-MailMessage -From $EmailSender -To $EmailRecipient -SmtpServer $EmailSmtp -Subject "${EmailSubjectPrefix} $($_.Exception.GetType().Name)" -Body "Total rows written: ${rowCount}`nError description: $($_.Exception.Message)"
            }
            "{0:yyyy-MM-dd HH:mm:ss}: {1} - Total rows written: {2} Error description: {3}" -f $(Get-Date), $($_.Exception.GetType().Name), $rowCount, $($_.Exception.Message) | Tee-Object -FilePath $LogFilePath -Append
            $success = $false
        }
    }
    # Here we snooze for 2 seconds between resumptions. As advised at https://developers.exlibrisgroup.com/discussions#!/forum/posts/list/63.page
    Start-Sleep -Seconds 2
} until ($complete -or $retryCount -eq $RetryAttempts)

# Handle occasional cases where all the records weren't returned, but the script finished normally
if ($rowCount -lt $ProblemRowCount -and $success -eq $true) {
    if ($EnableEmail) {
      Send-MailMessage -From $EmailSender -To $EmailRecipient -SmtpServer $EmailSmtp -Subject "${EmailSubjectPrefix} Rows written report" -Body "Total rows written: ${rowCount}"
    }
    "{0:yyyy-MM-dd HH:mm:ss}: Rows written report - Total rows written: {1}" -f $(Get-Date), $rowCount | Tee-Object -FilePath $LogFilePath -Append
}
