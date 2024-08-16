### 	Alma Analytics / Sentry Integration Daily Members Update via REST API

This is a Powershell script to download patron data from the Ex Libris's Alma Analytics RESTful API, and write the data to a comma-separated (CSV) file,	with the intention that the output file will then be imported by Sentry's Daily Update function, according to the Daily Update file path definition and schedule.

#### Prerequisites
- A working Alma Analytics report, with expiry dates formatted according to Sentry's expectations, DD/MM/YYYY
- An application profile must be created on the developers.exlibrisgroup.com site with read permissions for the Analytics API. Create the profile using your Institutional username/password. You will then have an API Key to use with this script.

#### Manual usage
Although the script is designed to be run as a scheduled task, it might be helpful to run it manually to get a feel for how it works. The first time you run the script, an API key will be requested, so you will want to put this in your paste buffer just before running it.

To begin, open a Powershell window, and `cd` to the directory containing `sentry-daily-update.ps1`. A number of named parameters exist, and should be appended to the base command as follows:
```
./sentry-daily-update.ps1 -NamedParam Value
```
Mandatory parameters with example values:
```
-EmailRecipient john.smith@example.org
-EmailSender do-not-reply@example.org
-EmailSmtp smtp.example.org
-ReportPath "/shared/Example University/Reports/Sentry/Sentry user export"
```

Optional parameters with example values:
```
-ApiKeysDirectoryPath
-ApiRegion cn
-BasePath /almaws/v1/analytics/reports
-BaseUrl https://api-eu.hosted.exlibrisgroup.com
-EmailSubjectPrefix "Some subject prefix:"
-EnableEmail
-OutputFileDirectoryPath .
-OutputFilename users.csv
-ProblemRowCount 30000
-RetryAttempts 10
-RowLimit 2000
```

For a full description of the parameters, see the parameter descriptions at the top of the script file.

#### Scheduling this script
This script can be run automatically by creating a new task in Windows Task Scheduler. The most important pieces of information required for this are the `Command` and `Arguments`. Here's an example:

Command: `C:\Windows\System32\WindowsPowerShell\v1.0\powershell.exe`  
Arguments: `-NoLogo -NoProfile -ExecutionPolicy ByPass -File "C:\path\to\sentry-daily-update.ps1" -EnableEmail -EmailSender "do-not-reply@example.org" -EmailRecipient "john.smith@example.org" -EmailSmtp "smtp.example.org" -ReportPath "/shared/Example University/Reports/Sentry/Sentry user export"`  
Start in: `c:\path\to`

You'll want to schedule this script to run daily at a time JUST BEFORE the Sentry Daily Update is scheduled to run. The script will take a few minutes to run but to be safe allow at least 10 minutes. Note that it's recommended to use double-quotes not single-quotes around parameter values in Windows scheduled tasks, [otherwise odd stuff can happen](https://stackoverflow.com/questions/44594179/pass-powershell-parameters-within-task-scheduler#comment110388918_44594978).

If you want to run the script without showing a Powershell window, there are two options:
* Add `-WindowStyle Hidden` to the first set of parameters (before the `-File` parameter). This is not ideal because the user will still see a flash of window before the parameter is processed.
* Use a VBScript wrapper script to launch Powershell, so your scheduled task Command/Arguments becomes:

Command: `C:\Windows\System32\WScript.exe`  
Arguments: `//B silent.vbs C:\Windows\System32\WindowsPowerShell\v1.0\powershell.exe -NoLogo -NoProfile -ExecutionPolicy ByPass -File "C:\path\to\sentry-daily-update.ps1" -EnableEmail -EmailSender "do-not-reply@example.org" -EmailRecipient "john.smith@example.org" -EmailSmtp "smtp.example.org" -ReportPath "/shared/Example University/Reports/Sentry/Sentry user export"`  
Start in: `c:\path\to`

The `silent.vbs` wrapper script needs to be in the `Start in` directory:

```
On Error Resume Next

ReDim args(WScript.Arguments.Count-1)

For i = 0 To WScript.Arguments.Count-1
    If InStr(WScript.Arguments(i), " ") > 0 Then
        args(i) = Chr(34) & WScript.Arguments(i) & Chr(34)
    Else
        args(i) = WScript.Arguments(i)
        End If

Next

CreateObject("WScript.Shell").Run Join(args, " "), 0, False
```
This does the job, completely eliminating the window, but again, not perfect for those who like minimalism.

#### Web pages that may be of interest
- https://developers.exlibrisgroup.com/alma/apis
- https://developers.exlibrisgroup.com/blog/How-we-re-building-APIs-at-Ex-Libris
- https://developers.exlibrisgroup.com/alma/apis/analytics
- https://developers.exlibrisgroup.com/blog/Working-with-Analytics-REST-APIs
- https://developers.exlibrisgroup.com/blog/alma_sentry_integration

### Local note for UoY Sysadmins
The specific parameters that we use are in a scheduled task export XML file located in the private repo https://github.com/digital-york/telepen-sentry-scripts-and-conf → `scheduled_tasks_export\Sentry Daily Update.xml` file
