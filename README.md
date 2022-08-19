### 	Alma Analytics / Sentry Integration Daily Members Update via REST API

This is a Visual Basic script to download patron data from the Ex Libris's Alma Analytics RESTful API, and write the data to a comma-separated (CSV) file,	with the intention that the output file will then be 'picked up' by Sentry's Daily Update function, according to the Daily Update file path definition and schedule.

#### Prerequisites
- A working Alma Analytics report, with expiry dates formatted according to Sentry's expectations, DD/MM/YYYY
- An application profile must be created on the developers.exlibrisgroup.com site with read permissions for the Analytics API. Create the profile using your Institutional username/password. You will then have an API Key to use with this script.

#### Usage
Open a command prompt and set the API key in the `SENTRY_DU_APIKEY` 'user' environment variable:
```
setx SENTRY_DU_APIKEY xxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxx
```
A number of named arguments exist, and should be appended to the base command as follows:
```
C:\Windows\System32\cscript.exe //nologo "sentry-daily-update.vbs" /named:argument
```
Mandatory arguments:
```
/emailsender:sendinguser@domain.org
/emailrecipient:receivinguser@domain.org
/emailsmtp:smtp.domain.org
/reportpath:%2Fpercent%2Fencoded%2Fpath%2Fsto%2Freport
```

Optional arguments:
```
/emailsubjectprefix:"Some subject prefix: "
/outputfilename:users.csv
/outputfilerelativepath:.
/apiregion:cn ( North America = na, Europe = eu (default), Asia Pacific = ap, Canada = ca, China = cn )
/retryattempts:10
/rowlimit:2000
```
#### Scheduling this script
This script can be run automatically by adding a job to Windows Task Scheduler. Example:
```
<Command>C:\Windows\System32\cscript.exe</Command>
<Arguments>//nologo "c:\path\to\sentry-daily-update.vbs" /emailsender:sendinguser@domain.org /emailrecipient:receivinguser@domain.org /emailsmtp:smtp.domain.org /reportpath:%2Fpercent%2Fencoded%2Fpath%2Fsto%2Freport</Arguments>
<WorkingDirectory>c:\path\to</WorkingDirectory>
```
You'll want to schedule this script to run daily at a time JUST BEFORE the Sentry Daily Update is scheduled to run. The script will take a few minutes to run but to be safe allow at least 10 minutes.

A scheduled task export XML file from that in use at the UoY is included in this repo.

#### Web pages that may be of interest
- https://developers.exlibrisgroup.com/alma/apis
- https://developers.exlibrisgroup.com/blog/How-we-re-building-APIs-at-Ex-Libris
- https://developers.exlibrisgroup.com/alma/apis/analytics
- https://developers.exlibrisgroup.com/blog/Working-with-Analytics-REST-APIs
- https://developers.exlibrisgroup.com/blog/alma_sentry_integration
