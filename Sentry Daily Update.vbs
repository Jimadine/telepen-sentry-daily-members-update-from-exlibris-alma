'********************************************************************
' 	ALMA ANALYTICS / SENTRY INTEGRATION
' 	DAILY MEMBERS UPDATE VIA REST API
'
'	Created:		15/10/2014
'	Modified:		12/10/2015
'	Version:		1.2
'	Author:			Jim Adamson
'	Description:	A Visual Basic script to download patron data from the Ex Libris's Alma Analytics RESTful API, and write the data to a comma-separated (CSV) file,
'					with the intention that the output file will then be 'picked up' by Sentry's Daily Update function, according to the Daily Update file path definition and schedule.
'
'	== Prerequisites ==
'	* A working Alma Analytics report, with dates formatted according to Sentry's expectations
'	* An application profile must be created on the developers.exlibrisgroup.com site with read permissions for the Analytics API. Create the profile using your Institutional username/password. You will then have an API Key to add to this script.
'
'	== Editing this script ==
'	The values for the CONSTANTS within the lines surrounded by asterisks are the basic changes you will need to make to this script.
'	Beyond that you will probably need to make some adjustments to allow for a differing number of columns in your report; this script assumes 5 columns.
'
'	== Testing this script ==
'	Open a command prompt and enter 'cscript.exe //nologo aa_api_req_v1.2_public.vbs'
'
'	== Scheduling this script ==
'	This script can be run automatically by adding a job to Windows Task Scheduler. Example:
'	<Command>cscript.exe</Command>
'	<Arguments>//nologo "c:\path\to\aa_api_req_v1.2_public.vbs"</Arguments>
'	<WorkingDirectory>c:\path\to</WorkingDirectory>
'
'	== Advice ==
' 	* You'll want to schedule this script to run daily at a time JUST BEFORE the Sentry Daily Update is scheduled to run. The script will take a few minutes to run but to be safe allow at least 10 minutes.
'	* Store this script on a secure drive and set the permissions as appropriate to avoid revealing the embedded API Key to people who might abuse it.
'
'	== Web pages that may be of interest ==
'	https://developers.exlibrisgroup.com/alma/apis
'	https://developers.exlibrisgroup.com/blog/How-we-re-building-APIs-at-Ex-Libris
'	https://developers.exlibrisgroup.com/alma/apis/analytics
'	https://developers.exlibrisgroup.com/blog/Working-with-Analytics-REST-APIs
'	https://developers.exlibrisgroup.com/blog/alma_sentry_integration
'
'********************************************************************

Option Explicit

'************** Edit these CONSTANTS ******************************
'Const AA_REPORT_PATH = "%2Fshared%2FYork University%2FReports%2FSentry%2FSentry%20user%20export%20(Jtest)%20-%20York"
Const AA_REPORT_PATH = "%2Fshared%2FYork%20University%2FReports%2FSentry%2FSentry%20user%20export%20PBD2%20-%20York"
Const API_KEY = "l7xxd1e52027a1324b26b63117efa78e3f79"
Const BASE_URL = "https://api-eu.hosted.exlibrisgroup.com/almaws/v1/analytics/reports?"
Const ERROR_EMAIL_RECIPIENT = "infodir-digital+sentry-daily-update@york.ac.uk"
Const ERROR_EMAIL_SENDER = "do-not-reply@york.ac.uk"
Const ERROR_EMAIL_SMTP = "smtp.york.ac.uk"
Const ERROR_EMAIL_SUBJECT = "Sentry Daily Update: "
Const OUTPUT_FILE_NAME = "daily_update.csv"
Const OUTPUT_FILE_PATH = ".."
Const RETRY_ATTEMPTS = 5
'Const ROW_LIMIT = "25"
Const ROW_LIMIT = "1000"
'******************************************************************

Const FOR_READING = 1, FOR_WRITING = 2, FOR_APPENDING = 8
Dim allDone,child,column0,column1,column2,column3,csvFile,csvFileName,csvLine,emailObj,emailConfig,fail,fin,fso,objShell,remoteError,restReq,restXml,retryCount,row,rowCount,rows,token,url

fail = false
rowCount = 0
retryCount = 0
allDone = false
url = BASE_URL & "path=" & AA_REPORT_PATH & "&limit=" & ROW_LIMIT

Function sendEmail(errorType, rowCount, errorMessage)
	Set emailObj      = CreateObject("CDO.Message")
	emailObj.From     = ERROR_EMAIL_SENDER
	emailObj.To       = ERROR_EMAIL_RECIPIENT
	emailObj.Subject  = ERROR_EMAIL_SUBJECT & errorType
	emailObj.TextBody = "Total rows written: " & rowCount
	If not errorMessage = "" Then
		 emailObj.TextBody = emailObj.TextBody & VbCrLf & "Error description: " & errorMessage
	End If
	Set emailConfig = emailObj.Configuration
	emailConfig.Fields("http://schemas.microsoft.com/cdo/configuration/smtpserver")       = ERROR_EMAIL_SMTP
	emailConfig.Fields("http://schemas.microsoft.com/cdo/configuration/smtpserverport")   = 25
	emailConfig.Fields("http://schemas.microsoft.com/cdo/configuration/sendusing")        = 2
	emailConfig.Fields("http://schemas.microsoft.com/cdo/configuration/smtpauthenticate") = 0
	emailConfig.Fields("http://schemas.microsoft.com/cdo/configuration/smtpusessl")       = false
	emailConfig.Fields("http://schemas.microsoft.com/cdo/configuration/sendusername")     = ERROR_EMAIL_SENDER
	emailConfig.Fields("http://schemas.microsoft.com/cdo/configuration/sendpassword")     = ""
	emailConfig.Fields.Update
	emailObj.Send
	Set emailobj	= Nothing
	Set emailConfig	= Nothing
End Function

Do Until allDone = true OR retryCount = RETRY_ATTEMPTS
	REM My testing revealed that using ye olde Microsoft.XMLHTTP yielded the same XML response data after the second resumption...very weird...so to solve the problem I use MSXML2.ServerXMLHTTP instead
	set restReq = CreateObject("MSXML2.ServerXMLHTTP")
	REM See https://msdn.microsoft.com/en-us/library/windows/desktop/ms760403(v=vs.85).aspx for detail of setTimeouts method. 
	REM resolveTimeout reduced from infinite. For connectTimeout & sendTimeout the defaults are used. receiveTimeout doubled.
	restReq.setTimeouts 10000,60000,30000,60000
	restReq.open "GET", url, false
	restReq.setRequestHeader "Authorization", "apikey " & API_KEY
	On Error Resume Next
	restReq.send
	
	If Err.Number = 0 Then
		If restReq.Status = 200 Then
			set restXml = restReq.responseXML
			set rows = restXml.getElementsByTagName("Row")
			set token = restXml.selectSingleNode("//report/QueryResult/ResumptionToken")
			set fin = restXml.selectSingleNode("//report/QueryResult/IsFinished")
			set remoteError = restXml.selectSingleNode("//web_service_result/errorsExist")
			
			REM First we check <Row>s were returned
			If rows.length > 0 AND remoteError Is Nothing Then
				REM We only need to open the text file once so we test whether the csvFileName has been assigned a value
				If csvFileName = "" Then
					set fso = CreateObject("Scripting.FileSystemObject")
					csvFileName = OUTPUT_FILE_PATH & "\tmp.csv"
					set csvFile = fso.OpenTextFile(csvFileName, FOR_WRITING, true)
				End If
		
				for each row in rows
					rowCount = rowCount + 1
					REM Here we parse each element's children
					for each child in row.ChildNodes
					  select case child.NodeName
						case "Column0"
						  column0 = child.Text
						case "Column1"
						  column1 = child.Text
						case "Column2"
							column2 = child.Text
						case "Column3"
							column3 = child.Text
					  end select
					next
					 REM Compose CSV-line: data is wrapped in double-quotes in case there are commas in the data itself
					 csvLine = """" & column0 & """,""" & column1 & """,""" & column2 & """,""" & column3 & """"
					 csvFile.writeline(csvLine)
					 REM Empty our columns before the next line
					 column0=empty : column1=empty : column2=empty : column3=empty
				next
			
				REM Here we check that resumption token was supplied in the response and rebuild the REST URL to include it
				If not token Is Nothing Then
					url = BASE_URL & "path=" & AA_REPORT_PATH & "&limit=" & ROW_LIMIT & "&token=" & token.text
				End If

				REM Here we perform the check that may end the loop, checking the <IsFinished> XML text and changing the value of allDone. We also close off the CSV file and release it from memory
				If not fin Is Nothing Then
					If fin.text = "true" Then
						csvFile.Close
						fso.CopyFile csvFileName, OUTPUT_FILE_PATH & "\" & OUTPUT_FILE_NAME
						set csvFile = nothing
						allDone = true
					End If
				End If
			
				REM Here we empty the Request and the Response, ready for another iteration
				set restReq = nothing
				set restXml = nothing

			REM Then we check whether the Analytics API responded with an error & if so, send an email
			ElseIf not remoteError Is Nothing Then
				If not csvFile Is Nothing Then
					csvFile.Close
				End If
				'WScript.Echo "An error occurred or no more data"
				sendEmail "XML response error", rowCount, remoteError
				fail = true
				Exit Do
				
			REM Commented out since this generates quite a few unhelpful emails
			'Else 
				'WScript.Echo "API not ready"
				'sendEmail "API not ready", rowCount, restReq.responseText
			End If
			
		Else 
			If not csvFile Is Nothing Then
				csvFile.Close
			End If
			'WScript.Echo "Giving up & sending email"
			sendEmail "Unexpected HTTP response code", rowCount, restReq.StatusText
			fail = true
			REM A retry probably not worth doing so we quit after the first non-200 response code!
			Exit Do
		End If
		
	Else
		retryCount = retryCount + 1
		'WScript.Echo Err.description & "Retrying (attempt number " & retryCount & ")"
		If retryCount = RETRY_ATTEMPTS Then
			If not csvFile Is Nothing Then
				csvFile.Close
			End If
			'WScript.Echo "Giving up & sending email"
			sendEmail Err.source, rowCount, Err.description
			fail = true
		End If
		Err.Clear
	End If
	On Error Goto 0
	REM Here we snooze for 2 seconds between resumptions. As advised at https://developers.exlibrisgroup.com/discussions#!/forum/posts/list/63.page
	WScript.Sleep 2000
Loop

REM Delete temporary file
If IsObject(fso) Then 
	If fso.FileExists(csvFileName) Then
		fso.DeleteFile csvFileName
	End If
End If

REM If the script completed successfully an email is sent. The number of rows returned is not taken into account. This can be uncommented if need be.
'WScript.Echo "Total rows written: " & rowCount
'If fail = false Then
'	sendEmail "Rows written report", rowCount, ""
'End If

REM An email is sent when less than 20000 rows were returned, indicating a problem because typically there are always more than 20000 user records at York.
If rowCount < 20000 Then
	sendEmail "Rows written report", rowCount, ""
End If