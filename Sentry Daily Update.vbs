' Created:    15/10/2014
' Modified:   30/04/2021
' Version:    1.0
' Author:     Jim Adamson
' Code formatted with http://www.vbindent.com
Option Explicit
Dim blnAllDone,blnFail,colNamedArguments,intRetryAttempts,intRetryCount,intRowCount,intRowLimit,objChild,objCsvFile,objEmail,objEmailConfig,objFin,objFso,objRemoteError,objRestReq,objRow,objRows,objShell,objToken,restXml,strAnalyticsReportPath,strApiKey,strBaseUrl,strColumn0,strColumn1,strColumn2,strColumn3,strCsvFileName,strCsvLine,strErrorEmailRecipient,strErrorEmailSender,strErrorEmailSmtpServer,strErrorEmailSubjectPrefix,strOutputFilename,strOutputFilepath,strUrl
Set colNamedArguments = WScript.Arguments.Named
Set objShell = WScript.CreateObject("WScript.Shell")
Const FOR_READING = 1, FOR_WRITING = 2, FOR_APPENDING = 8
blnFail = false
intRowCount = 0
intRetryCount = 0
blnAllDone = false

' Mandatory
If objShell.Environment("USER").Item("SENTRY_DU_APIKEY") = "" Then
    WScript.Echo "An API key must be set as a user environment variable with name SENTRY_DU_APIKEY"
    WScript.Quit 1
Else
    strApiKey = objShell.Environment("USER").Item("SENTRY_DU_APIKEY")
End If

If IsEmpty(colNamedArguments.Item("emailsender")) Then
    WScript.Echo "A sender email address must be supplied using /emailsender:sendinguser@domain.org"
    WScript.Quit 1
Else
    strErrorEmailSender = colNamedArguments.Item("emailsender")
End If

If IsEmpty(colNamedArguments.Item("emailrecipient")) Then
    WScript.Echo "A recipient email address must be supplied using /emailrecipient:receivinguser@domain.org"
    WScript.Quit 1
Else 
    strErrorEmailRecipient = colNamedArguments.Item("emailrecipient")
End If

If IsEmpty(colNamedArguments.Item("emailsmtp")) Then
    WScript.Echo "An SMTP server must be supplied using /emailsmtp:smtp.domain.org"
    WScript.Quit 1
Else 
    strErrorEmailSmtpServer = colNamedArguments.Item("emailsmtp")
End If

If IsEmpty(colNamedArguments.Item("reportpath")) Then
    WScript.Echo "An Alma Analytics report path must be supplied using /reportpath:%2Fpercent%2Fencoded%2Fpath%2Fsto%2Freport"
    WScript.Quit 1
Else 
    strAnalyticsReportPath = colNamedArguments.Item("reportpath")
End If

' Optional
If IsEmpty(colNamedArguments.Item("emailsubjectprefix")) Then
    strErrorEmailSubjectPrefix = "Sentry Daily Update: "
Else 
    strErrorEmailSubjectPrefix = colNamedArguments.Item("emailsubjectprefix")
End If

If IsEmpty(colNamedArguments.Item("outputfilename")) Then
    strOutputFilename = "daily_update.csv"
Else 
    strOutputFilename = colNamedArguments.Item("outputfilename")
End If

If IsEmpty(colNamedArguments.Item("outputfilerelativepath")) Then
    strOutputFilepath = "."
Else 
    strOutputFilepath = colNamedArguments.Item("outputfilerelativepath")
End If

If IsEmpty(colNamedArguments.Item("apiregion")) Then
    strBaseUrl = "https://api-eu.hosted.exlibrisgroup.com/almaws/v1/analytics/reports?"
Else 
    strBaseUrl = "https://api-" & colNamedArguments.Exists("apiregion") & ".hosted.exlibrisgroup.com/almaws/v1/analytics/reports?"
End If

If IsEmpty(colNamedArguments.Item("retryattempts")) Then
    intRetryAttempts = 5
Else 
    intRetryAttempts = colNamedArguments.Item("retryattempts")
End If

If IsEmpty(colNamedArguments.Item("rowlimit")) Then
    intRowLimit = 1000
Else 
    intRowLimit = colNamedArguments.Item("rowlimit")
End If

strUrl = strBaseUrl & "path=" & strAnalyticsReportPath & "&limit=" & intRowLimit

Function sendEmail(errorType, intRowCount, errorMessage)
    Set objEmail      = CreateObject("CDO.Message")
    objEmail.From     = strErrorEmailSender
    objEmail.To       = strErrorEmailRecipient
    objEmail.Subject  = strErrorEmailSubjectPrefix & errorType
    objEmail.TextBody = "Total rows written: " & intRowCount
    If not errorMessage = "" Then
        objEmail.TextBody = objEmail.TextBody & VbCrLf & "Error description: " & errorMessage
    End If
    Set objEmailConfig = objEmail.Configuration
    objEmailConfig.Fields("http://schemas.microsoft.com/cdo/configuration/smtpserver")       = strErrorEmailSmtpServer
    objEmailConfig.Fields("http://schemas.microsoft.com/cdo/configuration/smtpserverport")   = 25
    objEmailConfig.Fields("http://schemas.microsoft.com/cdo/configuration/sendusing")        = 2
    objEmailConfig.Fields("http://schemas.microsoft.com/cdo/configuration/smtpauthenticate") = 0
    objEmailConfig.Fields("http://schemas.microsoft.com/cdo/configuration/smtpusessl")       = false
    objEmailConfig.Fields("http://schemas.microsoft.com/cdo/configuration/sendusername")     = strErrorEmailSender
    objEmailConfig.Fields("http://schemas.microsoft.com/cdo/configuration/sendpassword")     = ""
    objEmailConfig.Fields.Update
    objEmail.Send
    Set objEmail    = Nothing
    Set objEmailConfig = Nothing
End Function

Do Until blnAllDone = true OR intRetryCount = intRetryAttempts
    REM My testing revealed that using ye olde Microsoft.XMLHTTP yielded the same XML response data after the second resumption...very weird...so to solve the problem I use MSXML2.ServerXMLHTTP instead
    set objRestReq = CreateObject("MSXML2.ServerXMLHTTP")
    REM See https://msdn.microsoft.com/en-us/library/windows/desktop/ms760403(v=vs.85).aspx for detail of setTimeouts method. 
    REM resolveTimeout reduced from infinite. For connectTimeout & sendTimeout the defaults are used. receiveTimeout doubled.
    objRestReq.setTimeouts 10000,60000,30000,60000
    objRestReq.open "GET", strUrl, false
    objRestReq.setRequestHeader "Authorization", "apikey " & strApiKey
    On Error Resume Next
    objRestReq.send
    
    If Err.Number = 0 Then
        If objRestReq.Status = 200 Then
            set restXml = objRestReq.responseXML
            set objRows = restXml.getElementsByTagName("Row")
            set objToken = restXml.selectSingleNode("//report/QueryResult/ResumptionToken")
            set objFin = restXml.selectSingleNode("//report/QueryResult/IsFinished")
            set objRemoteError = restXml.selectSingleNode("//web_service_result/errorsExist")
            
            REM First we check <Row>s were returned
            If objRows.length > 0 AND objRemoteError Is Nothing Then
                REM We only need to open the text file once so we test whether the strCsvFileName has been assigned a value
                If strCsvFileName = "" Then
                    set objFso = CreateObject("Scripting.FileSystemObject")
                    strCsvFileName = strOutputFilepath & "\tmp.csv"
                    set objCsvFile = objFso.OpenTextFile(strCsvFileName, FOR_WRITING, true)
                End If
                
                for each objRow in objRows
                    intRowCount = intRowCount + 1
                    REM Here we parse each element's children
                    for each objChild in objRow.ChildNodes
                        select case objChild.NodeName
                        case "Column0"
                            strColumn0 = objChild.Text
                        case "Column1"
                            strColumn1 = objChild.Text
                        case "Column2"
                            strColumn2 = objChild.Text
                        case "Column3"
                            strColumn3 = objChild.Text
                        end select
                    next
                    REM Compose CSV-line: data is wrapped in double-quotes in case there are commas in the data itself
                    strCsvLine = """" & strColumn0 & """,""" & strColumn1 & """,""" & strColumn2 & """,""" & strColumn3 & """"
                    objCsvFile.writeline(strCsvLine)
                    REM Empty our columns before the next line
                    strColumn0=empty : strColumn1=empty : strColumn2=empty : strColumn3=empty
                next
                
                REM Here we check that resumption token was supplied in the response and rebuild the REST URL to include it
                If not objToken Is Nothing Then
                    strUrl = strBaseUrl & "path=" & strAnalyticsReportPath & "&limit=" & intRowLimit & "&token=" & objToken.text
                End If
                
                REM Here we perform the check that may end the loop, checking the <IsFinished> XML text and changing the value of blnAllDone. We also close off the CSV file and release it from memory
                If not objFin Is Nothing Then
                    If objFin.text = "true" Then
                        objCsvFile.Close
                        objFso.CopyFile strCsvFileName, strOutputFilepath & "\" & strOutputFilename
                        set objCsvFile = nothing
                        blnAllDone = true
                    End If
                End If
                
                REM Here we empty the Request and the Response, ready for another iteration
                set objRestReq = nothing
                set restXml = nothing
                
                REM Then we check whether the Analytics API responded with an error & if so, send an email
            ElseIf not objRemoteError Is Nothing Then
                If not objCsvFile Is Nothing Then
                    objCsvFile.Close
                End If
'WScript.Echo "An error occurred or no more data"
                sendEmail "XML response error", intRowCount, objRemoteError
                blnFail = true
                Exit Do
                
                REM Commented out since this generates quite a few unhelpful emails
'Else 
'WScript.Echo "API not ready"
'sendEmail "API not ready", intRowCount, objRestReq.responseText
            End If
            
        Else 
            If not objCsvFile Is Nothing Then
                objCsvFile.Close
            End If
'WScript.Echo "Giving up & sending email"
            sendEmail "Unexpected HTTP response code", intRowCount, objRestReq.StatusText
            blnFail = true
            REM A retry probably not worth doing so we quit after the first non-200 response code!
            Exit Do
        End If
        
    Else
        intRetryCount = intRetryCount + 1
'WScript.Echo Err.description & "Retrying (attempt number " & intRetryCount & ")"
        If intRetryCount = intRetryAttempts Then
            If not objCsvFile Is Nothing Then
                objCsvFile.Close
            End If
'WScript.Echo "Giving up & sending email"
            sendEmail Err.source, intRowCount, Err.description
            blnFail = true
        End If
        Err.Clear
    End If
    On Error Goto 0
    REM Here we snooze for 2 seconds between resumptions. As advised at https://developers.exlibrisgroup.com/discussions#!/forum/posts/list/63.page
    WScript.Sleep 2000
Loop

REM Delete temporary file
If IsObject(objFso) Then 
    If objFso.FileExists(strCsvFileName) Then
        objFso.DeleteFile strCsvFileName
    End If
End If

REM If the script completed successfully an email is sent. The number of rows returned is not taken into account. This can be uncommented if need be.
'WScript.Echo "Total rows written: " & intRowCount
'If blnFail = false Then
'   sendEmail "Rows written report", intRowCount, ""
'End If

REM An email is sent when less than 20000 rows were returned, indicating a problem because typically there are always more than 20000 user records at York.
If intRowCount < 20000 And blnFail = false Then
    sendEmail "Rows written report", intRowCount, ""
End If
