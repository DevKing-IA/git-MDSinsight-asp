<!--#include file="../inc/subsandfuncs.asp"-->
<!--#include file="../inc/InsightFuncs.asp"-->
<%
'baseURL should alwats have a trailing /slash, just in case, handle either way
If right(baseURL,1)="/" Then maildomain = Left(right(baseURL,len(baseURL)-7),len(right(baseURL,len(baseURL)-7))-1) Else maildomain = right(baseURL,len(baseURL)-7)


CstNum = "1003"
rtdte = "3/15/2016"
Reason  = "121"
If Reason = "" Then Reason = 0

Response.Write("CstNum :" & CstNum & "<br>")
Response.Write("rtdte :" & rtdte & "<br>")
Response.Write("Reason  :" & Reason  & "<br>")
Response.Write("Session(Userno)  :" & Session("UserNo")  & "<br>")
Response.Write("Session(ClientCnnString)  :" & Session("ClientCnnString")  & "<br>")
Session("ClientCnnString") = "Driver={SQL Server};Server=66.201.99.15;Database=CorpCoffeedev;Uid=" & MUV_Read("SQL_Owner") & ";Pwd=5um47AS;"
Session("UserNo") = "2"

If  Month(rtdte) < 10 Then
	dteHold = "0" & Month(rtdte) & "/"
Else
	dteHold = Month(rtdte) & "/"
End IF

If  Day(rtdte) < 10 Then
	dteHold = dteHold  & "0" & Day(rtdte) & "/"
Else
	dteHold = dteHold  & Day(rtdte) & "/"
End IF
dteHold = dteHold & Right(rtdte,2)
rtdte = dteHold

Set cnnrteDte = Server.CreateObject("ADODB.Connection")
cnnrteDte.open (Session("ClientCnnString"))

Set rsrteDte = Server.CreateObject("ADODB.Recordset")
rsrteDte.CursorLocation = 3 

SQL = "Select ReturnDate from AR_Customer WHERE CustNum= " & CstNum
Set rsrteDte = cnnrteDte.Execute(SQL)
If not rsrteDte.Eof Then ORDte = rsrteDte("ReturnDate")

SQL = "UPDATE AR_Customer Set ReturnDate = '" & rtdte & "' WHERE CustNum= " & CstNum
Set rsrteDte = cnnrteDte.Execute(SQL)

set rsrteDte = Nothing
cnnrteDte.Close
Set cnnrteDte = Nothing

Description = ""
Description = Description & "The return date was manually changed for account # "  & CstNum 
Description = Description & "     The new return date is: "  & rtdte
Description = Description & "     The original return date was: "  & ORDte
 
CreateAuditLogEntry "Return Date Changed","Return Date Changed","Minor",0,Description

'Post to MDS goes here
data = "<DATASTREAM>"
data = data & "<IDENTITY>Pm8316wyc011</IDENTITY>"
data = data & "<MODE>TEST</MODE>"
data = data & "<RECORD_TYPE>UPDATE_CUSTOMER</RECORD_TYPE>"
data = data & "<RECORD_SUBTYPE>RETURN_DATE</RECORD_SUBTYPE>"
data = data & "<CLIENT_ID>CCS</CLIENT_ID>"
data = data & "<SERNO>1071d</SERNO>"
data = data & "<SUBMISSION_SOURCE>MDS Insight</SUBMISSION_SOURCE>"
data = data & "<ACCOUNT_NUM>" & CstNum & "</ACCOUNT_NUM>"
data = data & "<FIELD_DATA>" & rtdte & "</FIELD_DATA>"
data = data & "<FIELD_DATA1>" & Reason & "</FIELD_DATA1>"
data = data & "</DATASTREAM>"

Description = "Post to " & "http://23.115.72.137:3291/ocsmds/ocsapi"
CreateINSIGHTAuditLogEntry sURL,Description,"TEST"
Description = "data:" & data 
CreateINSIGHTAuditLogEntry sURL,Description,"TEST"

Set httpRequest = Server.CreateObject("MSXML2.ServerXMLHTTP")
httpRequest.Open "POST", "http://23.115.72.137:3291/ocsmds/ocsapi", False
httpRequest.SetRequestHeader "Content-Type", "application/x-www-form-urlencoded"
httpRequest.Send data
	
IF httpRequest.status = 200 THEN 
	Description = "httpRequest.responseText:" & httpRequest.responseText
	CreateINSIGHTAuditLogEntry sURL,Description,"TEST",SERNO,SERNO,"PostReturnDateToMDS"
	If Instr(httpRequest.responseText,"success") = 0 Then
		Response.Write("POST RESPONSE:------X" & httpRequest.responseText & "<---------------<br>")
		emailBody = httpRequest.responseText
		emailBody = emailBody & "    PAGE: " & Request.ServerVariables("SERVER_NAME") & Request.ServerVariables("URL")
		emailBody = emailBody & "    POSTED DATA:" & data
		SendMail "mailsender@" & maildomain ,"projects@metroplexdata.com","BAD POST",emailBody
	Else
		Response.Write("<br><b>SUCCESS!</b><br>")
	End If
ELSE
	'In here it must email us if there are problems
	Description = "httpRequest.status:" & httpRequest.status
	CreateINSIGHTAuditLogEntry sURL,Description,"TEST",SERNO,SERNO,"PostReturnDateToMDS"
	emailBody = "httpRequest.responseText:" & httpRequest.responseText
	emailBody = emailBody & "    Unable to communicate with Tomcat at http://23.115.72.137:3291/ocsmds/servicememo" 
	emailBody = emailBody & "    POSTED DATA:" & data
	emailBody = emailBody & "    httpRequest.status:" & httpRequest.status
	SendMail "mailsender@" & maildomain ,"projects@metroplexdata.com","POST ERROR",emailBody
END IF

Response.write("Finished<br>")	
Response.write("postResponse :" & postResponse  & "<br>")	


Sub SendMail(emailFrom,emailTo,emailSubject,emailBody)

	dim Mailer
	emailTo = Replace(emailTo, "`", "'")
	Set Mailer = Server.CreateObject("Persits.MailSender")
	Mailer.Host = "mail.vps1-mdsinsight-com.vps.ezhostingserver.com"
	Mailer.Port = 25
    Mailer.Username = "mailsender@dev.mdsinsight.com"
    Mailer.Password = "8WQ&9IKs"
	Mailer.AddAddress emailTo, emailTo
	Mailer.AddBcc "projects@metroplexdata.com" , "projects@metroplexdata.com"
	Mailer.From = emailFrom
	Mailer.FromName = "MDS Insight"
	Mailer.Subject = emailSubject
	Mailer.IsHTML = True
	Mailer.Body = emailBody

	strErr = ""
	bSuccess = False
	emailStatus = ""
	On Error Resume Next ' catch errors
	If Session("MAILOFF") = 0 Then Mailer.Send	' send message
	If Err <> 0 Then ' error occurred
	  strErr = Err.Description
	  emailStatus = strErr 
	else
	  bSuccess = True
	  emailStatus ="Sent Successfully"
	End If
	
End Sub


Sub LogEmailToSC_EmailLog(emailFrom,emailFromName,emailTo,emailSubject,emailBody,emailAttch,"","projects@metroplexdata.com",emailStatus,"PostReturnDateToMDS","Post Return Date")

	'Creates an entry in SC_EmailLog
	
	SQLRecord_SC_EmailLog = "INSERT INTO SC_EmailLog (RecordCreationDateTime, EmailDate, EmailTime, EmailTo, EmailFrom, EmailFromName, Subject, Body, CCs, BCCs, Attachment, ASPMailStatus) "
	SQLRecord_SC_EmailLog = SQLRecord_SC_EmailLog &  " VALUES (" & Now() & "," & GetDateStamp() & "," & GetTimeStamp() & ",'"  & emailTo & "','"  & emailFrom & "', "
	SQLRecord_SC_EmailLog = SQLRecord_SC_EmailLog & "'"  & emailFromName & "','"  & emailTo& "','"  & emailSubject & "','"  & emailBody & "','"  & emailAttch & "','"  & emailCCs & "', "
	SQLRecord_SC_EmailLog = SQLRecord_SC_EmailLog & "'"  & emailBCCs & "','"  & emailCategory1 & "','"  & emailCateogry2 & "')"
	
	Set cnnRecord_SC_EmailLog = Server.CreateObject("ADODB.Connection")
	cnnRecord_SC_EmailLog.open (Session("ClientCnnString"))

	Set rsRecord_SC_EmailLog = Server.CreateObject("ADODB.Recordset")
	rsRecord_SC_EmailLog.CursorLocation = 3 
	Set rsRecord_SC_EmailLog = cnnRecord_SC_EmailLog.Execute(SQLRecord_SC_EmailLog)
	set rsRecord_SC_EmailLog = Nothing

End Sub

Function LeadingZeros(ByVal Number, ByVal Places)
  Dim Zeros
  Zeros = String(CInt(Places), "0")
  LeadingZeros = Right(Zeros & CStr(Number), Places)
End Function

Function GetTimeStamp
  Dim CurrTime
  CurrTime = Now()
  GetTimeStamp =LeadingZeros(Hour(CurrTime),   2) & ":" & LeadingZeros(Minute(CurrTime), 2) & ":" & LeadingZeros(Second(CurrTime), 2)
End Function

Function GetDateStamp
  Dim CurrDate
  CurrDate = Now()
  GetDateStamp = LeadingZeros(Month(CurrDate),2) & "/" & LeadingZeros(Day(CurrDate),2)  & "/" & LeadingZeros(Year(CurrDate),4)    
End Function


Sub CreateAuditLogEntry(passedElementOrEventName,passedElementOrEventNav,passedMajorMinor,passedSettingChange,passedDescription) 

	'Creates an entry in SC_AuditLog
	
	passedDescription= replace(passedDescription,"'","")
	
	Dim UserIPAddress
	
	UserIPAddress = Request.ServerVariables("HTTP_X_FORWARDED_FOR")
	If UserIPAddress = "" Then
		UserIPAddress = Request.ServerVariables("REMOTE_ADDR")
	End If

	NameToLog = ""
	NameToLog = MUV_Read("DisplayName")
	If passedElementOrEventName = "Service Realtime Alert Sent" or passedElementOrEventName = "Service Realtime Escalation Alert Sent" Then NameToLog = "" ' So it will use SYSTEM
	If NameToLog = "" Then NameToLog = "System"
	
	SQL = "INSERT INTO SC_AuditLog (AuditElementOrEventNav,AuditElementOrEventName,AuditUserEmail, "
	SQL = SQL & "AuditDescription,AuditSettingChange,AuditIPAddress,AuditUserDisplayName,AuditMajorMinor)"
	SQL = SQL &  " VALUES ('" & passedElementOrEventNav & "'"
	SQL = SQL & ",'"  & passedElementOrEventName & "'"
	SQL = SQL & ",'"  & Session("userEmail") & "'"
	SQL = SQL & ",'"  & passedDescription & "'"		
	SQL = SQL & ","  & passedSettingChange
	SQL = SQL & ",'"  & UserIPAddress & "'"
	SQL = SQL & ",'"  & NameToLog & "'"
	SQL = SQL & ",'"  & passedMajorMinor & "')"
	
	'response.write(SQL)
	
	Set cnn8 = Server.CreateObject("ADODB.Connection")
	cnn8.open (Session("ClientCnnString"))

	Set rs8 = Server.CreateObject("ADODB.Recordset")
	rs8.CursorLocation = 3 
	Set rs8 = cnn8.Execute(SQL)
	set rs8 = Nothing
	
End Sub 

Sub CreateINSIGHTAuditLogEntry(passedIdentity,passedLogEntry,passedMode) 

	on error resume next
	'Creates an entry in API_AuditLog
	
	passedLogEntry= replace(passedLogEntry,"'","")

	SQL = "INSERT INTO API_AuditLog([Identity],LogEntry,Mode)"
	SQL = SQL &  " VALUES ('" & passedIdentity & "'"
	SQL = SQL & ",'"  & passedLogEntry & "'"
	SQL = SQL & ",'"  & passedMode & "')"
	
	'response.write(SQL)
	
	Set cnn8 = Server.CreateObject("ADODB.Connection")
	cnn8.open ("Driver={SQL Server};Server=66.201.99.15;Database=_BIInsight;Uid=biinsight;Pwd=Z32#kje4217;")

	Set rs8 = Server.CreateObject("ADODB.Recordset")
	rs8.CursorLocation = 3 
		
	Set rs8 = cnn8.Execute(SQL)

	set rs8 = Nothing
	
	On error goto 0
	
End Sub 

%>
