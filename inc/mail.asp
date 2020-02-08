<%
If Session("ClientCnnString") <> ""Then
	'SEE IF MAIL IS ON OR OFF
	SQLtoggle = "Select STOPALLEMAIL from " & MUV_Read("SQL_Owner") & ".Settings_Global"
	Set cnntoggle = Server.CreateObject("ADODB.Connection")
	cnntoggle.open (Session("ClientCnnString"))
	Set rstoggle = Server.CreateObject("ADODB.Recordset")
	rstoggle.CursorLocation = 3 
	Set rstoggle = cnntoggle.Execute(SQLtoggle)
	If rstoggle.Eof Then 
		Session("MAILOFF") = 1 ' If eof then set email to off
	Else
		Session("MAILOFF") = rstoggle("STOPALLEMAIL")
	End If
	set rstoggle = Nothing
	cnntoggle.close
	set cnntoggle = Nothing
Else
	Session("MAILOFF") = 0 ' There was no valid ccn string, so assume it is on
End If

Dim emailTo, emailSubject, emailBody, emailBody2, MailSent
MailSent = "False"

Sub SendMail(emailFrom,emailTo,emailSubject,emailBody,emailCategory1,emailCategory2)

	If Instr(ucase(Request.ServerVariables("SERVER_NAME")),"DEV") <> 0 Then emailTo = "rsmith@ocsaccess.com"
	
	dim Mailer
	emailTo = Replace(emailTo, "`", "'")
	Set Mailer = Server.CreateObject("Persits.MailSender")
	Mailer.Host = "mail.vps1-mdsinsight-com.vps.ezhostingserver.com"
	Mailer.Port = 25
    Mailer.Username = "mailsender@dev.mdsinsight.com"
    Mailer.Password = "8WQ&9IKs"
	Mailer.AddAddress emailTo, emailTo
	Mailer.AddBcc "archive@mdsinsight.com","archive@mdsinsight.com"
	Mailer.From = emailFrom
	Mailer.FromName = "MDS Insight"
	Mailer.Subject = emailSubject
	Mailer.IsHTML = True
	Mailer.Body = emailBody

	strErr = ""
	bSuccess = False
	emailStatus = ""

	If Session("MAILOFF") = 0 Then Mailer.Send	' send message
	If Err <> 0 Then ' error occurred
	  strErr = Err.Description
	  emailStatus = strErr 
	else
	  bSuccess = True
	  emailStatus ="Sent Successfully"
	End If
	
	LogEmailToSC_EmailLog emailFrom,emailFromName,emailTo,emailSubject,emailBody,"","","projects@metroplexdata.com",emailStatus,emailCategory1,emailCategory2
	
	
End Sub







Sub SendMailWithCCs(emailFrom,emailTo,emailSubject,emailBody,emailCCs,emailBCCs,emailCategory1,emailCategory2)

	If Instr(ucase(Request.ServerVariables("SERVER_NAME")),"DEV") <> 0 Then emailTo = "rsmith@ocsaccess.com"
	If Instr(ucase(Request.ServerVariables("SERVER_NAME")),"DEV") <> 0 Then emailCCs = "rsmith@ocsaccess.com"
	If Instr(ucase(Request.ServerVariables("SERVER_NAME")),"DEV") <> 0 Then emailBCCs = "rsmith@ocsaccess.com"
	
	dim Mailer
	emailTo = Replace(emailTo, "`", "'")
	Set Mailer = Server.CreateObject("Persits.MailSender")
	Mailer.Host = "mail.vps1-mdsinsight-com.vps.ezhostingserver.com"
	Mailer.Port = 25
    Mailer.Username = "mailsender@dev.mdsinsight.com"
    Mailer.Password = "8WQ&9IKs"
	Mailer.AddAddress emailTo, emailTo
	
	If emailCCs <> "" Then
		CCArray = split(emailCCs,",")
		For i = 0 To Ubound(CCArray)
			Mailer.AddCC CCArray(i),CCArray(i)
		Next
	End If
	
	If emailBCCs <> "" Then
		BCCArray = split(emailBCCs,",")
		For i = 0 To Ubound(BCCArray)
			Mailer.AddBCC BCCArray(i),BCCArray(i)
		Next
	End If
	
	Mailer.AddBCC "archive@mdsinsight.com","archive@mdsinsight.com"
	Mailer.From = emailFrom
	Mailer.FromName = "MDS Insight"
	Mailer.Subject = emailSubject
	Mailer.IsHTML = True
	Mailer.Body = emailBody

	strErr = ""
	bSuccess = False
	emailStatus = ""

	If Session("MAILOFF") = 0 Then Mailer.Send	' send message
	If Err <> 0 Then ' error occurred
	  strErr = Err.Description
	  emailStatus = strErr 
	else
	  bSuccess = True
	  emailStatus ="Sent Successfully"
	End If
	
	LogEmailToSC_EmailLog emailFrom,emailFromName,emailTo,emailSubject,emailBody,"","","projects@metroplexdata.com",emailStatus,emailCategory1,emailCategory2
	
	
End Sub













Sub SendMailWatt(emailFrom,emailTo,emailSubject,emailBody,atfn,emailCategory1,emailCategory2)

	If Instr(ucase(Request.ServerVariables("SERVER_NAME")),"DEV") <> 0 Then emailTo = "rich@ocsaccess.com"
	
	dim Mailer
	emailTo = Replace(emailTo, "`", "'")
	Set Mailer = Server.CreateObject("Persits.MailSender")
	Mailer.Host = "mail.vps1-mdsinsight-com.vps.ezhostingserver.com"
	Mailer.Port = 25
    Mailer.Username = "mailsender@dev.mdsinsight.com"
    Mailer.Password = "8WQ&9IKs"
	Mailer.AddAddress emailTo, emailTo
	Mailer.AddBcc "archive@mdsinsight.com" , "archive@mdsinsight.com"
	Mailer.From = emailFrom
	Mailer.FromName = "MDS Insight"
	Mailer.Subject = emailSubject
	Mailer.IsHTML = True
	Mailer.AddAttachment atfn
	Mailer.Body = emailBody

	strErr = ""
	emailStatus = ""
	bSuccess = False

	If Session("MAILOFF") = 0 Then Mailer.Send	' send message
	
	If Err <> 0 Then ' error occurred
	  strErr = Err.Description
	  emailStatus = strErr 
	else
	  bSuccess = True
	  emailStatus ="Sent Successfully"
	End If
	
	LogEmailToSC_EmailLog emailFrom,emailFromName,emailTo,emailSubject,emailBody,atfn,"","projects@metroplexdata.com",emailStatus,emailCategory1,emailCategory2
	
End Sub





Sub SendMailWMultipleAtt(emailFrom,emailTo,emailSubject,emailBody,attFileNames,emailCategory1,emailCategory2,fromname)
	
	If Instr(ucase(Request.ServerVariables("SERVER_NAME")),"DEV") <> 0 Then emailTo = "rsmith@ocsaccess.com"
	
	dim Mailer
	emailTo = Replace(emailTo, "`", "'")
	Set Mailer = Server.CreateObject("Persits.MailSender")
	Mailer.Host = "mail.vps1-mdsinsight-com.vps.ezhostingserver.com"
	Mailer.Port = 25
    Mailer.Username = "mailsender@dev.mdsinsight.com"
    Mailer.Password = "8WQ&9IKs"
	Mailer.AddAddress emailTo, emailTo
	Mailer.AddBcc "archive@mdsinsight.com","archive@mdsinsight.com"	
	Mailer.From = emailFrom
	Mailer.FromName = fromname
	Mailer.Subject = emailSubject
	Mailer.IsHTML = True
	
	
	attFileNamesArray = split(attFileNames, ",")
	
	For i = 0 To UBound(attFileNamesArray)
		Mailer.AddAttachment attFileNamesArray(i)
	Next
	
	Mailer.Body = emailBody

	strErr = ""
	emailStatus = ""
	bSuccess = False

	If Session("MAILOFF") = 0 Then Mailer.Send	' send message
	
	If Err <> 0 Then ' error occurred
	  strErr = Err.Description
	  emailStatus = strErr 
	else
	  bSuccess = True
	  emailStatus ="Sent Successfully"
	End If
	
	LogEmailToSC_EmailLog emailFrom,emailFromName,emailTo,emailSubject,emailBody,attFileNames,"","projects@metroplexdata.com",emailStatus,emailCategory1,emailCategory2
	
End Sub








Sub SendMailWattCCs(emailFrom,emailTo,emailSubject,emailBody,emailCCs,emailBCCs,atfn,emailCategory1,emailCategory2)

	If Instr(ucase(Request.ServerVariables("SERVER_NAME")),"DEV") <> 0 Then emailTo = "rsmith@ocsaccess.com"
	If Instr(ucase(Request.ServerVariables("SERVER_NAME")),"DEV") <> 0 Then emailCCs = "rsmith@ocsaccess.com"
	If Instr(ucase(Request.ServerVariables("SERVER_NAME")),"DEV") <> 0 Then emailBCCs = "rsmith@ocsaccess.com"
	
	dim Mailer
	emailTo = Replace(emailTo, "`", "'")
	emailTo = "rich@ocsaccess.com"	
	Set Mailer = Server.CreateObject("Persits.MailSender")
	Mailer.Host = "mail.vps1-mdsinsight-com.vps.ezhostingserver.com"
	Mailer.Port = 25
    Mailer.Username = "mailsender@dev.mdsinsight.com"
    Mailer.Password = "8WQ&9IKs"
	Mailer.AddAddress emailTo, emailTo
	
	
	If emailCCs <> "" Then
		CCArray = split(emailCCs,",")
		For i = 0 To Ubound(CCArray)
			Mailer.AddCC CCArray(i),CCArray(i)
		Next
	End If
	
	If emailBCCs <> "" Then
		BCCArray = split(emailBCCs,",")
		For i = 0 To Ubound(BCCArray)
			Mailer.AddBCC BCCArray(i),BCCArray(i)
		Next
	End If
	
	Mailer.AddBCC "archive@mdsinsight.com" , "archive@mdsinsight.com"
	Mailer.From = emailFrom
	Mailer.FromName = "MDS Insight"
	Mailer.Subject = emailSubject
	Mailer.IsHTML = True
	Mailer.AddAttachment atfn
	Mailer.Body = emailBody

	strErr = ""
	emailStatus = ""
	bSuccess = False

	If Session("MAILOFF") = 0 Then Mailer.Send	' send message
	
	If Err <> 0 Then ' error occurred
	  strErr = Err.Description
	  emailStatus = strErr 
	else
	  bSuccess = True
	  emailStatus ="Sent Successfully"
	End If
	
	LogEmailToSC_EmailLog emailFrom,emailFromName,emailTo,emailSubject,emailBody,atfn,"","projects@metroplexdata.com",emailStatus,emailCategory1,emailCategory2
	
End Sub







Sub LogEmailToSC_EmailLog(emailFrom,emailFromName,emailTo,emailSubject,emailBody,emailAttch,emailCCs,emailBCCs,emailStatus,emailCategory1,emailCateogry2)

	'Trim Vars to Avoid Errors
	If Len(emailSubject) > 255 Then emailSubject = Left(emailSubject,255)
	If Len(emailBody) > 8000 Then emailBody = Left(emailBody,8000)
	

	'Creates an entry in SC_EmailLog
	
	SQLRecord_SC_EmailLog = "INSERT INTO SC_EmailLog (RecordCreationDateTime, EmailDate, EmailTime, EmailTo, EmailFrom, EmailFromName, Subject, Body, Attachment,CCs, BCCs,  ASPMailStatus,emailCategory1,emailCategory2) "
	SQLRecord_SC_EmailLog = SQLRecord_SC_EmailLog &  " VALUES ('" & Now() & "','" & GetDateStamp() & "','" & GetTimeStamp() & "','"  & emailTo & "','"  & emailFrom & "', "
	SQLRecord_SC_EmailLog = SQLRecord_SC_EmailLog & "'"  & emailFromName & "','"  & emailSubject & "','"  & Replace(emailbody,"'","''") & "','"  & emailAttch & "','"  & emailCCs & "', "
	SQLRecord_SC_EmailLog = SQLRecord_SC_EmailLog & "'"  & emailBCCs & "','" & emailStatus & "','" & emailCategory1 & "','"  & emailCateogry2 & "')"
	
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


Function isEmailValid(passedemail) 

	isEmailValidresult = 0 ' Assume it is no good
	
	Set Mailer = Server.CreateObject("Persits.MailSender")
	If Mailer.ValidateAddress(passedemail) <> 0 Then isEmailValidresult = Mailer.ValidateAddress(passedemail)
	Set Mailer = Nothing
	
	isEmailValid = isEmailValidresult 
	
	'	0	Valid
	'	1	Too short
	'	2	Too long (greater than 256 chars)
	'	3	No @
	'	4	Nothing before @
	'	5	Characters before @ must be a-z A-Z 0-9 ' _ . - +
	'	6	No dots after @
	'	7	Zero-length subdomain
	'	8	Characters in a subdomain must be a-z A-Z 0-9 -
	'	9	Characters in a top-level subdomain must be a-z A-Z 0-9
	'	10	Top-level subdomain must be at least 2 characters long
	'	11	Name part of address cannot start or end with a dot
	'	12	A subdomain cannot start or end with a dash (-)

End Function
%>