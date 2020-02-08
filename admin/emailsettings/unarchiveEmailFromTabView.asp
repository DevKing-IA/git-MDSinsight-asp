<!--#include file="../../inc/header.asp"-->
<!--#include file="../../inc/InsightFuncs.asp"-->


<%
emailReceivedAsArray = false
currentEmailCategory1ViewedIDTab = Request.QueryString("cat1")
currentEmailCategory2ViewedIDTab = Request.QueryString("cat2")


'******************************************************************************************
'IF WE DO NOT RECEIVE A RECORD NUMBER FROM THE QUERYSTRING, THEN THE USER HAS REQUESTED TO 
'UNARCHIVE A SINGLE EMAIL OR MULTIPLE EMAILS FROM THE VIEW ALL EMAILS OUTBOX
'WE NEED TO SPLIT ON A COMMA TO SEE IF WE HAVE AN ARRAY/MULTIPLE EMAILS TO UNARCHIVE 
'******************************************************************************************

InternalRecordNumber = Request.Form("i")
emailReceivedAsArray = false
InternalRecordNumberArray = Split(InternalRecordNumber,",")

If Ubound(InternalRecordNumberArray) = 0 Then
	emailReceivedAsArray = false
	InternalRecordNumber = InternalRecordNumberArray(0)
Else
	emailReceivedAsArray = true
End If


If emailReceivedAsArray = false AND InternalRecordNumber <> "" Then
	
	SQL7 = "UPDATE SC_EmailLog SET Archived=0 WHERE InternalRecordNumber = " & InternalRecordNumber 
	Set cnn7 = Server.CreateObject("ADODB.Connection")
	cnn7.open (Session("ClientCnnString"))
	Set rs7 = Server.CreateObject("ADODB.Recordset")
	rs7.CursorLocation = 3 
	Set rs7 = cnn7.Execute(SQL7)
	set rs7 = Nothing
	set cnn7  = Nothing


	SQL10 = "SELECT * FROM SC_EmailLog WHERE InternalRecordNumber = " & InternalRecordNumber 
	
	Set cnn10 = Server.CreateObject("ADODB.Connection")
	cnn10.open (Session("ClientCnnString"))
	Set rs10 = Server.CreateObject("ADODB.Recordset")
	rs10.CursorLocation = 3 
	Set rs10 = cnn10.Execute(SQL10)
	
	If not rs10.eof then
	
		EmailTo = rs10("EmailTo")
		EmailFrom = rs10("EmailFrom")
		EmailFromName = rs10("EmailFromName")
		EmailDate = rs10("EmailDate")
		EmailTime = FormatDateTime(rs10("EmailTime"),3)
		Subject = rs10("Subject")
		Body = rs10("Body")
		CCs = rs10("CCs")
		BCCs = rs10("BCCs")
		Attachment = rs10("Attachment")

	End If	
	
	set rs10 = Nothing
	set cnn10  = Nothing


	Description = "Email with subject, " & Subject & ", sent on " & EmailDate & " at " & EmailTime & " was un-archived. "

	CreateAuditLogEntry "Email Unarchived From Admin","Email Unarchived From Admin","Minor",0,Description 

	Response.Redirect ("allSentEmails.asp?cat1ID=" & currentEmailCategory1ViewedIDTab & "&tab=" & currentEmailCategory2ViewedIDTab)
	
'*******************************************************************************
'ARCHIVE MULTIPLE EMAILS AT ONCE
'*******************************************************************************
ElseIf emailReceivedAsArray = true Then


	For z = 0 to uBound(InternalRecordNumberArray)
			
		SQL7 = "UPDATE SC_EmailLog SET Archived=0 WHERE InternalRecordNumber = " & InternalRecordNumberArray(z)
		Set cnn7 = Server.CreateObject("ADODB.Connection")
		cnn7.open (Session("ClientCnnString"))
		Set rs7 = Server.CreateObject("ADODB.Recordset")
		rs7.CursorLocation = 3 
		Set rs7 = cnn7.Execute(SQL7)
		set rs7 = Nothing
		set cnn7  = Nothing
	
	
		SQL10 = "SELECT * FROM SC_EmailLog WHERE InternalRecordNumber = " & InternalRecordNumberArray(i)
		
		Set cnn10 = Server.CreateObject("ADODB.Connection")
		cnn10.open (Session("ClientCnnString"))
		Set rs10 = Server.CreateObject("ADODB.Recordset")
		rs10.CursorLocation = 3 
		Set rs10 = cnn10.Execute(SQL10)
		
		If not rs10.eof then
		
			EmailTo = rs10("EmailTo")
			EmailFrom = rs10("EmailFrom")
			EmailFromName = rs10("EmailFromName")
			EmailDate = rs10("EmailDate")
			EmailTime = FormatDateTime(rs10("EmailTime"),3)
			Subject = rs10("Subject")
			Body = rs10("Body")
			CCs = rs10("CCs")
			BCCs = rs10("BCCs")
			Attachment = rs10("Attachment")
	
		End If	
		
		set rs10 = Nothing
		set cnn10  = Nothing
	
	
		Description = "Email with subject, " & Subject & ", sent on " & EmailDate & " at " & EmailTime & " was un-archived. "
	
		CreateAuditLogEntry "Email Unarchived From Admin","Email Unarchived From Admin","Minor",0,Description 
			
	Next
	
	set rs10 = Nothing
	set cnn10  = Nothing	
	
	Response.Redirect ("allSentEmails.asp?cat1ID=" & currentEmailCategory1ViewedIDTab & "&tab=" & currentEmailCategory2ViewedIDTab)

Else

	%><div><br />
	Unable to archive email, could not parse querystring for unqiue email identifier.
	</div>
	<%
	
End If


%><!--#include file="../../inc/footer-main.asp"-->