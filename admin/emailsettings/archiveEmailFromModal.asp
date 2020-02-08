<!--#include file="../../inc/header.asp"-->
<!--#include file="../../inc/InsightFuncs.asp"-->


<%


'****************************************************************************************
'IF WE RECEIVE A RECORD NUMBER FROM THE QUERYSTRING, THEN THE USER HAS REQUESTED TO 
'ARCHIVE A SINGLE EMAIL FROM THE VIEW FULL EMAIL MODAL WINDOW
'****************************************************************************************
InternalRecordNumber = Request.Form("i")
currentEmailCategory1ViewedID = Request.Form("cat1")
currentEmailCategory2ViewedIDTab = Request.Form("cat2")


'*******************************************************************************
'ARCHIVE SINGLE EMAIL
'*******************************************************************************
If InternalRecordNumber <> "" Then
	
	SQL9 = "UPDATE SC_EmailLog SET Archived=1 WHERE InternalRecordNumber = " & InternalRecordNumber 
	
	Set cnn9 = Server.CreateObject("ADODB.Connection")
	cnn9.open (Session("ClientCnnString"))
	Set rs9 = Server.CreateObject("ADODB.Recordset")
	rs9.CursorLocation = 3 
	Set rs9 = cnn9.Execute(SQL9)


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


	set rs9 = Nothing
	set cnn9  = Nothing

	set rs10 = Nothing
	set cnn10  = Nothing

	Description = "Email with subject, " & Subject & ", sent on " & EmailDate & " at " & EmailTime & " was archived. "
		
	CreateAuditLogEntry "Email Archived From Admin","Email Archived From Admin","Minor",0,Description 

Else

	%>
	Unable to archive email, could not parse querystring for unqiue email identifier.
	<%
	
End If


%><!--#include file="../../inc/footer-main.asp"-->