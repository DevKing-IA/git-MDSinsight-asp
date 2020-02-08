<!--#include file="../../inc/InSightFuncs.asp"-->
<!--#include file="../../inc/InSightFuncs_Users.asp"-->
<%
InternalRecordIdentifier = Request.QueryString("i")

If InternalRecordIdentifier <> "" Then

	'First look it uo so we can get the alert name
	SQLDelete = "SELECT * FROM AR_Terms WHERE InternalRecordIdentifier = "& InternalRecordIdentifier 
	
	Set cnnDelete = Server.CreateObject("ADODB.Connection")
	cnnDelete.open (Session("ClientCnnString"))
	Set rsDelete = Server.CreateObject("ADODB.Recordset")
	rsDelete.CursorLocation = 3 
	Set rsDelete = cnnDelete.Execute(SQLDelete)
	
	If not rsDelete.Eof Then TermDescription = rsDelete("Description")
	
	Set rsDelete = Nothing
	cnnDelete.Close
	Set cnnDelete = Nothing

	
	SQLDelete = "Delete FROM AR_Terms WHERE InternalRecordIdentifier = "& InternalRecordIdentifier 
	
	Set cnnDelete = Server.CreateObject("ADODB.Connection")
	cnnDelete.open (Session("ClientCnnString"))
	Set rsDelete = Server.CreateObject("ADODB.Recordset")
	rsDelete.CursorLocation = 3 
	Set rsDelete = cnnDelete.Execute(SQLDelete)
	
	
	Description = "The Accounts Receivable module term " & TermDescription & " was deleted by " & GetUserDisplayNameByUserNo(Session("UserNo")) 
	CreateAuditLogEntry "Accounts Receivable module term deleted","Accounts Receivable module term deleted","Major",0,Description
	
	set rsDelete = Nothing
	cnnDelete.Close
	set cnnDelete = Nothing
	
End If

Response.Redirect ("main.asp")
%>