<!--#include file="../../inc/InSightFuncs.asp"-->
<!--#include file="../../inc/InSightFuncs_Users.asp"-->
<%
InternalRecordIdentifier = Request.QueryString("i")

If InternalRecordIdentifier <> "" Then

	'First look it uo so we can get the alert name
	SQLDelete = "SELECT * FROM FS_Parts WHERE InternalRecordIdentifier = "& InternalRecordIdentifier 
	
	Set cnnDelete = Server.CreateObject("ADODB.Connection")
	cnnDelete.open (Session("ClientCnnString"))
	Set rsDelete = Server.CreateObject("ADODB.Recordset")
	rsDelete.CursorLocation = 3 
	Set rsDelete = cnnDelete.Execute(SQLDelete)
	
	If not rsDelete.Eof Then PartNumber = rsDelete("PartNumber")
	
	Set rsDelete = Nothing
	cnnDelete.Close
	Set cnnDelete = Nothing

	
	SQLDelete = "Delete FROM FS_Parts WHERE InternalRecordIdentifier = "& InternalRecordIdentifier 
	
	Set cnnDelete = Server.CreateObject("ADODB.Connection")
	cnnDelete.open (Session("ClientCnnString"))
	Set rsDelete = Server.CreateObject("ADODB.Recordset")
	rsDelete.CursorLocation = 3 
	Set rsDelete = cnnDelete.Execute(SQLDelete)
	
	
	Description = "The service module part number " & PartNumber & " was deleted by " & GetUserDisplayNameByUserNo(Session("UserNo")) 
	CreateAuditLogEntry "Service module part number deleted","Service module part number deleted","Major",0,Description
	
	set rsDelete = Nothing
	cnnDelete.Close
	set cnnDelete = Nothing
	
End If

Response.Redirect ("main.asp")
%>