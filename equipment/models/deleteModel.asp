<!--#include file="../../inc/InSightFuncs.asp"-->
<!--#include file="../../inc/InSightFuncs_Users.asp"-->
<%
InternalRecordIdentifier = Request.QueryString("i")

If InternalRecordIdentifier <> "" Then

	'First look it uo so we can get the alert name
	SQLDelete = "Select * FROM EQ_Models WHERE InternalRecordIdentifier = "& InternalRecordIdentifier 
	
	Set cnnDelete = Server.CreateObject("ADODB.Connection")
	cnnDelete.open (Session("ClientCnnString"))
	Set rsDelete = Server.CreateObject("ADODB.Recordset")
	rsDelete.CursorLocation = 3 
	Set rsDelete = cnnDelete.Execute(SQLDelete)
	
	If not rsDelete.Eof Then Model = rsDelete("Model")
	
	Set rsDelete = Nothing
	cnnDelete.Close
	Set cnnDelete = Nothing

	
	SQLDelete = "Delete FROM EQ_Models WHERE InternalRecordIdentifier = "& InternalRecordIdentifier 
	
	Set cnnDelete = Server.CreateObject("ADODB.Connection")
	cnnDelete.open (Session("ClientCnnString"))
	Set rsDelete = Server.CreateObject("ADODB.Recordset")
	rsDelete.CursorLocation = 3 
	Set rsDelete = cnnDelete.Execute(SQLDelete)
	
	
	Description = "The " & GetTerm("Equipment") & " model named " & Model & " was deleted by " & GetUserDisplayNameByUserNo(Session("UserNo")) 
	CreateAuditLogEntry GetTerm("Equipment") & " Model Deleted",GetTerm("Equipment") & " Model Deleted","Major",0,Description
	
	set rsDelete = Nothing
	cnnDelete.Close
	set cnnDelete = Nothing
	
End If

Response.Redirect ("main.asp")
%>