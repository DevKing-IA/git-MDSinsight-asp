<!--#include file="../../inc/InsightFuncs.asp"-->
<!--#include file="../../inc/InSightFuncs_Users.asp"-->

<%
InternalAlertRecNumber = Request.QueryString("a")
ActiveTab = Request.QueryString("tab")

If InternalAlertRecNumber <> "" Then

	'First look it uo so we can get the alert name
	SQLDelete = "Select * FROM SC_Alerts WHERE InternalAlertRecNumber = "& InternalAlertRecNumber 
	
	Set cnnDelete = Server.CreateObject("ADODB.Connection")
	cnnDelete.open (Session("ClientCnnString"))
	Set rsDelete = Server.CreateObject("ADODB.Recordset")
	rsDelete.CursorLocation = 3 
	Set rsDelete = cnnDelete.Execute(SQLDelete)
	
	If not rsDelete.Eof Then AlertName = rsDelete("AlertName")
	
	Set rsDelete = Nothing
	cnnDelete.Close
	Set cnnDelete = Nothing

	
	SQLDelete = "Delete  FROM SC_Alerts WHERE InternalAlertRecNumber = "& InternalAlertRecNumber 
	
	Set cnnDelete = Server.CreateObject("ADODB.Connection")
	cnnDelete.open (Session("ClientCnnString"))
	Set rsDelete = Server.CreateObject("ADODB.Recordset")
	rsDelete.CursorLocation = 3 
	Set rsDelete = cnnDelete.Execute(SQLDelete)
	
	
	Description = "The alert named " & AlertName & " was deleted by " & GetUserDisplayNameByUserNo(Session("UserNo")) 
	CreateAuditLogEntry "Alert Deleted","Alert Deleted","Minor",0,Description
	
	set rsDelete = Nothing
	cnnDelete.Close
	set cnnDelete = Nothing
	
End If

Response.Redirect ("main.asp#" & ActiveTab)
%>