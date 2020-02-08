<!--#include file="../../inc/InSightFuncs.asp"-->
<!--#include file="../../inc/InSightFuncs_Users.asp"-->
<%
InternalRecordIdentifier = Request.QueryString("i")

If InternalRecordIdentifier <> "" Then

	'First look it uo so we can get the alert name
	SQLDisable = "Select * FROM IC_Partners WHERE InternalRecordIdentifier = "& InternalRecordIdentifier 
	
	Set cnnDisable = Server.CreateObject("ADODB.Connection")
	cnnDisable.open (Session("ClientCnnString"))
	Set rsDisable = Server.CreateObject("ADODB.Recordset")
	rsDisable.CursorLocation = 3 
	Set rsDisable = cnnDisable.Execute(SQLDisable)
	
	If not rsDisable.Eof Then 
		partnerCompanyName = rsDisable("partnerCompanyName")
		partnerAPIKey = rsDisable("partnerAPIKey")
	End If
	
	Set rsDisable = Nothing
	cnnDisable.Close
	Set cnnDisable = Nothing

	
	SQLDisable = "UPDATE IC_Partners SET partnerDisabled = 1 WHERE InternalRecordIdentifier = "& InternalRecordIdentifier 
	
	Set cnnDisable = Server.CreateObject("ADODB.Connection")
	cnnDisable.open (Session("ClientCnnString"))
	Set rsDisable = Server.CreateObject("ADODB.Recordset")
	rsDisable.CursorLocation = 3 
	Set rsDisable = cnnDisable.Execute(SQLDisable)
	
	
	Description = "The " & GetTerm("Inventory") & " Partner, " & partnerCompanyName & " with API Key " & partnerCompanyName & ", was disabled by " & GetUserDisplayNameByUserNo(Session("UserNo")) 
	CreateAuditLogEntry GetTerm("Inventory Control") & " Partner Disabled",GetTerm("Inventory Control") & " Partner Disabled","Major",0,Description
	
	set rsDisable = Nothing
	cnnDisable.Close
	set cnnDisable = Nothing
	
End If

Response.Redirect ("main.asp")
%>