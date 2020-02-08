<!--#include file="../../inc/InSightFuncs.asp"-->
<!--#include file="../../inc/InSightFuncs_Users.asp"-->
<%
InternalRecordIdentifier = Request.QueryString("i")

If InternalRecordIdentifier <> "" Then

	'First look it uo so we can get the alert name
	SQLEnable = "Select * FROM IC_Partners WHERE InternalRecordIdentifier = "& InternalRecordIdentifier 
	
	Set cnnEnable = Server.CreateObject("ADODB.Connection")
	cnnEnable.open (Session("ClientCnnString"))
	Set rsEnable = Server.CreateObject("ADODB.Recordset")
	rsEnable.CursorLocation = 3 
	Set rsEnable = cnnEnable.Execute(SQLEnable)
	
	If not rsEnable.Eof Then 
		partnerCompanyName = rsEnable("partnerCompanyName")
		partnerAPIKey = rsEnable("partnerAPIKey")
	End If
	
	Set rsEnable = Nothing
	cnnEnable.Close
	Set cnnEnable = Nothing

	
	SQLEnable = "UPDATE IC_Partners SET partnerDisabled = 0 WHERE InternalRecordIdentifier = "& InternalRecordIdentifier 
	
	Set cnnEnable = Server.CreateObject("ADODB.Connection")
	cnnEnable.open (Session("ClientCnnString"))
	Set rsEnable = Server.CreateObject("ADODB.Recordset")
	rsEnable.CursorLocation = 3 
	Set rsEnable = cnnEnable.Execute(SQLEnable)
	
	
	Description = "The " & GetTerm("Inventory") & " Partner, " & partnerCompanyName & " with API Key " & partnerCompanyName & ", was Enabled by " & GetUserDisplayNameByUserNo(Session("UserNo")) 
	CreateAuditLogEntry GetTerm("Inventory Control") & " Partner Enabled",GetTerm("Inventory Control") & " Partner Enabled","Major",0,Description
	
	set rsEnable = Nothing
	cnnEnable.Close
	set cnnEnable = Nothing
	
End If

Response.Redirect ("main.asp")
%>