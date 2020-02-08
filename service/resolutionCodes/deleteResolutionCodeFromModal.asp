<!--#include file="../../inc/InsightFuncs_Prospecting.asp"-->
<!--#include file="../../inc/InsightFuncs.asp"-->
<!--#include file="../../inc/InsightFuncs_Users.asp"-->

<%
ResolutionCodeToReplace = Request.Form("txtResolutionCodeToReplace")
ResolutionCodeReplaceWith = Request.Form("seldeleteResolutionCodeFromModal")

If ResolutionCodeToReplace <> "" AND ResolutionCodeReplaceWith <> "" Then
	
	'Now replace all records with the new resolution code
	Set cnnDelete = Server.CreateObject("ADODB.Connection")
	cnnDelete.open (Session("ClientCnnString"))
	Set rsDelete = Server.CreateObject("ADODB.Recordset")

	SQLDelete = "UPDATE FS_ServiceMemos Set ResolutionCode = " & ResolutionCodeReplaceWith & " WHERE ResolutionCode = " & ResolutionCodeToReplace 
	Set rsDelete = cnnDelete.Execute(SQLDelete)
	
	'Now Do the deletion
	
	SQLDelete = "Delete FROM FS_ResolutionCodes WHERE InternalRecordIdentifier = "& ResolutionCodeToReplace 
	rsDelete.CursorLocation = 3 
	Set rsDelete = cnnDelete.Execute(SQLDelete)
	
	Description = "The service module resolution code " & ResolutionCodeToReplace & " was deleted by " & GetUserDisplayNameByUserNo(Session("UserNo")) 
	CreateAuditLogEntry "Service module resolution code deleted","Service module resolution code deleted","Major",0,Description
	
	set rsDelete = Nothing
	cnnDelete.Close
	set cnnDelete = Nothing
	
End If

Response.Redirect ("main.asp")
%>