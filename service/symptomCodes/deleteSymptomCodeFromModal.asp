<!--#include file="../../inc/InsightFuncs_Prospecting.asp"-->
<!--#include file="../../inc/InsightFuncs.asp"-->
<!--#include file="../../inc/InsightFuncs_Users.asp"-->
<%
SymptomCodeToReplace = Request.Form("txtSymptomCodeToReplace")
SymptomCodeReplaceWith = Request.Form("seldeleteSymptomCodeFromModal")

If SymptomCodeToReplace <> "" AND SymptomCodeReplaceWith <> "" Then
	
	'Now replace all records with the new symptom code
	Set cnnDelete = Server.CreateObject("ADODB.Connection")
	cnnDelete.open (Session("ClientCnnString"))
	Set rsDelete = Server.CreateObject("ADODB.Recordset")

	SQLDelete = "UPDATE FS_ServiceMemos Set SymptomCode = " & SymptomCodeReplaceWith & " WHERE SymptomCode = " & SymptomCodeToReplace 
	Set rsDelete = cnnDelete.Execute(SQLDelete)
	
	'Now Do the deletion
	
	SQLDelete = "Delete FROM FS_SymptomCodes WHERE InternalRecordIdentifier = "& SymptomCodeToReplace 
	rsDelete.CursorLocation = 3 
	Set rsDelete = cnnDelete.Execute(SQLDelete)
	
	Description = "The service module symptom code " & SymptomCodeToReplace & " was deleted by " & GetUserDisplayNameByUserNo(Session("UserNo")) 
	CreateAuditLogEntry "Service module symptom code deleted","Service module symptom code deleted","Major",0,Description
	
	set rsDelete = Nothing
	cnnDelete.Close
	set cnnDelete = Nothing
	
End If

Response.Redirect ("main.asp")
%>