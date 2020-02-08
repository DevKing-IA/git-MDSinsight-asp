<!--#include file="../../inc/InsightFuncs_Prospecting.asp"-->
<!--#include file="../../inc/InsightFuncs.asp"-->
<!--#include file="../../inc/InsightFuncs_Users.asp"-->
<%
ProblemCodeToReplace = Request.Form("txtProblemCodeToReplace")
ProblemCodeReplaceWith = Request.Form("seldeleteProblemCodeFromModal")

If ProblemCodeToReplace <> "" AND ProblemCodeReplaceWith <> "" Then
	
	'Now replace all records with the new problem code
	Set cnnDelete = Server.CreateObject("ADODB.Connection")
	cnnDelete.open (Session("ClientCnnString"))
	Set rsDelete = Server.CreateObject("ADODB.Recordset")

	SQLDelete = "UPDATE FS_ServiceMemos Set ProblemCode = " & ProblemCodeReplaceWith & " WHERE ProblemCode = " & ProblemCodeToReplace 
	Set rsDelete = cnnDelete.Execute(SQLDelete)
	
	'Now Do the deletion
	
	SQLDelete = "Delete FROM FS_ProblemCodes WHERE InternalRecordIdentifier = "& ProblemCodeToReplace 
	rsDelete.CursorLocation = 3 
	Set rsDelete = cnnDelete.Execute(SQLDelete)
	
	Description = "The service module problem code " & ProblemCodeToReplace & " was deleted by " & GetUserDisplayNameByUserNo(Session("UserNo")) 
	CreateAuditLogEntry "Service module problem code deleted","Service module problem code deleted","Major",0,Description
	
	set rsDelete = Nothing
	cnnDelete.Close
	set cnnDelete = Nothing
	
End If

Response.Redirect ("main.asp")
%>