<!--#include file="../../inc/InsightFuncs_Prospecting.asp"-->
<!--#include file="../../inc/InSightFuncs.asp"-->
<!--#include file="../../inc/InSightFuncs_Users.asp"-->
<%
StageNoToReplace = Request.Form("txtStageNoToReplace")
StageNoReplaceWith = Request.Form("seldeleteStagesFromModal")

If StageNoToReplace <> "" AND StageNoReplaceWith <> "" Then

	'We need to loop through all the records in PR_ProspectStages so we can make entries in the PR_Audit table
	
	SQLDelete = "Select * From PR_ProspectStages WHERE StageRecID = " & StageNoToReplace
	Set cnnDelete = Server.CreateObject("ADODB.Connection")
	cnnDelete.open (Session("ClientCnnString"))
	Set rsDelete = Server.CreateObject("ADODB.Recordset")
	rsDelete.CursorLocation = 3 
	Set rsDelete = cnnDelete.Execute(SQLDelete)

	Activity = "The stage for this prospect was changed from ''" & GetStageByNum(StageNoToReplace) & "'' to ''" & GetStageByNum(StageNoReplaceWith) & "'' to allow for the deletion of ''" & GetStageByNum(StageNoToReplace) & "''"
	
	StagesDescription = GetStageByNum(StageNoToReplace) ' For audit trail below
	
	If not rsDelete.Eof Then
		Do
			Record_PR_Activity rsDelete("ProspectRecID"),Activity,Session("UserNo")
		
			rsDelete.movenext
		Loop Until rsDelete.Eof
	End If
	rsDelete.Close
	
	'Now replace all records with the new stage number
	
	SQLDelete = "UPDATE PR_ProspectStages Set StageRecID = " & StageNoReplaceWith & " WHERE StageRecID = " & StageNoToReplace
	rsDelete.CursorLocation = 3 
	Set rsDelete = cnnDelete.Execute(SQLDelete)
	
	'Now Do the deletion
	
	SQLDelete = "Delete FROM PR_Stages WHERE InternalRecordIdentifier = "& StageNoToReplace
	rsDelete.CursorLocation = 3 
	Set rsDelete = cnnDelete.Execute(SQLDelete)
	
	Description = "The " & GetTerm("Prospecting") & " stage named " & StagesDescription & " was deleted by " & GetUserDisplayNameByUserNo(Session("UserNo")) 
	CreateAuditLogEntry GetTerm("Prospecting") & " stage deleted",GetTerm("Prospecting") & " stage deleted","Major",0,Description
	
	set rsDelete = Nothing
	cnnDelete.Close
	set cnnDelete = Nothing
	
End If

Response.Redirect ("main.asp")
%>