<!--#include file="../../inc/InsightFuncs_Prospecting.asp"-->
<!--#include file="../../inc/InSightFuncs.asp"-->
<!--#include file="../../inc/InSightFuncs_Users.asp"-->
<%
ReasonNoToReplace = Request.Form("txtReasonNoToReplace")
ReasonNoReplaceWith = Request.Form("seldeleteReasonFromModal")

If ReasonNoToReplace <> "" AND ReasonNoReplaceWith <> "" Then

	'We need to loop through all the records so we can make entries in the PR_Activty table
	
	SQLDelete = "Select InternalRecordIdentifier From PR_Prospects WHERE ReasonNumber = " & ReasonNoToReplace
	Set cnnDelete = Server.CreateObject("ADODB.Connection")
	cnnDelete.open (Session("ClientCnnString"))
	Set rsDelete = Server.CreateObject("ADODB.Recordset")
	rsDelete.CursorLocation = 3 
	Set rsDelete = cnnDelete.Execute(SQLDelete)

	Activity = "The reason for this prospect was changed from ''" & GetReasonByNum(ReasonNoToReplace) & "'' to ''" & GetReasonByNum(ReasonNoReplaceWith) & "'' to allow for the deletion of ''" & GetReasonByNum(ReasonNoToReplace) & "''"
	ReasonDescription = GetReasonByNum(ReasonNoToReplace) ' For audit trail below
	
	If not rsDelete.Eof Then
		Do
			Record_PR_Activity rsDelete("InternalRecordIdentifier"),Activity,Session("UserNo")
		
			rsDelete.movenext
		Loop Until rsDelete.Eof
	End If
	rsDelete.Close
	
	'Now replace all records with the new Reason number
	
	SQLDelete = "UPDATE PR_Prospects Set ReasonNumber = " & ReasonNoReplaceWith & " WHERE ReasonNumber = " & ReasonNoToReplace
	rsDelete.CursorLocation = 3 
	Set rsDelete = cnnDelete.Execute(SQLDelete)
	
	'Now Do the deletion
	
	SQLDelete = "Delete FROM PR_Reasons WHERE InternalRecordIdentifier = "& ReasonNoToReplace
	rsDelete.CursorLocation = 3 
	Set rsDelete = cnnDelete.Execute(SQLDelete)
	
	Description = "The " & GetTerm("Prospecting") & " reason named " & ReasonDescription & " was deleted by " & GetUserDisplayNameByUserNo(Session("UserNo")) 
	CreateAuditLogEntry GetTerm("Prospecting") & " reason deleted",GetTerm("Prospecting") & " reason deleted","Major",0,Description
	
	set rsDelete = Nothing
	cnnDelete.Close
	set cnnDelete = Nothing
	
End If

Response.Redirect ("main.asp")
%>