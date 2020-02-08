<!--#include file="../../inc/InsightFuncs_Prospecting.asp"-->
<!--#include file="../../inc/InSightFuncs.asp"-->
<!--#include file="../../inc/InSightFuncs_Users.asp"-->
<%
ActivityNoToReplace = Request.Form("txtActivityNoToReplace")
ActivityNoReplaceWith = Request.Form("seldeleteActivityFromModal")

If ActivityNoToReplace <> "" AND ActivityNoReplaceWith <> "" Then

	'We need to loop through all the records in PR_ProspectActivities so we can make entries in the PR_Audit table
	
	SQLDelete = "Select * From PR_ProspectActivities WHERE ActivityRecID = " & ActivityNoToReplace
	Set cnnDelete = Server.CreateObject("ADODB.Connection")
	cnnDelete.open (Session("ClientCnnString"))
	Set rsDelete = Server.CreateObject("ADODB.Recordset")
	rsDelete.CursorLocation = 3 
	Set rsDelete = cnnDelete.Execute(SQLDelete)

	Activity = "The activity for this prospect was changed from ''" & GetActivityByNum(ActivityNoToReplace) & "'' to ''" & GetActivityByNum(ActivityNoReplaceWith) & "'' to allow for the deletion of ''" & GetActivityByNum(ActivityNoToReplace) & "''"
	ActivityDescription = GetActivityByNum(ActivityNoToReplace) ' For audit trail below
	
	If not rsDelete.Eof Then
		Do
			Record_PR_Activity rsDelete("ProspectRecID"),Activity,Session("UserNo")
		
			rsDelete.movenext
		Loop Until rsDelete.Eof
	End If
	rsDelete.Close
	
	'Now replace all records with the new ActivityRecID 
	
	SQLDelete = "UPDATE PR_ProspectActivities Set ActivityRecID = " & ActivityNoReplaceWith & " WHERE ActivityRecID = " & ActivityNoToReplace
	rsDelete.CursorLocation = 3 
	Set rsDelete = cnnDelete.Execute(SQLDelete)
	
	'Now Do the deletion
	
	SQLDelete = "Delete FROM PR_Activities WHERE InternalRecordIdentifier = "& ActivityNoToReplace
	rsDelete.CursorLocation = 3 
	Set rsDelete = cnnDelete.Execute(SQLDelete)
	
	Description = "The " & GetTerm("Prospecting") & " activity named " & ActivityDescription & " was deleted by " & GetUserDisplayNameByUserNo(Session("UserNo")) 
	CreateAuditLogEntry GetTerm("Prospecting") & " activity deleted",GetTerm("Prospecting") & " activity deleted","Major",0,Description
	
	set rsDelete = Nothing
	cnnDelete.Close
	set cnnDelete = Nothing
	
End If

Response.Redirect ("main.asp")
%>