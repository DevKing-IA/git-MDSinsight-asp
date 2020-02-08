<!--#include file="../../inc/InsightFuncs_Prospecting.asp"-->
<!--#include file="../../inc/InSightFuncs.asp"-->
<!--#include file="../../inc/InSightFuncs_Users.asp"-->
<%
CompetitorNoToReplace = Request.Form("txtCompetitorNoToReplace")
CompetitorNoReplaceWith = Request.Form("seldeleteCompetitorFromModal")

If CompetitorNoToReplace <> "" AND CompetitorNoReplaceWith <> "" Then

	'We need to loop through all the records so we can make entries in the PR_Audit table
	
	SQLDelete = "Select InternalRecordIdentifier From PR_ProspectCompetitors WHERE CompetitorRecID = " & CompetitorNoToReplace
	Set cnnDelete = Server.CreateObject("ADODB.Connection")
	cnnDelete.open (Session("ClientCnnString"))
	Set rsDelete = Server.CreateObject("ADODB.Recordset")
	rsDelete.CursorLocation = 3 
	Set rsDelete = cnnDelete.Execute(SQLDelete)

	Activity = "The Competitor for this prospect was changed from ''" & GetCompetitorByNum(CompetitorNoToReplace) & "'' to ''" & GetCompetitorByNum (CompetitorNoReplaceWith) & "'' to allow for the deletion of ''" & GetCompetitorByNum (CompetitorNoToReplace) & "''"
	CompetitorDescription = GetCompetitorByNum(CompetitorNoToReplace) ' For audit trail below
	
	If not rsDelete.Eof Then
		Do
			Record_PR_Activity rsDelete("InternalRecordIdentifier"),Activity,Session("UserNo")
		
			rsDelete.movenext
		Loop Until rsDelete.Eof
	End If
	rsDelete.Close
	
	'Now replace all records with the new Competitor number
	
	SQLDelete = "UPDATE PR_ProspectCompetitors Set CompetitorRecID = " & CompetitorNoReplaceWith & " WHERE CompetitorRecID = " & CompetitorNoToReplace
	rsDelete.CursorLocation = 3 
	Set rsDelete = cnnDelete.Execute(SQLDelete)
	
	'Now Do the deletion
	
	SQLDelete = "Delete FROM PR_Competitors WHERE InternalRecordIdentifier = "& CompetitorNoToReplace
	rsDelete.CursorLocation = 3 
	Set rsDelete = cnnDelete.Execute(SQLDelete)
	
	Description = "The " & GetTerm("Prospecting") & " Competitor named " & CompetitorName  & " was deleted by " & GetUserDisplayNameByUserNo(Session("UserNo")) 
	CreateAuditLogEntry GetTerm("Prospecting") & " Competitor deleted",GetTerm("Prospecting") & " Competitor deleted","Major",0,Description
	
	set rsDelete = Nothing
	cnnDelete.Close
	set cnnDelete = Nothing
	
End If

Response.Redirect ("main.asp")
%>