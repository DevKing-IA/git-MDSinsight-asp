<!--#include file="../../inc/InsightFuncs_Prospecting.asp"-->
<!--#include file="../../inc/InSightFuncs.asp"-->
<!--#include file="../../inc/InSightFuncs_Users.asp"-->
<%
LeadSourceNoToReplace = Request.Form("txtLeadSourceNoToReplace")
LeadSourceNoReplaceWith = Request.Form("seldeleteLeadSourceFromModal")

If LeadSourceNoToReplace <> "" AND LeadSourceNoReplaceWith <> "" Then

	'We need to loop through all the records so we can make entries in the PR_Activty table
	
	SQLDelete = "Select InternalRecordIdentifier From PR_Prospects WHERE LeadSourceNumber = " & LeadSourceNoToReplace
	Set cnnDelete = Server.CreateObject("ADODB.Connection")
	cnnDelete.open (Session("ClientCnnString"))
	Set rsDelete = Server.CreateObject("ADODB.Recordset")
	rsDelete.CursorLocation = 3 
	Set rsDelete = cnnDelete.Execute(SQLDelete)

	Activity = "The lead source for this prospect was changed from ''" & GetLeadSourceByNum(LeadSourceNoToReplace) & "'' to ''" & GetLeadSourceByNum (LeadSourceNoReplaceWith) & "'' to allow for the deletion of ''" & GetLeadSourceByNum (LeadSourceNoToReplace) & "''"
	LeadSourceDescription = GetLeadSourceByNum(LeadSourceNoToReplace) ' For audit trail below
	
	If not rsDelete.Eof Then
		Do
			Record_PR_Activity rsDelete("InternalRecordIdentifier"),Activity,Session("UserNo")
		
			rsDelete.movenext
		Loop Until rsDelete.Eof
	End If
	rsDelete.Close
	
	'Now replace all records with the new lead source number
	
	SQLDelete = "UPDATE PR_Prospects Set LeadSourceNumber = " & LeadSourceNoReplaceWith & " WHERE LeadSourceNumber = " & LeadSourceNoToReplace
	rsDelete.CursorLocation = 3 
	Set rsDelete = cnnDelete.Execute(SQLDelete)
	
	'Now Do the deletion
	
	SQLDelete = "Delete FROM PR_LeadSources WHERE InternalRecordIdentifier = "& LeadSourceNoToReplace
	rsDelete.CursorLocation = 3 
	Set rsDelete = cnnDelete.Execute(SQLDelete)
	
	Description = "The " & GetTerm("Prospecting") & " lead source named " & LeadSourceDescription & " was deleted by " & GetUserDisplayNameByUserNo(Session("UserNo")) 
	CreateAuditLogEntry GetTerm("Prospecting") & " lead source deleted",GetTerm("Prospecting") & " lead source deleted","Major",0,Description
	
	set rsDelete = Nothing
	cnnDelete.Close
	set cnnDelete = Nothing
	
End If

Response.Redirect ("main.asp")
%>