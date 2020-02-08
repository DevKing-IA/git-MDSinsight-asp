<!--#include file="../../inc/InsightFuncs_Prospecting.asp"-->
<!--#include file="../../inc/InsightFuncs.asp"-->
<!--#include file="../../inc/InSightFuncs_Users.asp"-->
<%
IndustryNoToReplace = Request.Form("txtIndustryNoToReplace")
IndustryNoReplaceWith = Request.Form("seldeleteIndustryFromModal")

If IndustryNoToReplace <> "" AND IndustryNoReplaceWith <> "" Then

	'We need to loop through all the records so we can make entries in the PR_Activty table
	
	SQLDelete = "Select InternalRecordIdentifier From PR_Prospects WHERE IndustryNumber = " & IndustryNoToReplace
	Set cnnDelete = Server.CreateObject("ADODB.Connection")
	cnnDelete.open (Session("ClientCnnString"))
	Set rsDelete = Server.CreateObject("ADODB.Recordset")
	rsDelete.CursorLocation = 3 
	Set rsDelete = cnnDelete.Execute(SQLDelete)

	Activity = "The industry for this prospect was changed from ''" & GetIndustryByNum(IndustryNoToReplace) & "'' to ''" & GetIndustryByNum(IndustryNoReplaceWith) & "'' to allow for the deletion of ''" & GetIndustryByNum(IndustryNoToReplace) & "''"
	IndustryDescription = GetIndustryByNum(IndustryNoToReplace) ' For audit trail below
	
	If not rsDelete.Eof Then
		Do
			Record_PR_Activity rsDelete("InternalRecordIdentifier"),Activity,Session("UserNo")
		
			rsDelete.movenext
		Loop Until rsDelete.Eof
	End If
	rsDelete.Close
	
	'Now replace all records with the new industry number
	
	SQLDelete = "UPDATE PR_Prospects Set IndustryNumber = " & IndustryNoReplaceWith & " WHERE IndustryNumber = " & IndustryNoToReplace
	rsDelete.CursorLocation = 3 
	Set rsDelete = cnnDelete.Execute(SQLDelete)
	
	'Now Do the deletion
	
	SQLDelete = "Delete FROM PR_Industries WHERE InternalRecordIdentifier = "& IndustryNoToReplace
	rsDelete.CursorLocation = 3 
	Set rsDelete = cnnDelete.Execute(SQLDelete)
	
	Description = "The " & GetTerm("Prospecting") & " industry named " & IndustryDescription & " was deleted by " & GetUserDisplayNameByUserNo(Session("UserNo")) 
	CreateAuditLogEntry GetTerm("Prospecting") & " industry deleted",GetTerm("Prospecting") & " industry deleted","Major",0,Description
	
	set rsDelete = Nothing
	cnnDelete.Close
	set cnnDelete = Nothing
	
End If

Response.Redirect ("main.asp")
%>