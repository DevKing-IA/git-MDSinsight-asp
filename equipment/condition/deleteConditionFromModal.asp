<!--#include file="../../inc/InsightFuncs_Equipment.asp"-->
<!--#include file="../../inc/InSightFuncs.asp"-->
<!--#include file="../../inc/InSightFuncs_Users.asp"-->
<%
ConditionIntRecIDToReplace = Request.Form("txtConditionIntRecIDToReplace")
ConditionIntRecIDReplaceWith = Request.Form("seldeleteConditionFromModal")


ConditionToDeleteName = GetConditionNameByIntRecID(ConditionIntRecIDToReplace)
ConditionToReplaceWithName = GetConditionNameByIntRecID(ConditionIntRecIDReplaceWith)

If ConditionIntRecIDToReplace <> "" AND ConditionIntRecIDReplaceWith <> "" Then

	'We need to loop through all the records so we can make entries in the EQ_Activty table
	
	SQLDelete = "Select InternalRecordIdentifier From EQ_Equipment WHERE CurrentConditionIntRecID = " & ConditionIntRecIDToReplace
	Set cnnDelete = Server.CreateObject("ADODB.Connection")
	cnnDelete.open (Session("ClientCnnString"))
	Set rsDelete = Server.CreateObject("ADODB.Recordset")
	rsDelete.CursorLocation = 3 
	Set rsDelete = cnnDelete.Execute(SQLDelete)

	Activity = "The condition for this prospect was changed from ''" & ConditionToDeleteName & "'' to ''" & ConditionToReplaceWithName & "'' to allow for the deletion of ''" & ConditionToDeleteName & "''"

	If not rsDelete.Eof Then
		Do
			Record_EQ_Activity rsDelete("InternalRecordIdentifier"),Activity,Session("UserNo")
			rsDelete.movenext
		Loop Until rsDelete.Eof
	End If
	rsDelete.Close
	
	'Now replace all records with the new condition number
	
	SQLDelete = "UPDATE EQ_Equipment Set CurrentConditionIntRecID = " & ConditionIntRecIDReplaceWith & " WHERE CurrentConditionIntRecID = " & ConditionIntRecIDToReplace
	rsDelete.CursorLocation = 3 
	Set rsDelete = cnnDelete.Execute(SQLDelete)
	
	'Now Do the deletion
	
	SQLDelete = "Delete FROM EQ_Condition WHERE InternalRecordIdentifier = "& ConditionIntRecIDToReplace
	rsDelete.CursorLocation = 3 
	Set rsDelete = cnnDelete.Execute(SQLDelete)
	
	Description = "The " & GetTerm("Equipment") & " condition named " & ConditionToDeleteName & " was deleted by " & GetUserDisplayNameByUserNo(Session("UserNo")) 
	CreateAuditLogEntry GetTerm("Equipment") & " Condition Deleted",GetTerm("Equipment") & " Condition Deleted","Major",0,Description
	
	set rsDelete = Nothing
	cnnDelete.Close
	set cnnDelete = Nothing
	
End If

Response.Redirect ("main.asp")
%>