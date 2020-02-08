<!--#include file="../../inc/InsightFuncs_Equipment.asp"-->
<!--#include file="../../inc/InSightFuncs.asp"-->
<!--#include file="../../inc/InSightFuncs_Users.asp"-->
<%
StatusCodeIntRecIDToReplace = Request.Form("txtStatusCodeIntRecIDToReplace")
StatusCodeIntRecIDReplaceWith = Request.Form("selDeleteStatusCodeFromModal")


StatusCodeToDeleteName = GetStatusCodeNameByIntRecID(StatusCodeIntRecIDToReplace)
StatusCodeToReplaceWithName = GetStatusCodeNameByIntRecID(StatusCodeIntRecIDReplaceWith)

If StatusCodeIntRecIDToReplace <> "" AND StatusCodeIntRecIDReplaceWith <> "" Then

	'We need to loop through all the records so we can make entries in the EQ_Activty table
	
	SQLDelete = "SELECT InternalRecordIdentifier From EQ_Equipment WHERE StatusCodeIntRecID = " & StatusCodeIntRecIDToReplace
	Set cnnDelete = Server.CreateObject("ADODB.Connection")
	cnnDelete.open (Session("ClientCnnString"))
	Set rsDelete = Server.CreateObject("ADODB.Recordset")
	rsDelete.CursorLocation = 3 
	Set rsDelete = cnnDelete.Execute(SQLDelete)

	Activity = "The Status Code for Equipment with Serial # " & rsDelete("SerialNumber") & " was changed from ''" & StatusCodeToDeleteName & "'' to ''" & StatusCodeToReplaceWithName & "'' to allow for the deletion of ''" & StatusCodeToDeleteName & "''"

	If not rsDelete.Eof Then
		Do
			Record_EQ_Activity rsDelete("InternalRecordIdentifier"),Activity,Session("UserNo")
			rsDelete.movenext
		Loop Until rsDelete.Eof
	End If
	rsDelete.Close
	
	'Now replace all records with the new status code number
	
	SQLDelete = "UPDATE EQ_Equipment Set StatusCodeIntRecID = " & StatusCodeIntRecIDReplaceWith & " WHERE CurrentStatusCodeIntRecID = " & StatusCodeIntRecIDToReplace
	rsDelete.CursorLocation = 3 
	Set rsDelete = cnnDelete.Execute(SQLDelete)
	
	'Now Do the deletion
	
	SQLDelete = "DELETE FROM EQ_StatusCodes WHERE InternalRecordIdentifier = "& StatusCodeIntRecIDToReplace
	rsDelete.CursorLocation = 3 
	Set rsDelete = cnnDelete.Execute(SQLDelete)
	
	Description = "The " & GetTerm("Equipment") & " status code named " & StatusCodeToDeleteName & " was deleted by " & GetUserDisplayNameByUserNo(Session("UserNo")) 
	CreateAuditLogEntry GetTerm("Equipment") & " Status Code Deleted",GetTerm("Equipment") & " Status Code Deleted","Major",0,Description
	
	set rsDelete = Nothing
	cnnDelete.Close
	set cnnDelete = Nothing
	
End If

Response.Redirect ("main.asp")
%>