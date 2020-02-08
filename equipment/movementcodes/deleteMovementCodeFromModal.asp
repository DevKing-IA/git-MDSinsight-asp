<!--#include file="../../inc/InsightFuncs_Equipment.asp"-->
<!--#include file="../../inc/InSightFuncs.asp"-->
<!--#include file="../../inc/InSightFuncs_Users.asp"-->
<%
MovementCodeIntRecIDToReplace = Request.Form("txtMovementCodeIntRecIDToReplace")
MovementCodeIntRecIDReplaceWith = Request.Form("selDeleteMovementCodeFromModal")

MovementCodeToDelete = GetMovementCodeByIntRecID(MovementCodeIntRecIDToReplace)
MovementCodeToReplaceWith = GetMovementCodeByIntRecID(MovementCodeIntRecIDReplaceWith)

MovementCodeToDeleteDesc = GetMovementCodeDescByIntRecID(MovementCodeIntRecIDToReplace)
MovementCodeToReplaceWithDesc = GetMovementCodeDescByIntRecID(MovementCodeIntRecIDReplaceWith)

If MovementCodeIntRecIDToReplace <> "" AND MovementCodeIntRecIDReplaceWith <> "" Then

	'We need to loop through all the records so we can make entries in the EQ_Activty table
	
	SQLDelete = "SELECT InternalRecordIdentifier From EQ_Equipment WHERE MovementCodeIntRecID = " & MovementCodeIntRecIDToReplace
	Set cnnDelete = Server.CreateObject("ADODB.Connection")
	cnnDelete.open (Session("ClientCnnString"))
	Set rsDelete = Server.CreateObject("ADODB.Recordset")
	rsDelete.CursorLocation = 3 
	Set rsDelete = cnnDelete.Execute(SQLDelete)

	Activity = "The Movement Code for Equipment with Serial # " & rsDelete("SerialNumber") & " was changed from ''" & MovementCodeToDeleteDesc & "'' to ''" & MovementCodeToReplaceWithDesc & "'' to allow for the deletion of ''" & MovementCodeToDeleteDesc & "''"

	If not rsDelete.Eof Then
		Do
			Record_EQ_Activity rsDelete("InternalRecordIdentifier"),Activity,Session("UserNo")
			rsDelete.movenext
		Loop Until rsDelete.Eof
	End If
	rsDelete.Close
	
	'Now replace all records with the new Movement Code number
	
	SQLDelete = "UPDATE EQ_Equipment Set MovementCodeIntRecID = " & MovementCodeIntRecIDReplaceWith & " WHERE CurrentMovementCodeIntRecID = " & MovementCodeIntRecIDToReplace
	rsDelete.CursorLocation = 3 
	Set rsDelete = cnnDelete.Execute(SQLDelete)
	
	'Now Do the deletion
	
	SQLDelete = "DELETE FROM EQ_MovementCodes WHERE InternalRecordIdentifier = "& MovementCodeIntRecIDToReplace
	rsDelete.CursorLocation = 3 
	Set rsDelete = cnnDelete.Execute(SQLDelete)
	
	Description = "The " & GetTerm("Equipment") & " Movement Code " & MovementCodeToDelete & " (" & MovementCodeToDeleteDesc & ") was deleted by " & GetUserDisplayNameByUserNo(Session("UserNo")) 
	CreateAuditLogEntry GetTerm("Equipment") & " Movement Code Deleted",GetTerm("Equipment") & " Movement Code Deleted","Major",0,Description
	
	set rsDelete = Nothing
	cnnDelete.Close
	set cnnDelete = Nothing
	
End If

Response.Redirect ("main.asp")
%>