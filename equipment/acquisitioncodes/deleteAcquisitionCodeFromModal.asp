<!--#include file="../../inc/InsightFuncs_Equipment.asp"-->
<!--#include file="../../inc/InSightFuncs.asp"-->
<!--#include file="../../inc/InSightFuncs_Users.asp"-->

<%
acquisitionCodeIntRecIDToReplace = Request.Form("txtacquisitionCodeIntRecIDToReplace")
acquisitionCodeIntRecIDReplaceWith = Request.Form("selDeleteacquisitionCodeFromModal")

acquisitionCodeToDelete = GetAcquisitionCodeByIntRecID(acquisitionCodeIntRecIDToReplace)
acquisitionCodeToReplaceWith = GetAcquisitionCodeByIntRecID(acquisitionCodeIntRecIDReplaceWith)

acquisitionCodeToDeleteDesc = GetAcquisitionCodeDescByIntRecID(acquisitionCodeIntRecIDToReplace)
acquisitionCodeToReplaceWithDesc = GetAcquisitionCodeDescByIntRecID(acquisitionCodeIntRecIDReplaceWith)

If acquisitionCodeIntRecIDToReplace <> "" AND acquisitionCodeIntRecIDReplaceWith <> "" Then

	'We need to loop through all the records so we can make entries in the EQ_Activty table
	
	SQLDelete = "SELECT InternalRecordIdentifier From EQ_Equipment WHERE CurrentConditionIntRecID = " & acquisitionCodeIntRecIDToReplace
	Set cnnDelete = Server.CreateObject("ADODB.Connection")
	cnnDelete.open (Session("ClientCnnString"))
	Set rsDelete = Server.CreateObject("ADODB.Recordset")
	rsDelete.CursorLocation = 3 
	Set rsDelete = cnnDelete.Execute(SQLDelete)

	Activity = "The Acquisition Code for Equipment with Serial # " & rsDelete("SerialNumber") & " was changed from ''" & acquisitionCodeToDeleteDesc & "'' to ''" & acquisitionCodeToReplaceWithDesc & "'' to allow for the deletion of ''" & acquisitionCodeToDeleteDesc & "''"

	If not rsDelete.Eof Then
		Do
			Record_EQ_Activity rsDelete("InternalRecordIdentifier"),Activity,Session("UserNo")
			rsDelete.movenext
		Loop Until rsDelete.Eof
	End If
	rsDelete.Close
	
	'Now replace all records with the new Acquisition Code number
	
	SQLDelete = "UPDATE EQ_Equipment SET CurrentConditionIntRecID = " & acquisitionCodeIntRecIDReplaceWith & " WHERE CurrentAcquisitionCodeIntRecID = " & acquisitionCodeIntRecIDToReplace
	rsDelete.CursorLocation = 3 
	Set rsDelete = cnnDelete.Execute(SQLDelete)
	
	'Now Do the deletion
	
	SQLDelete = "DELETE FROM EQ_AcquisitionCodes WHERE InternalRecordIdentifier = "& acquisitionCodeIntRecIDToReplace
	rsDelete.CursorLocation = 3 
	Set rsDelete = cnnDelete.Execute(SQLDelete)
	
	Description = "The " & GetTerm("Equipment") & " Acquisition Code " & acquisitionCodeToDelete & " (" & acquisitionCodeToDeleteDesc & ") was deleted by " & GetUserDisplayNameByUserNo(Session("UserNo")) 
	CreateAuditLogEntry GetTerm("Equipment") & " Acquisition Code Deleted",GetTerm("Equipment") & " Acquisition Code Deleted","Major",0,Description
	
	set rsDelete = Nothing
	cnnDelete.Close
	set cnnDelete = Nothing
	
End If

Response.Redirect ("main.asp")
%>