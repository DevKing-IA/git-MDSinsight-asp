<!--#include file="../../inc/InsightFuncs_Equipment.asp"-->
<!--#include file="../../inc/InSightFuncs.asp"-->
<!--#include file="../../inc/InSightFuncs_Users.asp"-->
<%

GroupIntRecIDToReplace = Request.Form("txtGroupIntRecIDToReplace")
GroupIntRecIDReplaceWith = Request.Form("selDeleteGroupFromModal")
GroupToBeReplacedDescription = GetGroupNameByIntRecID(GroupIntRecIDToReplace)
GroupToReplaceWithDescription = GetGroupNameByIntRecID(GroupIntRecIDReplaceWith)

If GroupIntRecIDToReplace <> "" AND GroupIntRecIDReplaceWith <> "" Then

	'We need to loop through all the records so we can make entries in the EQ_Activty table
	
	SQLDelete = "SELECT * FROM EQ_Equipment WHERE GroupNumber = " & GroupIntRecIDToReplace
	Set cnnDelete = Server.CreateObject("ADODB.Connection")
	cnnDelete.open (Session("ClientCnnString"))
	Set rsDelete = Server.CreateObject("ADODB.Recordset")
	rsDelete.CursorLocation = 3 
	Set rsDelete = cnnDelete.Execute(SQLDelete)

	Activity = "The Equipment Group for Equipment with Serial # " & rsDelete("SerialNumber") & " was changed from ''" &  GroupToBeReplacedDescription & "'' to ''" &  GroupToReplaceWithDescription & "'' to allow for the deletion of ''" &  GroupToBeReplacedDescription & "''"
	
	If not rsDelete.Eof Then
		Do
			Record_EQ_Activity rsDelete("InternalRecordIdentifier"),Activity,Session("UserNo")
			rsDelete.movenext
		Loop Until rsDelete.Eof
	End If
	rsDelete.Close
	
	'Now replace all records with the new Group number
	
	SQLDelete = "UPDATE EQ_Equipment SET GroupIntRecID = " & GroupIntRecIDReplaceWith & " WHERE GroupIntRecID = " & GroupIntRecIDToReplace
	rsDelete.CursorLocation = 3 
	Set rsDelete = cnnDelete.Execute(SQLDelete)
	
	'Now Do the deletion
	
	SQLDelete = "DELETE FROM EQ_Groups WHERE InternalRecordIdentifier = "& GroupIntRecIDToReplace
	rsDelete.CursorLocation = 3 
	Set rsDelete = cnnDelete.Execute(SQLDelete)
	
	Description = "The " & GetTerm("Equipment") & " Group named " & GroupToBeReplacedDescription & " was deleted by " & GetUserDisplayNameByUserNo(Session("UserNo")) 
	CreateAuditLogEntry GetTerm("Equipment") & " Group Deleted",GetTerm("Equipment") & " Group Deleted","Major",0,Description
	
	set rsDelete = Nothing
	cnnDelete.Close
	set cnnDelete = Nothing
	
End If

Response.Redirect ("main.asp")
%>