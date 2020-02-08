<!--#include file="../../inc/InsightFuncs_Equipment.asp"-->
<!--#include file="../../inc/InSightFuncs.asp"-->
<!--#include file="../../inc/InSightFuncs_Users.asp"-->
<%
ModelIntRecIDToReplace = Request.Form("txtModelIntRecIDToReplace")
ModelToDeleteName = GetModelNameByIntRecID(ModelIntRecIDToReplace)
ModelIntRecIDToReplaceWith = Request.Form("seldeleteModelFromModel")
ModelToReplaceWithName = GetModelNameByIntRecID(ModelIntRecIDToReplaceWith) 

If  ModelIntRecIDToReplace <> "" AND  ModelIntRecIDToReplaceWith <> "DELETE_MODEL_AND_EQUIPMENT" Then

	'---------------------------------------------------------------------------------------------------
	'update the audit trail with all the equipment records that are about to be assigned to a new model
	'---------------------------------------------------------------------------------------------------

	SQLDelete = "SELECT * FROM EQ_Equipment WHERE ModelIntRecID = " & ModelIntRecIDToReplace
	Set cnnDelete = Server.CreateObject("ADODB.Connection")
	cnnDelete.open (Session("ClientCnnString"))
	Set rsDelete = Server.CreateObject("ADODB.Recordset")
	rsDelete.CursorLocation = 3 
	Set rsDelete = cnnDelete.Execute(SQLDelete)
	If not rsDelete.Eof Then
		Do
			Activity = "The  Model for Equipment with Serial # " & rsDelete(" SerialNumber") & " was changed from ''" &  ModelToDeleteName & "'' to ''" &  ModelToReplaceWithName & "'' to allow for the deletion of ''" &  ModelToDeleteName & "''"
			CreateAuditLogEntry GetTerm("Equipment") & " Model Changed For Equipment Record",GetTerm("Equipment") & " Model Changed For Equipment Record","Major",0,Activity
			rsDelete.movenext
		Loop Until rsDelete.Eof
	End If
	rsDelete.Close

	'--------------------------------------------------------------------------------------------
	'Now replace all Model records with a new Model Name
	'--------------------------------------------------------------------------------------------
	
	SQLDelete = "UPDATE EQ_Models Set Model = " &  ModelToReplaceWithName & " WHERE InternalRecordIdentifier = " &  ModelIntRecIDToReplace
	rsDelete.CursorLocation = 3 
	Set rsDelete = cnnDelete.Execute(SQLDelete)	
	
	'--------------------------------------------------------------------------------------------
	'Now replace all EQUIPMENT records with a new Model Int RecID
	'--------------------------------------------------------------------------------------------
	
	SQLDelete = "UPDATE EQ_Equipment Set ModelIntRecID = " &  ModelIntRecIDToReplaceWith & " WHERE  ModelIntRecID = " &  ModelIntRecIDToReplace
	rsDelete.CursorLocation = 3 
	Set rsDelete = cnnDelete.Execute(SQLDelete)
	
	'--------------------------------------------------------------------------------------------
	'Now delete the Model from the EQ_Models tables
	'--------------------------------------------------------------------------------------------
	
	'Now Do the deletion
	
	SQLDelete = "Delete FROM EQ_Models WHERE InternalRecordIdentifier = " & ModelIntRecIDToReplace
	rsDelete.CursorLocation = 3 
	Set rsDelete = cnnDelete.Execute(SQLDelete)
	
	Description = "The " & GetTerm("Equipment") & " model named " & ModelToDeleteName & " was deleted by " & GetUserDisplayNameByUserNo(Session("UserNo")) 
	CreateAuditLogEntry GetTerm("Equipment") & " Model Deleted",GetTerm("Equipment") & " Model Deleted","Major",0,Description
	
	set rsDelete = Nothing
	cnnDelete.Close
	set cnnDelete = Nothing
	
Else

'******************************************************************************
'User wants to delete Model AND all its associated equipment
'******************************************************************************
	'--------------------------------------------------------------------------------------------
	'Delete the Model from the EQ_Models tables
	'--------------------------------------------------------------------------------------------
	
	'Now Do the deletion
	
	SQLDelete = "Delete FROM EQ_Models WHERE InternalRecordIdentifier = " & ModelIntRecIDToReplace
	rsDelete.CursorLocation = 3 
	Set rsDelete = cnnDelete.Execute(SQLDelete)
	
	Description = "The " & GetTerm("Equipment") & " model named " & ModelToDeleteName & " was deleted by " & GetUserDisplayNameByUserNo(Session("UserNo")) 
	CreateAuditLogEntry GetTerm("Equipment ") & " Model Deleted",GetTerm("Equipment") & " Model Deleted","Major",0,Description
	
	set rsDelete = Nothing
	cnnDelete.Close
	set cnnDelete = Nothing

	'--------------------------------------------------------------------------------------------
	'Delete the equipment associated with this Manufacturer from the equipment table
	'--------------------------------------------------------------------------------------------
	
	SQLDelete = "DELETE FROM EQ_Equipment WHERE ModelIntRecID = " & ModelIntRecIDToReplace
	rsDelete.CursorLocation = 3 
	Set rsDelete = cnnDelete.Execute(SQLDelete)
	
	Description = "The " & GetTerm("Equipment") & " table records associated with the model, " &  ModelToDeleteName & ", were deleted by " & GetUserDisplayNameByUserNo(Session("UserNo")) 
	CreateAuditLogEntry GetTerm("Equipment") & "  Equipment records deleted by model deletion",GetTerm("Equipment") & " Equipment records deleted by model deletion","Major",0,Description
	
End If

Response.Redirect ("main.asp")
%>