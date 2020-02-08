<!--#include file="../../inc/InsightFuncs_Equipment.asp"-->
<!--#include file="../../inc/InSightFuncs.asp"-->
<!--#include file="../../inc/InSightFuncs_Users.asp"-->
<%

ClassIntRecIDToReplace = Request.Form("txtClassIntRecIDToReplace")
ClassIntRecIDReplaceWith = Request.Form("selDeleteClassFromModal")
ClassToBeReplacedDescription = GetClassNameByIntRecID(ClassIntRecIDToReplace)
ClassToReplaceWithDescription = GetClassNameByIntRecID(ClassIntRecIDReplaceWith)

If ClassIntRecIDToReplace <> "" AND ClassIntRecIDReplaceWith <> "" Then

	'We need to loop through all the records so we can make entries in the EQ_Activty table
	
	SQLDelete = "SELECT * FROM EQ_Equipment WHERE ClassNumber = " & ClassIntRecIDToReplace
	Set cnnDelete = Server.CreateObject("ADODB.Connection")
	cnnDelete.open (Session("ClientCnnString"))
	Set rsDelete = Server.CreateObject("ADODB.Recordset")
	rsDelete.CursorLocation = 3 
	Set rsDelete = cnnDelete.Execute(SQLDelete)

	Activity = "The Class for Equipment with Serial # " & rsDelete("SerialNumber") & " was changed from ''" &  ClassToBeReplacedDescription & "'' to ''" &  ClassToReplaceWithDescription & "'' to allow for the deletion of ''" &  ClassToBeReplacedDescription & "''"
	
	If not rsDelete.Eof Then
		Do
			Record_EQ_Activity rsDelete("InternalRecordIdentifier"),Activity,Session("UserNo")
			rsDelete.movenext
		Loop Until rsDelete.Eof
	End If
	rsDelete.Close
	
	'Now replace all records with the new class number
	
	SQLDelete = "UPDATE EQ_Equipment SET ClassIntRecID = " & ClassIntRecIDReplaceWith & " WHERE ClassIntRecID = " & ClassIntRecIDToReplace
	rsDelete.CursorLocation = 3 
	Set rsDelete = cnnDelete.Execute(SQLDelete)
	
	'Now Do the deletion
	
	SQLDelete = "DELETE FROM EQ_Classes WHERE InternalRecordIdentifier = "& ClassIntRecIDToReplace
	rsDelete.CursorLocation = 3 
	Set rsDelete = cnnDelete.Execute(SQLDelete)
	
	Description = "The " & GetTerm("Equipment") & " class named " & ClassToBeReplacedDescription & " was deleted by " & GetUserDisplayNameByUserNo(Session("UserNo")) 
	CreateAuditLogEntry GetTerm("Equipment") & " Class Deleted",GetTerm("Equipment") & " Class Deleted","Major",0,Description
	
	set rsDelete = Nothing
	cnnDelete.Close
	set cnnDelete = Nothing
	
End If

Response.Redirect ("main.asp")
%>