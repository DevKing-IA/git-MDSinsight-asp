<!--#include file="../../../inc/InsightFuncs_Service.asp"-->
<!--#include file="../../../inc/InSightFuncs.asp"-->
<!--#include file="../../../inc/InSightFuncs_Users.asp"-->
<%
FilterIntRecIDToReplace = Request.Form("txtFilterIntRecIDToReplace")
FilterIntRecIDReplaceWith = Request.Form("selDeleteFilterIntRecIDFromModal")

FilterToDelete = GetFilterIDByIntRecID(FilterIntRecIDToReplace)
FilterToReplaceWith = GetFilterIDByIntRecID(FilterIntRecIDReplaceWith)

FilterToDeleteDesc = GetFilterDescByIntRecID(FilterIntRecIDToReplace)
FilterToReplaceWithDesc = GetFilterDescByIntRecID(FilterIntRecIDReplaceWith)

If FilterIntRecIDToReplace <> "" AND FilterIntRecIDReplaceWith <> "" Then

	'We need to loop through all the records so we can make entries in the EQ_Activty table
	
	SQLDelete = "SELECT * From FS_CustomerFilters WHERE FilterIntRecID = " & FilterIntRecIDToReplace
	Set cnnDelete = Server.CreateObject("ADODB.Connection")
	cnnDelete.open (Session("ClientCnnString"))
	Set rsDelete = Server.CreateObject("ADODB.Recordset")
	rsDelete.CursorLocation = 3 
	Set rsDelete = cnnDelete.Execute(SQLDelete)

	Activity = "The Filter For Customer # " & rsDelete("CustID") & " was changed from ''" & FilterToDeleteDesc & "'' to ''" & FilterToReplaceWithDesc & "'' to allow for the deletion of ''" & FilterToDeleteDesc & "''"

	If not rsDelete.Eof Then
		Do
			CreateAuditLogEntry GetTerm("Service") & " Filter Replaced",GetTerm("Service") & " Filter Replaced","Minor",0,Activity
			rsDelete.movenext
		Loop Until rsDelete.Eof
	End If
	rsDelete.Close
	
	'Now replace all records with the new filter int rec id number
	
	SQLDelete = "UPDATE FS_CustomerFilters Set FilterIntRecID = " & FilterIntRecIDReplaceWith & " WHERE FilterIntRecID = " & FilterIntRecIDToReplace
	rsDelete.CursorLocation = 3 
	Set rsDelete = cnnDelete.Execute(SQLDelete)
	
	'Now Do the deletion
	
	SQLDelete = "DELETE FROM IC_Filters WHERE InternalRecordIdentifier = " & FilterIntRecIDToReplace
	rsDelete.CursorLocation = 3 
	Set rsDelete = cnnDelete.Execute(SQLDelete)
	
	Description = "The " & GetTerm("Service") & " Filter " & FilterToDelete & " (" & FilterToDeleteDesc & ") was deleted by " & GetUserDisplayNameByUserNo(Session("UserNo")) 
	CreateAuditLogEntry GetTerm("Service") & " Filter Deleted",GetTerm("Service") & " Filter Deleted","Major",0,Description
	
	set rsDelete = Nothing
	cnnDelete.Close
	set cnnDelete = Nothing
	
End If

Response.Redirect ("main.asp")
%>