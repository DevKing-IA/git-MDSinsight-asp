<!--#include file="../../inc/InsightFuncs_Equipment.asp"-->
<!--#include file="../../inc/InSightFuncs.asp"-->
<!--#include file="../../inc/InSightFuncs_Users.asp"-->
<%

 ManufacturerIntRecIDToReplace = Request.Form("txtManufacturerIntRecIDToReplace")
 ManufacturerToDeleteName = GetManufacturerNameByIntRecID(ManufacturerIntRecIDToReplace)
 ManufacturerIntRecIDToReplaceWith = Request.Form("selDeleteManufacturerFromModal")
 ManufacturerToReplaceWithName = GetManufacturerNameByIntRecID(ManufacturerIntRecIDToReplaceWith) 

If ManufacturerIntRecIDToReplace <> "" AND ManufacturerIntRecIDToReplaceWith <> "DELETE_MANUFACTURER_AND_EQUIPMENT" Then

	 ManufacturerToReplaceWithName = GetManufacturerNameByIntRecID(ManufacturerIntRecIDToReplaceWith)

	'**************************************************************************************
	'User wants to delete  Manufacturer and select a replacement  Manufacturer for all defined SKUS
	'**************************************************************************************
	
	'--------------------------------------------------------------------------------------------
	'update the audit trail with all the skus that are about to be assigned to a new  Manufacturer
	'--------------------------------------------------------------------------------------------

	SQLDelete = "SELECT * FROM EQ_Equipment WHERE ManufacIntRecID = " &  ManufacturerIntRecIDToReplace
	Set cnnDelete = Server.CreateObject("ADODB.Connection")
	cnnDelete.open (Session("ClientCnnString"))
	Set rsDelete = Server.CreateObject("ADODB.Recordset")
	rsDelete.CursorLocation = 3 
	Set rsDelete = cnnDelete.Execute(SQLDelete)
	If not rsDelete.Eof Then
		Do
			Activity = "The  Manufacturer for Equipment with Serial # " & rsDelete("SerialNumber") & " was changed from ''" &  ManufacturerToDeleteName & "'' to ''" &  ManufacturerToReplaceWithName & "'' to allow for the deletion of ''" &  ManufacturerToDeleteName & "''"
			CreateAuditLogEntry GetTerm("Equipment") & "  Manufacturer Changed For Equipment Record",GetTerm("Equipment") & "  Manufacturer Changed For Equipment Record","Major",0,Activity
			rsDelete.movenext
		Loop Until rsDelete.Eof
	End If
	rsDelete.Close

	'--------------------------------------------------------------------------------------------
	'Now replace all BRAND records with a new Manufacturer rec id
	'--------------------------------------------------------------------------------------------
	
	SQLDelete = "UPDATE EQ_Brands Set ManufacIntRecID = " & ManufacturerIntRecIDToReplaceWith & " WHERE ManufacIntRecID = " & ManufacturerIntRecIDToReplace
	rsDelete.CursorLocation = 3 
	Set rsDelete = cnnDelete.Execute(SQLDelete)
	
	'--------------------------------------------------------------------------------------------
	'Now replace all MODEL records with a new Manufacturer rec id
	'--------------------------------------------------------------------------------------------
	
	SQLDelete = "UPDATE EQ_Models Set ManufacIntRecID = " & ManufacturerIntRecIDToReplaceWith & " WHERE ManufacIntRecID = " & ManufacturerIntRecIDToReplace
	rsDelete.CursorLocation = 3 
	Set rsDelete = cnnDelete.Execute(SQLDelete)
	
	
	'--------------------------------------------------------------------------------------------
	'Now replace all EQUIPMENT records with a new Manufacturer rec id
	'--------------------------------------------------------------------------------------------
	
	SQLDelete = "UPDATE EQ_Equipment Set ManufacIntRecID = " & ManufacturerIntRecIDToReplaceWith & " WHERE ManufacIntRecID = " & ManufacturerIntRecIDToReplace
	rsDelete.CursorLocation = 3 
	Set rsDelete = cnnDelete.Execute(SQLDelete)
	
	'--------------------------------------------------------------------------------------------
	'Now delete the Manufacturer
	'--------------------------------------------------------------------------------------------
	
	SQLDelete = "DELETE FROM EQ_Manufacturers WHERE InternalRecordIdentifier = " & ManufacturerIntRecIDToReplace
	rsDelete.CursorLocation = 3 
	Set rsDelete = cnnDelete.Execute(SQLDelete)
	
	Description = "The " & GetTerm("Equipment") & "  Manufacturer named " & ManufacturerToDeleteName & " was deleted by " & GetUserDisplayNameByUserNo(Session("UserNo")) 
	CreateAuditLogEntry GetTerm("Equipment") & " Manufacturer Deleted",GetTerm("Equipment") & " Manufacturer Deleted","Major",0,Description
	
Else

'******************************************************************************
'User wants to delete Manufacturer AND all their associated equipment
'******************************************************************************

	'--------------------------------------------------------------------------------------------
	'Delete the  Manufacturer from the Manufacturers table
	'--------------------------------------------------------------------------------------------
	
	SQLDelete = "DELETE FROM EQ_ Manufacturers WHERE InternalRecordIdentifier = " & ManufacturerIntRecIDToReplace
	Set cnnDelete = Server.CreateObject("ADODB.Connection")
	cnnDelete.open (Session("ClientCnnString"))
	Set rsDelete = Server.CreateObject("ADODB.Recordset")
	rsDelete.CursorLocation = 3 
	Set rsDelete = cnnDelete.Execute(SQLDelete)
	
	Description = "The " & GetTerm("Equipment") & "  Manufacturer named " & ManufacturerToDeleteName & " was deleted by " & GetUserDisplayNameByUserNo(Session("UserNo")) 
	CreateAuditLogEntry GetTerm("Equipment") & "  Manufacturer Deleted",GetTerm("Equipment") & " Manufacturer Deleted","Major",0,Description


	'--------------------------------------------------------------------------------------------
	'Delete the equipment associated with this Manufacturer from the equipment table
	'--------------------------------------------------------------------------------------------
	
	SQLDelete = "DELETE FROM EQ_Equipment WHERE ManufacIntRecID = " & ManufacturerIntRecIDToReplace
	rsDelete.CursorLocation = 3 
	Set rsDelete = cnnDelete.Execute(SQLDelete)
	
	Description = "The " & GetTerm("Equipment") & " table records associated with the Manufacturer, " &  ManufacturerToDeleteName & ", were deleted by " & GetUserDisplayNameByUserNo(Session("UserNo")) 
	CreateAuditLogEntry GetTerm("Equipment") & "  Equipment records deleted by manufacturer deletion",GetTerm("Equipment") & " Equipment records deleted by manufacturer deletion","Major",0,Description

	
End If

	
set rsDelete = Nothing
cnnDelete.Close
set cnnDelete = Nothing



Response.Redirect ("main.asp")
%>