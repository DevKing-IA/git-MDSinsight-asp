<!--#include file="../../inc/InsightFuncs_Equipment.asp"-->
<!--#include file="../../inc/InSightFuncs.asp"-->
<!--#include file="../../inc/InSightFuncs_Users.asp"-->
<%
BrandIntRecIDToReplace = Request.Form("txtBrandIntRecIDToReplace")
BrandToDeleteName = GetBrandNameByIntRecID(BrandIntRecIDToReplace)
BrandIntRecIDToReplaceWith = Request.Form("seldeleteBrandFromBrand")
BrandToReplaceWithName = GetBrandNameByIntRecID(BrandIntRecIDToReplaceWith) 

If  BrandIntRecIDToReplace <> "" AND  BrandIntRecIDToReplaceWith <> "DELETE_BRAND_AND_EQUIPMENT" Then

	'---------------------------------------------------------------------------------------------------
	'update the audit trail with all the equipment records that are about to be assigned to a new Brand
	'---------------------------------------------------------------------------------------------------

	SQLDelete = "SELECT * FROM EQ_Equipment WHERE BrandIntRecID = " & BrandIntRecIDToReplace
	Set cnnDelete = Server.CreateObject("ADODB.Connection")
	cnnDelete.open (Session("ClientCnnString"))
	Set rsDelete = Server.CreateObject("ADODB.Recordset")
	rsDelete.CursorLocation = 3 
	Set rsDelete = cnnDelete.Execute(SQLDelete)
	If not rsDelete.Eof Then
		Do
			Activity = "The  Brand for Equipment with Serial # " & rsDelete(" SerialNumber") & " was changed from ''" &  BrandToDeleteName & "'' to ''" &  BrandToReplaceWithName & "'' to allow for the deletion of ''" &  BrandToDeleteName & "''"
			CreateAuditLogEntry GetTerm("Equipment") & " Brand Changed For Equipment Record",GetTerm("Equipment") & " Brand Changed For Equipment Record","Major",0,Activity
			rsDelete.movenext
		Loop Until rsDelete.Eof
	End If
	rsDelete.Close

	'--------------------------------------------------------------------------------------------
	'Now replace all Brand records with a new Brand Name
	'--------------------------------------------------------------------------------------------
	
	SQLDelete = "UPDATE EQ_Brands Set Brand = " &  BrandToReplaceWithName & " WHERE InternalRecordIdentifier = " &  BrandIntRecIDToReplace
	rsDelete.CursorLocation = 3 
	Set rsDelete = cnnDelete.Execute(SQLDelete)	
	
	'--------------------------------------------------------------------------------------------
	'Now replace all EQUIPMENT records with a new Brand Int RecID
	'--------------------------------------------------------------------------------------------
	
	SQLDelete = "UPDATE EQ_Equipment Set BrandIntRecID = " &  BrandIntRecIDToReplaceWith & " WHERE  BrandIntRecID = " &  BrandIntRecIDToReplace
	rsDelete.CursorLocation = 3 
	Set rsDelete = cnnDelete.Execute(SQLDelete)
	
	'--------------------------------------------------------------------------------------------
	'Now delete the Brand from the EQ_Brands tables
	'--------------------------------------------------------------------------------------------
	
	'Now Do the deletion
	
	SQLDelete = "Delete FROM EQ_Brands WHERE InternalRecordIdentifier = " & BrandIntRecIDToReplace
	rsDelete.CursorLocation = 3 
	Set rsDelete = cnnDelete.Execute(SQLDelete)
	
	Description = "The " & GetTerm("Equipment") & " Brand named " & BrandToDeleteName & " was deleted by " & GetUserDisplayNameByUserNo(Session("UserNo")) 
	CreateAuditLogEntry GetTerm("Equipment") & " Brand Deleted",GetTerm("Equipment") & " Brand Deleted","Major",0,Description
	
	set rsDelete = Nothing
	cnnDelete.Close
	set cnnDelete = Nothing
	
Else

'******************************************************************************
'User wants to delete Brand AND all its associated equipment
'******************************************************************************
	'--------------------------------------------------------------------------------------------
	'Delete the Brand from the EQ_Brands tables
	'--------------------------------------------------------------------------------------------
	
	'Now Do the deletion
	
	SQLDelete = "Delete FROM EQ_Brands WHERE InternalRecordIdentifier = " & BrandIntRecIDToReplace
	rsDelete.CursorLocation = 3 
	Set rsDelete = cnnDelete.Execute(SQLDelete)
	
	Description = "The " & GetTerm("Equipment") & " Brand named " & BrandToDeleteName & " was deleted by " & GetUserDisplayNameByUserNo(Session("UserNo")) 
	CreateAuditLogEntry GetTerm("Equipment ") & " Brand Deleted",GetTerm("Equipment") & " Brand Deleted","Major",0,Description
	
	set rsDelete = Nothing
	cnnDelete.Close
	set cnnDelete = Nothing

	'--------------------------------------------------------------------------------------------
	'Delete the equipment associated with this Manufacturer from the equipment table
	'--------------------------------------------------------------------------------------------
	
	SQLDelete = "DELETE FROM EQ_Equipment WHERE BrandIntRecID = " & BrandIntRecIDToReplace
	rsDelete.CursorLocation = 3 
	Set rsDelete = cnnDelete.Execute(SQLDelete)
	
	Description = "The " & GetTerm("Equipment") & " table records associated with the Brand, " &  BrandToDeleteName & ", were deleted by " & GetUserDisplayNameByUserNo(Session("UserNo")) 
	CreateAuditLogEntry GetTerm("Equipment") & "  Equipment records deleted by Brand deletion",GetTerm("Equipment") & " Equipment records deleted by Brand deletion","Major",0,Description
	
End If

Response.Redirect ("main.asp")
%>