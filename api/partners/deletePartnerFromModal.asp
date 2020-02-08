<!--#include file="../../inc/InsightFuncs_InventoryControl.asp"-->
<!--#include file="../../inc/InSightFuncs.asp"-->
<!--#include file="../../inc/InSightFuncs_Users.asp"-->
<%

PartnerIntRecIDToReplace = Request.Form("txtPartnerIntRecIDToReplace")
PartnerToDeleteName = GetPartnerNameByIntRecID(PartnerIntRecIDToReplace)
PartnerToReplaceWith = Request.Form("selDeletePartnerFromModal")

If PartnerIntRecIDToReplace <> "" AND PartnerToReplaceWith <> "DELETE_PARTNER_AND_SKUS" Then

	PartnerToReplaceWithName = GetPartnerNameByIntRecID(PartnerToReplaceWith)

	'**************************************************************************************
	'User wants to delete partner and select a replacement partner for all defined SKUS
	'**************************************************************************************
	
	'--------------------------------------------------------------------------------------------
	'update the audit trail with all the skus that are about to be assigned to a new partner
	'--------------------------------------------------------------------------------------------

	SQLDelete = "SELECT * FROM IC_ProductMapping WHERE partnerIntRecID = " & PartnerIntRecIDToReplace
	Set cnnDelete = Server.CreateObject("ADODB.Connection")
	cnnDelete.open (Session("ClientCnnString"))
	Set rsDelete = Server.CreateObject("ADODB.Recordset")
	rsDelete.CursorLocation = 3 
	Set rsDelete = cnnDelete.Execute(SQLDelete)
	If not rsDelete.Eof Then
		Do
			Activity = "The partner for SKU " & rsDelete("SKU") & " was changed from ''" & PartnerToDeleteName & "'' to ''" & PartnerToReplaceWithName & "'' to allow for the deletion of ''" & PartnerToDeleteName & "''"
			CreateAuditLogEntry GetTerm("Inventory Control") & " partner equivalent sku partner ID changed",GetTerm("Inventory Control") & " partner equivalent sku partner ID changed","Major",0,Activity
			rsDelete.movenext
		Loop Until rsDelete.Eof
	End If
	rsDelete.Close
	
	'--------------------------------------------------------------------------------------------
	'Now replace all sku records with a new partner rec id
	'--------------------------------------------------------------------------------------------
	
	SQLDelete = "UPDATE IC_ProductMapping Set partnerIntRecID = " & PartnerToReplaceWith & " WHERE partnerIntRecID = " & PartnerIntRecIDToReplace
	rsDelete.CursorLocation = 3 
	Set rsDelete = cnnDelete.Execute(SQLDelete)
	
	'--------------------------------------------------------------------------------------------
	'Now delete the partner
	'--------------------------------------------------------------------------------------------
	
	SQLDelete = "DELETE FROM IC_Partners WHERE InternalRecordIdentifier = " & PartnerIntRecIDToReplace
	rsDelete.CursorLocation = 3 
	Set rsDelete = cnnDelete.Execute(SQLDelete)
	
	Description = "The " & GetTerm("Inventory Control") & " partner named " & PartnerToDeleteName & " was deleted by " & GetUserDisplayNameByUserNo(Session("UserNo")) 
	CreateAuditLogEntry GetTerm("Inventory Control") & " partner deleted",GetTerm("Inventory Control") & " partner deleted","Major",0,Description
	
Else

'******************************************************************************
'User wants to delete partner AND all their equivalent product SKUs
'******************************************************************************

	'--------------------------------------------------------------------------------------------
	'Delete the partner from the partners table
	'--------------------------------------------------------------------------------------------
	
	SQLDelete = "DELETE FROM IC_Partners WHERE InternalRecordIdentifier = " & PartnerIntRecIDToReplace
	Set cnnDelete = Server.CreateObject("ADODB.Connection")
	cnnDelete.open (Session("ClientCnnString"))
	Set rsDelete = Server.CreateObject("ADODB.Recordset")
	rsDelete.CursorLocation = 3 
	Set rsDelete = cnnDelete.Execute(SQLDelete)
	
	Description = "The " & GetTerm("Inventory Control") & " partner named " & PartnerToDeleteName & " was deleted by " & GetUserDisplayNameByUserNo(Session("UserNo")) 
	CreateAuditLogEntry GetTerm("Inventory Control") & " partner deleted",GetTerm("Inventory Control") & " partner deleted","Major",0,Description


	'--------------------------------------------------------------------------------------------
	'Delete the skus associated with this partner from the product equivalent SKUs table
	'--------------------------------------------------------------------------------------------
	
	SQLDelete = "DELETE FROM IC_ProductMapping WHERE partnerIntRecID= " & PartnerIntRecIDToReplace
	rsDelete.CursorLocation = 3 
	Set rsDelete = cnnDelete.Execute(SQLDelete)
	
	Description = "The " & GetTerm("Inventory Control") & " SKUs associated with the partner, " & PartnerToDeleteName & ", were deleted by " & GetUserDisplayNameByUserNo(Session("UserNo")) 
	CreateAuditLogEntry GetTerm("Inventory Control") & " partner equivalent sku set deleted",GetTerm("Inventory Control") & " partner equivalent sku set deleted","Major",0,Description

	
End If

	
set rsDelete = Nothing
cnnDelete.Close
set cnnDelete = Nothing



Response.Redirect ("main.asp")
%>