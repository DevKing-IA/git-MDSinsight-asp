<!--#include file="../../inc/InSightFuncs.asp"-->
<!--#include file="../../inc/InSightFuncs_Users.asp"-->
<%
PartnerInternalRecordIdentifier = Request.QueryString("x")
SKUInternalRecordIdentifier = Request.QueryString("i")
SKUtoDelete = Request.QueryString("p")
UMToDelete = Request.QueryString("u")
CategoryID = Request.QueryString("c")

If SKUInternalRecordIdentifier <> "" Then

	SQLDelete = "Delete FROM IC_ProductMapping WHERE InternalRecordIdentifier = " & SKUInternalRecordIdentifier
	
	Set cnnDelete = Server.CreateObject("ADODB.Connection")
	cnnDelete.open (Session("ClientCnnString"))
	Set rsDelete = Server.CreateObject("ADODB.Recordset")
	rsDelete.CursorLocation = 3 
	Set rsDelete = cnnDelete.Execute(SQLDelete)
	
	
	Description = "The " & GetTerm("Inventory") & " Partner SKU <strong>" & SKUtoDelete & "</strong> with unit <strong>" & UMToDelete & "</strong>, in category <strong>" & CategoryID & "</strong>, was deleted by " & GetUserDisplayNameByUserNo(Session("UserNo")) 
	CreateAuditLogEntry GetTerm("Inventory Control") & " Partner Equivalent SKU Deleted",GetTerm("Inventory Control") & " Partner Equivalent SKU Deleted","Major",0,Description
	
	set rsDelete = Nothing
	cnnDelete.Close
	set cnnDelete = Nothing
	
End If

Response.Redirect ("editPartnerSKUCategoryToEdit.asp?i=" & PartnerInternalRecordIdentifier & "&c=" & CategoryID)
%>