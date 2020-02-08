<!--#include file="../../inc/subsandfuncs.asp"-->
<!--#include file="../../inc/InSightFuncs.asp"-->
<%

partnerAPIKey = Request.Form("txtPartnerAPIKey")
partnerCompanyName = Request.Form("txtPartnerCompanyName")
partnerPrimaryContactName = Request.Form("txtPrimaryContactName")
partnerPrimaryContactEmail = Request.Form("txtPrimaryContactEmailAddress")
partnerTechnicalContactName = Request.Form("txtTechnicalContactName")
partnerTechnicalContactEmail = Request.Form("txtTechnicalContactEmailAddress")
partnerAddress = Request.Form("txtAddressLine1")
partnerAddress2 = Request.Form("txtAddressLine2")
partnerCity = Request.Form("txtCity")
partnerState = Request.Form("txtState")
partnerZip = Request.Form("txtZipCode")
partnerPhone = Request.Form("txtPhoneNumber")
partnerFax = Request.Form("txtFaxNumber")

partnerRejectsBlankProdDescs = Request.Form("chkRejectsBlankProdDescs")
If (partnerRejectsBlankProdDescs <> "" AND partnerRejectsBlankProdDescs = "on") Then partnerRejectsBlankProdDescs = 1 Else partnerRejectsBlankProdDescs = 0

partnerRejectsBlankProdUOMS = Request.Form("chkRejectsBlankProdUOMS")
If (partnerRejectsBlankProdUOMS <> "" AND partnerRejectsBlankProdUOMS = "on") Then partnerRejectsBlankProdUOMS = 1 Else partnerRejectsBlankProdUOMS = 0


partnerUnmappedTaxablePassOriginalSKU_RADIO = Request.Form("optMappedOrPassedTaxable")

If partnerUnmappedTaxablePassOriginalSKU_RADIO = "DefinedCode" Then
	partnerUnmappedTaxablePassOriginalSKU = 0
	partnerUnmappedTaxableSKU = Request.Form("txtUnmappedTaxableSKUToPass")
	partnerUnmappedTaxableUM = Request.Form("txtUnmappedTaxableUM")
Else
	partnerUnmappedTaxablePassOriginalSKU = 1
	partnerUnmappedTaxableSKU = ""
	partnerUnmappedTaxableUM = ""
End If


partnerUnmappedNonTaxablePassOriginalSKU_RADIO = Request.Form("optMappedOrPassedNonTaxable")

If partnerUnmappedNonTaxablePassOriginalSKU_RADIO = "DefinedCode" Then
	partnerUnmappedNonTaxablePassOriginalSKU = 0
	partnerUnmappedNonTaxableSKU = Request.Form("txtUnmappedNonTaxableSKUToPass")
	partnerUnmappedNonTaxableUM = Request.Form("txtUnmappedNonTaxableUM")
Else
	partnerUnmappedNonTaxablePassOriginalSKU = 1
	partnerUnmappedNonTaxableSKU = ""
	partnerUnmappedNonTaxableUM = ""
End If


partnerUnmappedPassOriginalCustomerID_RADIO = Request.Form("optMappedOrPassedCustAccount")

If partnerUnmappedPassOriginalCustomerID_RADIO = "DefinedAccount" Then
	partnerUnmappedPassOriginalCustomerID = 0
	partnerUnmappedCustomerID = Request.Form("txtUnmappedCustomerIDToPass")
Else
	partnerUnmappedPassOriginalCustomerID = 1
	partnerUnmappedCustomerID = ""
End If


SQL = "INSERT INTO IC_Partners (partnerAPIKey, partnerCompanyName, partnerPrimaryContactName, partnerPrimaryContactEmail, "
SQL = SQL & " partnerTechnicalContactName, partnerTechnicalContactEmail, partnerAddress, partnerAddress2, partnerCity, partnerState, partnerZip, partnerPhone, partnerFax, "
SQL = SQL & " partnerUnmappedTaxableSKU, partnerUnmappedTaxableUM, partnerUnmappedTaxablePassOriginalSKU, partnerUnmappedNonTaxableSKU, "
SQL = SQL & " partnerUnmappedNonTaxableUM, partnerUnmappedNonTaxablePassOriginalSKU, partnerUnmappedCustomerID, partnerUnmappedPassOriginalCustomerID, partnerRejectsBlankProdDescs, partnerRejectsBlankProdUOMS)"
SQL = SQL &  " VALUES (" 
SQL = SQL & "'" & partnerAPIKey & "','" & partnerCompanyName & "','" & partnerPrimaryContactName & "','" & partnerPrimaryContactEmail & "','" & partnerTechnicalContactName & "','" & partnerTechnicalContactEmail & "',"
SQL = SQL & "'" & partnerAddress & "','" & partnerAddress2 & "','" & partnerCity & "','" & partnerState & "','" & partnerZip & "','" & partnerPhone & "','" & partnerFax & "', "
SQL = SQL & "'" & partnerUnmappedTaxableSKU & "','" & partnerUnmappedTaxableUM & "'," & partnerUnmappedTaxablePassOriginalSKU & ", "
SQL = SQL & "'" & partnerUnmappedNonTaxableSKU & "','" & partnerUnmappedNonTaxableUM & "'," & partnerUnmappedNonTaxablePassOriginalSKU & ", "
SQL = SQL & "'" & partnerUnmappedCustomerID & "'," & partnerUnmappedPassOriginalCustomerID & "," & partnerRejectsBlankProdDescs & "," & partnerRejectsBlankProdUOMS& ")"

Set cnn8 = Server.CreateObject("ADODB.Connection")
cnn8.open (Session("ClientCnnString"))

Set rs8 = Server.CreateObject("ADODB.Recordset")
rs8.CursorLocation = 3 
Set rs8 = cnn8.Execute(SQL)
set rs8 = Nothing

Description = GetUserDisplayNameByUserNo(Session("Userno")) & " added the " & GetTerm("Inventory Control") & " Partner : " & partnerCompanyName & " with API Key " & partnerAPIKey
CreateAuditLogEntry GetTerm("Inventory Control") & " Partner Added",GetTerm("Inventory Control") & " Partner Added","Minor",0,Description

Response.Redirect("main.asp")

%>















