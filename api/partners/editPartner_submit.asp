<!--#include file="../../inc/subsandfuncs.asp"-->
<!--#include file="../../inc/InSightFuncs.asp"-->

<%
InternalRecordIdentifier = Request.Form("txtInternalRecordIdentifier")
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





'*******************************************************************
'Lookup the record as it exists now so we can fillin the audit trail

SQL = "SELECT * FROM IC_Partners where InternalRecordIdentifier = " & InternalRecordIdentifier
	
Set cnn8 = Server.CreateObject("ADODB.Connection")
cnn8.open (Session("ClientCnnString"))
Set rs = Server.CreateObject("ADODB.Recordset")
rs.CursorLocation = 3 
Set rs = cnn8.Execute(SQL)
	
If not rs.EOF Then
	Orig_partnerAPIKey = rs("partnerAPIKey")
	Orig_partnerCompanyName = rs("partnerCompanyName")
	Orig_partnerPrimaryContactName = rs("partnerPrimaryContactName")
	Orig_partnerPrimaryContactEmail = rs("partnerPrimaryContactEmail")
	Orig_partnerTechnicalContactName = rs("partnerTechnicalContactName")
	Orig_partnerTechnicalContactEmail = rs("partnerTechnicalContactEmail")
	Orig_partnerAddress = rs("partnerAddress")
	Orig_partnerAddress2 = rs("partnerAddress2")
	Orig_partnerCity = rs("partnerCity")
	Orig_partnerState = rs("partnerState")
	Orig_partnerZip = rs("partnerZip")
	Orig_partnerPhone = rs("partnerPhone")
	Orig_partnerFax = rs("partnerFax")
	Orig_partnerUnmappedTaxableSKU = rs("partnerUnmappedTaxableSKU")
	Orig_partnerUnmappedTaxableUM = rs("partnerUnmappedTaxableUM")
	Orig_partnerUnmappedTaxablePassOriginalSKU = rs("partnerUnmappedTaxablePassOriginalSKU")
	Orig_partnerUnmappedNonTaxableSKU = rs("partnerUnmappedNonTaxableSKU")
	Orig_partnerUnmappedNonTaxableUM = rs("partnerUnmappedNonTaxableUM")
	Orig_partnerUnmappedNonTaxablePassOriginalSKU = rs("partnerUnmappedNonTaxablePassOriginalSKU")
	Orig_partnerUnmappedCustomerID = rs("partnerUnmappedCustomerID")
	Orig_partnerUnmappedPassOriginalCustomerID = rs("partnerUnmappedPassOriginalCustomerID")
	Orig_partnerRejectsBlankProdDescs = rs("partnerRejectsBlankProdDescs")
	Orig_partnerRejectsBlankProdUOMS= rs("partnerRejectsBlankProdUOMS")
	
	If Orig_partnerRejectsBlankProdDescs = true Then Orig_partnerRejectsBlankProdDescs = 1
	If Orig_partnerRejectsBlankProdUOMS= true Then Orig_partnerRejectsBlankProdUOMS= 1
	If Orig_partnerUnmappedTaxablePassOriginalSKU = true Then Orig_partnerUnmappedTaxablePassOriginalSKU = 1
	If Orig_partnerUnmappedNonTaxablePassOriginalSKU = true Then Orig_partnerUnmappedNonTaxablePassOriginalSKU = 1
	If Orig_partnerUnmappedPassOriginalCustomerID = true Then Orig_partnerUnmappedPassOriginalCustomerID = 1

	If Orig_partnerRejectsBlankProdDescs = false Then Orig_partnerRejectsBlankProdDescs = 0
	If Orig_partnerRejectsBlankProdUOMS= false Then Orig_partnerRejectsBlankProdUOMS= 0
	If Orig_partnerUnmappedTaxablePassOriginalSKU = false Then Orig_partnerUnmappedTaxablePassOriginalSKU = 0
	If Orig_partnerUnmappedNonTaxablePassOriginalSKU = false Then Orig_partnerUnmappedNonTaxablePassOriginalSKU = 0
	If Orig_partnerUnmappedPassOriginalCustomerID = false Then Orig_partnerUnmappedPassOriginalCustomerID = 0
	
End If

set rs = Nothing
cnn8.close
set cnn8 = Nothing
'***********************************************************************
'End Lookup the record as it exists now so we can fillin the audit trail
'***********************************************************************

SQL = "UPDATE IC_Partners SET "
SQL = SQL &  "partnerAPIKey = '" & partnerAPIKey & "',partnerCompanyName = '" & partnerCompanyName & "', "
SQL = SQL &  "partnerPrimaryContactName = '" & partnerPrimaryContactName & "',partnerPrimaryContactEmail = '" & partnerPrimaryContactEmail & "', "
SQL = SQL &  "partnerTechnicalContactName = '" & partnerTechnicalContactName & "',partnerTechnicalContactEmail = '" & partnerTechnicalContactEmail & "', "
SQL = SQL &  "partnerAddress = '" & partnerAddress & "',partnerAddress2 = '" & partnerAddress2 & "', "
SQL = SQL &  "partnerCity = '" & partnerCity & "',partnerState = '" & partnerState & "', partnerFax = '" & partnerFax & "', "
SQL = SQL &  "partnerUnmappedTaxableSKU = '" & partnerUnmappedTaxableSKU & "',partnerUnmappedTaxableUM = '" & partnerUnmappedTaxableUM & "',"
SQL = SQL &  "partnerUnmappedTaxablePassOriginalSKU = " & partnerUnmappedTaxablePassOriginalSKU & ", "
SQL = SQL &  "partnerUnmappedNonTaxableSKU = '" & partnerUnmappedNonTaxableSKU & "',partnerUnmappedNonTaxableUM = '" & partnerUnmappedNonTaxableUM & "', "
SQL = SQL &  "partnerUnmappedNonTaxablePassOriginalSKU = " & partnerUnmappedNonTaxablePassOriginalSKU & ", "
SQL = SQL &  "partnerUnmappedCustomerID = '" & partnerUnmappedCustomerID & "',partnerUnmappedPassOriginalCustomerID = " & partnerUnmappedPassOriginalCustomerID & ", "
SQL = SQL &  "partnerRejectsBlankProdDescs = " & partnerRejectsBlankProdDescs & ", partnerRejectsBlankProdUOMS= " & partnerRejectsBlankProdUOMS & " "
SQL = SQL &  " WHERE InternalRecordIdentifier = " & InternalRecordIdentifier


Set cnn8 = Server.CreateObject("ADODB.Connection")
cnn8.open (Session("ClientCnnString"))

Set rs8 = Server.CreateObject("ADODB.Recordset")
rs8.CursorLocation = 3 
Set rs8 = cnn8.Execute(SQL)
set rs8 = Nothing



Description = ""

If Orig_partnerAPIKey <> partnerAPIKey Then
	Description = GetTerm("Inventory Control") & " Partner, " & partnerCompanyName & ", Partner API Key changed from " & Orig_partnerAPIKey & " to " & partnerAPIKey
	CreateAuditLogEntry GetTerm("Inventory Control") & " Partner Edited",GetTerm("Inventory Control") & " Partner Edited","Minor",0,Description
End If

If Orig_partnerCompanyName <> partnerCompanyName Then
	Description = GetTerm("Inventory Control") & " Partner Company Name changed from " & Orig_partnerCompanyName & " to " & partnerCompanyName 
	CreateAuditLogEntry GetTerm("Inventory Control") & " Partner Edited",GetTerm("Inventory Control") & " Partner Edited","Minor",0,Description
End If

If Orig_partnerPrimaryContactName <> partnerPrimaryContactName Then
	Description = GetTerm("Inventory Control") & " Partner, " & partnerCompanyName & ", Partner Primary Contact Name changed from " & Orig_partnerPrimaryContactName & " to " & partnerPrimaryContactName 
	CreateAuditLogEntry GetTerm("Inventory Control") & " Partner Edited",GetTerm("Inventory Control") & " Partner Edited","Minor",0,Description
End If

If Orig_partnerPrimaryContactEmail <> partnerPrimaryContactEmail Then
	Description = GetTerm("Inventory Control") & " Partner, " & partnerCompanyName & ", Partner Primary Contact Email changed from " & Orig_partnerPrimaryContactEmail & " to " & partnerPrimaryContactEmail 
	CreateAuditLogEntry GetTerm("Inventory Control") & " Partner Edited",GetTerm("Inventory Control") & " Partner Edited","Minor",0,Description
End If

If Orig_partnerTechnicalContactName <> partnerTechnicalContactName Then
	Description = GetTerm("Inventory Control") & " Partner, " & partnerCompanyName & ", Partner Technical Contact Name changed from " & Orig_partnerTechnicalContactName & " to " & partnerTechnicalContactName 
	CreateAuditLogEntry GetTerm("Inventory Control") & " Partner Edited",GetTerm("Inventory Control") & " Partner Edited","Minor",0,Description
End If

If Orig_partnerTechnicalContactEmail <> partnerTechnicalContactEmail Then
	Description = GetTerm("Inventory Control") & " Partner, " & partnerCompanyName & ", Partner Technical Contact Email changed from " & Orig_partnerTechnicalContactEmail & " to " & partnerTechnicalContactEmail 
	CreateAuditLogEntry GetTerm("Inventory Control") & " Partner Edited",GetTerm("Inventory Control") & " Partner Edited","Minor",0,Description
End If

If Orig_partnerAddress <> partnerAddress Then
	Description = GetTerm("Inventory Control") & " Partner, " & partnerCompanyName & ", Partner Address 1 changed from " & Orig_partnerAddress & " to " & partnerAddress 
	CreateAuditLogEntry GetTerm("Inventory Control") & " Partner Edited",GetTerm("Inventory Control") & " Partner Edited","Minor",0,Description
End If

If Orig_partnerAddress2 <> partnerAddress2 Then
	Description = GetTerm("Inventory Control") & " Partner, " & partnerCompanyName & ", Partner Address 2 changed from " & Orig_partnerAddress2 & " to " & partnerAddress2 
	CreateAuditLogEntry GetTerm("Inventory Control") & " Partner Edited",GetTerm("Inventory Control") & " Partner Edited","Minor",0,Description
End If

If Orig_partnerCity <> partnerCity Then
	Description = GetTerm("Inventory Control") & " Partner, " & partnerCompanyName & ", Partner City changed from " & Orig_partnerCity & " to " & partnerCity 
	CreateAuditLogEntry GetTerm("Inventory Control") & " Partner Edited",GetTerm("Inventory Control") & " Partner Edited","Minor",0,Description
End If

If Orig_partnerState <> partnerState Then
	Description = GetTerm("Inventory Control") & " Partner, " & partnerCompanyName & ", Partner State changed from " & Orig_partnerState & " to " & partnerState 
	CreateAuditLogEntry GetTerm("Inventory Control") & " Partner Edited",GetTerm("Inventory Control") & " Partner Edited","Minor",0,Description
End If

If Orig_partnerZip <> partnerZip Then
	Description = GetTerm("Inventory Control") & " Partner, " & partnerCompanyName & ", Partner Zip Code changed from " & Orig_partnerZip & " to " & partnerZip 
	CreateAuditLogEntry GetTerm("Inventory Control") & " Partner Edited",GetTerm("Inventory Control") & " Partner Edited","Minor",0,Description
End If

If Orig_partnerPhone <> partnerPhone Then
	Description = GetTerm("Inventory Control") & " Partner, " & partnerCompanyName & ", Partner Phone Number changed from " & Orig_partnerPhone & " to " & partnerPhone 
	CreateAuditLogEntry GetTerm("Inventory Control") & " Partner Edited",GetTerm("Inventory Control") & " Partner Edited","Minor",0,Description
End If

If Orig_partnerFax <> partnerFax Then
	Description = GetTerm("Inventory Control") & " Partner, " & partnerCompanyName & ", Partner Fax Number changed from " & Orig_partnerFax & " to " & partnerFax 
	CreateAuditLogEntry GetTerm("Inventory Control") & " Partner Edited",GetTerm("Inventory Control") & " Partner Edited","Minor",0,Description
End If



If Orig_partnerUnmappedTaxablePassOriginalSKU <> partnerUnmappedTaxablePassOriginalSKU Then

	If Orig_partnerUnmappedTaxablePassOriginalSKU = 1 Then 
	
		Description = GetTerm("Inventory Control") & " Partner, " & partnerCompanyName & ", unmapped taxable product changed from pass through original SKU to map to product code, <strong>" & partnerUnmappedTaxableSKU & "</strong>, and "
		Description = Description & " a UM of <strong>" & partnerUnmappedTaxableUM & "</strong>."
		CreateAuditLogEntry GetTerm("Inventory Control") & " Partner Edited",GetTerm("Inventory Control") & " Partner Edited","Minor",0,Description
		
	ElseIf Orig_partnerUnmappedTaxablePassOriginalSKU = 0 Then
	
		Description = GetTerm("Inventory Control") & " Partner, " & partnerCompanyName & ", unmapped taxable product changed from map to product code, <strong>" & Orig_partnerUnmappedTaxableSKU & "</strong>, and "
		Description = Description & " a UM of <strong>" & Orig_partnerUnmappedTaxableUM & "</strong> to pass through original SKU."
		CreateAuditLogEntry GetTerm("Inventory Control") & " Partner Edited",GetTerm("Inventory Control") & " Partner Edited","Minor",0,Description
		
	End If
	
End If



If Orig_partnerUnmappedNonTaxablePassOriginalSKU <> partnerUnmappedNonTaxablePassOriginalSKU Then

	If Orig_partnerUnmappedNonTaxablePassOriginalSKU = 1 Then 
	
		Description = GetTerm("Inventory Control") & " Partner, " & partnerCompanyName & ", unmapped non taxable product changed from pass through original SKU to map to product code, <strong>" & partnerUnmappedNonTaxableSKU & "</strong>, and "
		Description = Description & " a UM of <strong>" & partnerUnmappedNonTaxableUM & "</strong>."
		CreateAuditLogEntry GetTerm("Inventory Control") & " Partner Edited",GetTerm("Inventory Control") & " Partner Edited","Minor",0,Description
		
	ElseIf Orig_partnerUnmappedNonTaxablePassOriginalSKU = 0 Then
	
		Description = GetTerm("Inventory Control") & " Partner, " & partnerCompanyName & ", unmapped non taxable product changed from map to product code, <strong>" & Orig_partnerUnmappedNonTaxableSKU & "</strong>, and "
		Description = Description & " a UM of <strong>" & Orig_partnerUnmappedNonTaxableUM & "</strong> to pass through original SKU."
		CreateAuditLogEntry GetTerm("Inventory Control") & " Partner Edited",GetTerm("Inventory Control") & " Partner Edited","Minor",0,Description
		
	End If
	
End If



If Orig_partnerUnmappedPassOriginalCustomerID <> partnerUnmappedPassOriginalCustomerID Then

	If Orig_partnerUnmappedPassOriginalCustomerID = 1 Then 
	
		Description = GetTerm("Inventory Control") & " Partner, " & partnerCompanyName & ", unmapped customer changed from pass through original customer number to map to product customer number, <strong>" & partnerUnmappedCustomerID & "</strong>."
		CreateAuditLogEntry GetTerm("Inventory Control") & " Partner Edited",GetTerm("Inventory Control") & " Partner Edited","Minor",0,Description
		
	ElseIf Orig_partnerUnmappedPassOriginalCustomerID = 0 Then
	
		Description = GetTerm("Inventory Control") & " Partner, " & partnerCompanyName & ", unmapped customer changed changed from map to product customer number, <strong>" & Orig_partnerUnmappedCustomerID & "</strong>,"
		Description = Description & " to pass through original customer number."
		CreateAuditLogEntry GetTerm("Inventory Control") & " Partner Edited",GetTerm("Inventory Control") & " Partner Edited","Minor",0,Description
		
	End If
	
End If



If Orig_partnerUnmappedTaxablePassOriginalSKU = partnerUnmappedTaxablePassOriginalSKU Then

	If Orig_partnerUnmappedTaxableSKU <> partnerUnmappedTaxableSKU Then
	
		Description = GetTerm("Inventory Control") & " Partner, " & partnerCompanyName & ", unmapped taxable product changed from map to product code, <strong>" & Orig_partnerUnmappedTaxableSKU & "</strong>, and "
		Description = Description & " a UM of <strong>" & Orig_partnerUnmappedTaxableUM & "</strong> to map to product code "
		Description = Description & " <strong>" & partnerUnmappedTaxableSKU & "</strong>, and "
		Description = Description & " a UM of <strong>" & partnerUnmappedTaxableUM & "</strong>."
		CreateAuditLogEntry GetTerm("Inventory Control") & " Partner Edited",GetTerm("Inventory Control") & " Partner Edited","Minor",0,Description
		
	End If
	
End If


If Orig_partnerUnmappedNonTaxablePassOriginalSKU = partnerUnmappedNonTaxablePassOriginalSKU Then

	If Orig_partnerUnmappedNonTaxableSKU <> partnerUnmappedNonTaxableSKU Then
	
		Description = GetTerm("Inventory Control") & " Partner, " & partnerCompanyName & ", unmapped non taxable product changed from map to product code, <strong>" & Orig_partnerUnmappedNonTaxableSKU & "</strong>, and "
		Description = Description & " a UM of <strong>" & Orig_partnerUnmappedNonTaxableUM  & "</strong> to map to product code "
		Description = Description & " <strong>" & partnerUnmappedNonTaxableSKU & "</strong>, and "
		Description = Description & " a UM of <strong>" & partnerUnmappedNonTaxableUM & "</strong>."
		CreateAuditLogEntry GetTerm("Inventory Control") & " Partner Edited",GetTerm("Inventory Control") & " Partner Edited","Minor",0,Description
		
	End If
End If


If Orig_partnerUnmappedPassOriginalCustomerID = partnerUnmappedPassOriginalCustomerID Then

	If Orig_partnerUnmappedCustomerID <> partnerUnmappedCustomerID Then
	
		Description = GetTerm("Inventory Control") & " Partner, " & partnerCompanyName & ", unmapped non taxable product changed from map to product customer number, <strong>" & Orig_partnerUnmappedCustomerID & "</strong>, "
		Description = Description & " to map to product customer number, "
		Description = Description & " <strong>" & partnerUnmappedCustomerID & "</strong>."
		CreateAuditLogEntry GetTerm("Inventory Control") & " Partner Edited",GetTerm("Inventory Control") & " Partner Edited","Minor",0,Description
		
	End If
End If



If Orig_partnerRejectsBlankProdDescs <> partnerRejectsBlankProdDescs Then

	If Orig_partnerRejectsBlankProdDescs = 1 Then 
	
		Description = GetTerm("Inventory Control") & " Partner, " & partnerCompanyName & ", changed API rejects products with blank descriptions from <strong>TRUE</strong> to <strong>FALSE</strong>."
		CreateAuditLogEntry GetTerm("Inventory Control") & " Partner Edited",GetTerm("Inventory Control") & " Partner Edited","Minor",0,Description
		
	ElseIf Orig_partnerRejectsBlankProdDescs = 0 Then
	
		Description = GetTerm("Inventory Control") & " Partner, " & partnerCompanyName & ", changed API rejects products with blank descriptions from <strong>FALSE</strong> to <strong>TRUE</strong>."
		CreateAuditLogEntry GetTerm("Inventory Control") & " Partner Edited",GetTerm("Inventory Control") & " Partner Edited","Minor",0,Description
		
	End If
	
End If

If Orig_partnerRejectsBlankProdUOMS <> partnerRejectsBlankProdUOMS Then

	If Orig_partnerRejectsBlankProdUOMS = 1 Then 
	
		Description = GetTerm("Inventory Control") & " Partner, " & partnerCompanyName & ", changed API rejects products with blank UOMS from <strong>TRUE</strong> to <strong>FALSE</strong>."
		CreateAuditLogEntry GetTerm("Inventory Control") & " Partner Edited",GetTerm("Inventory Control") & " Partner Edited","Minor",0,Description
		
	ElseIf Orig_partnerRejectsBlankProdUOMS = 0 Then
	
		Description = GetTerm("Inventory Control") & " Partner, " & partnerCompanyName & ", changed API rejects products with blank UOMS from <strong>FALSE</strong> to <strong>TRUE</strong>."
		CreateAuditLogEntry GetTerm("Inventory Control") & " Partner Edited",GetTerm("Inventory Control") & " Partner Edited","Minor",0,Description
		
	End If
	
End If


Response.Redirect("main.asp")

%>