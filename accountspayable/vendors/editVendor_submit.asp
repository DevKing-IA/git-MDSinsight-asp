<!--#include file="../../inc/subsandfuncs.asp"-->
<!--#include file="../../inc/InSightFuncs.asp"-->

<%
InternalRecordIdentifier = Request.Form("txtInternalRecordIdentifier")
vendorAPIKey = Request.Form("txtvendorAPIKey")
vendorCompanyName = Request.Form("txtvendorCompanyName")
vendorPrimaryContactName = Request.Form("txtPrimaryContactName")
vendorPrimaryContactEmail = Request.Form("txtPrimaryContactEmailAddress")
vendorTechnicalContactName = Request.Form("txtTechnicalContactName")
vendorTechnicalContactEmail = Request.Form("txtTechnicalContactEmailAddress")
vendorAddress = Request.Form("txtAddressLine1")
vendorAddress2 = Request.Form("txtAddressLine2")
vendorCity = Request.Form("txtCity")
vendorState = Request.Form("txtState")
vendorZip = Request.Form("txtZipCode")
vendorPhone = Request.Form("txtPhoneNumber")
vendorFax = Request.Form("txtFaxNumber")

vendorWebsite = Request.Form("txtWebsite")
vendorAccountNumber = Request.Form("txtAccountNumber")
vendorNotes = Request.Form("txtNotes")

'*******************************************************************
'Lookup the record as it exists now so we can fillin the audit trail

SQL = "SELECT * FROM AP_Vendor where InternalRecordIdentifier = " & InternalRecordIdentifier
	
Set cnn8 = Server.CreateObject("ADODB.Connection")
cnn8.open (Session("ClientCnnString"))
Set rs = Server.CreateObject("ADODB.Recordset")
rs.CursorLocation = 3 
Set rs = cnn8.Execute(SQL)
	
If not rs.EOF Then
	Orig_vendorAPIKey = rs("vendorAPIKey")
	Orig_vendorCompanyName = rs("vendorCompanyName")
	Orig_vendorPrimaryContactName = rs("vendorPrimaryContactName")
	Orig_vendorPrimaryContactEmail = rs("vendorPrimaryContactEmail")
	Orig_vendorTechnicalContactName = rs("vendorTechnicalContactName")
	Orig_vendorTechnicalContactEmail = rs("vendorTechnicalContactEmail")
	Orig_vendorAddress = rs("vendorAddress")
	Orig_vendorAddress2 = rs("vendorAddress2")
	Orig_vendorCity = rs("vendorCity")
	Orig_vendorState = rs("vendorState")
	Orig_vendorZip = rs("vendorZip")
	Orig_vendorPhone = rs("vendorPhone")
	Orig_vendorFax = rs("vendorFax")
	Orig_vendorWebsite = rs("Website")	
	Orig_vendorAccountNumber = rs("AccountNumber")	
	Orig_vendorNotes = rs("Notes")	
End If

set rs = Nothing
cnn8.close
set cnn8 = Nothing
'***********************************************************************
'End Lookup the record as it exists now so we can fillin the audit trail
'***********************************************************************

SQL = "UPDATE AP_Vendor SET "
SQL = SQL &  "vendorAPIKey = '" & vendorAPIKey & "',vendorCompanyName = '" & vendorCompanyName & "', "
SQL = SQL &  "vendorPrimaryContactName = '" & vendorPrimaryContactName & "',vendorPrimaryContactEmail = '" & vendorPrimaryContactEmail & "', "
SQL = SQL &  "vendorTechnicalContactName = '" & vendorTechnicalContactName & "',vendorTechnicalContactEmail = '" & vendorTechnicalContactEmail & "', "
SQL = SQL &  "vendorAddress = '" & vendorAddress & "',vendorAddress2 = '" & vendorAddress2 & "', "
SQL = SQL &  "vendorCity = '" & vendorCity & "',vendorState = '" & vendorState & "', vendorFax = '" & vendorFax & "',"
SQL = SQL &  "Website = '" & vendorWebsite & "',AccountNumber = '" & vendorAccountNumber & "', Notes = '" & vendorNotes & "'"
SQL = SQL &  " WHERE InternalRecordIdentifier = " & InternalRecordIdentifier


Set cnn8 = Server.CreateObject("ADODB.Connection")
cnn8.open (Session("ClientCnnString"))

Set rs8 = Server.CreateObject("ADODB.Recordset")
rs8.CursorLocation = 3 
Set rs8 = cnn8.Execute(SQL)
set rs8 = Nothing


Description = ""

If Orig_vendorAPIKey <> vendorAPIKey Then
	Description = GetTerm("Accounts Payable") & " Vendor API Key changed from " & Orig_vendorAPIKey & " to " & vendorAPIKey
	CreateAuditLogEntry GetTerm("Accounts Payable") & " Vendor Edited",GetTerm("Accounts Payable") & " Vendor Edited","Minor",0,Description
End If

If Orig_vendorCompanyName <> vendorCompanyName Then
	Description = GetTerm("Accounts Payable") & " Vendor Company Name changed from " & Orig_vendorCompanyName & " to " & vendorCompanyName 
	CreateAuditLogEntry GetTerm("Accounts Payable") & " Vendor Edited",GetTerm("Accounts Payable") & " Vendor Edited","Minor",0,Description
End If

If Orig_vendorPrimaryContactName <> vendorPrimaryContactName Then
	Description = GetTerm("Accounts Payable") & " Vendor Primary Contact Name changed from " & Orig_vendorPrimaryContactName & " to " & vendorPrimaryContactName 
	CreateAuditLogEntry GetTerm("Accounts Payable") & " Vendor Edited",GetTerm("Accounts Payable") & " Vendor Edited","Minor",0,Description
End If

If Orig_vendorPrimaryContactEmail <> vendorPrimaryContactEmail Then
	Description = GetTerm("Accounts Payable") & " Vendor Primary Contact Email changed from " & Orig_vendorPrimaryContactEmail & " to " & vendorPrimaryContactEmail 
	CreateAuditLogEntry GetTerm("Accounts Payable") & " Vendor Edited",GetTerm("Accounts Payable") & " Vendor Edited","Minor",0,Description
End If

If Orig_vendorTechnicalContactName <> vendorTechnicalContactName Then
	Description = GetTerm("Accounts Payable") & " Vendor Technical Contact Name changed from " & Orig_vendorTechnicalContactName & " to " & vendorTechnicalContactName 
	CreateAuditLogEntry GetTerm("Accounts Payable") & " Vendor Edited",GetTerm("Accounts Payable") & " Vendor Edited","Minor",0,Description
End If

If Orig_vendorTechnicalContactEmail <> vendorTechnicalContactEmail Then
	Description = GetTerm("Accounts Payable") & " Vendor Technical Contact Email changed from " & Orig_vendorTechnicalContactEmail & " to " & vendorTechnicalContactEmail 
	CreateAuditLogEntry GetTerm("Accounts Payable") & " Vendor Edited",GetTerm("Accounts Payable") & " Vendor Edited","Minor",0,Description
End If

If Orig_vendorAddress <> vendorAddress Then
	Description = GetTerm("Accounts Payable") & " Vendor Address 1 changed from " & Orig_vendorAddress & " to " & vendorAddress 
	CreateAuditLogEntry GetTerm("Accounts Payable") & " Vendor Edited",GetTerm("Accounts Payable") & " Vendor Edited","Minor",0,Description
End If

If Orig_vendorAddress2 <> vendorAddress2 Then
	Description = GetTerm("Accounts Payable") & " Vendor Address 2 changed from " & Orig_vendorAddress2 & " to " & vendorAddress2 
	CreateAuditLogEntry GetTerm("Accounts Payable") & " Vendor Edited",GetTerm("Accounts Payable") & " Vendor Edited","Minor",0,Description
End If

If Orig_vendorCity <> vendorCity Then
	Description = GetTerm("Accounts Payable") & " Vendor City changed from " & Orig_vendorCity & " to " & vendorCity 
	CreateAuditLogEntry GetTerm("Accounts Payable") & " Vendor Edited",GetTerm("Accounts Payable") & " Vendor Edited","Minor",0,Description
End If

If Orig_vendorState <> vendorState Then
	Description = GetTerm("Accounts Payable") & " Vendor State changed from " & Orig_vendorState & " to " & vendorState 
	CreateAuditLogEntry GetTerm("Accounts Payable") & " Vendor Edited",GetTerm("Accounts Payable") & " Vendor Edited","Minor",0,Description
End If

If Orig_vendorZip <> vendorZip Then
	Description = GetTerm("Accounts Payable") & " Vendor Zip Code changed from " & Orig_vendorZip & " to " & vendorZip 
	CreateAuditLogEntry GetTerm("Accounts Payable") & " Vendor Edited",GetTerm("Accounts Payable") & " Vendor Edited","Minor",0,Description
End If

If Orig_vendorPhone <> vendorPhone Then
	Description = GetTerm("Accounts Payable") & " Vendor Phone Number changed from " & Orig_vendorPhone & " to " & vendorPhone 
	CreateAuditLogEntry GetTerm("Accounts Payable") & " Vendor Edited",GetTerm("Accounts Payable") & " Vendor Edited","Minor",0,Description
End If

If Orig_vendorFax <> vendorFax Then
	Description = GetTerm("Accounts Payable") & " Vendor Fax Number changed from " & Orig_vendorFax & " to " & vendorFax 
	CreateAuditLogEntry GetTerm("Accounts Payable") & " Vendor Edited",GetTerm("Accounts Payable") & " Vendor Edited","Minor",0,Description
End If

If Orig_vendorWebsite <> vendorWebsite Then
	Description = GetTerm("Accounts Payable") & " Vendor website changed from " & Orig_vendorWebsite & " to " & vendorWebsite
	CreateAuditLogEntry GetTerm("Accounts Payable") & " Vendor Edited",GetTerm("Accounts Payable") & " Vendor Edited","Minor",0,Description
End If

If Orig_vendorAccountNumber <> vendorAccountNumber Then
	Description = GetTerm("Accounts Payable") & " Vendor account number changed from " & Orig_vendorAccountNumber & " to " & vendorAccountNumber
	CreateAuditLogEntry GetTerm("Accounts Payable") & " Vendor Edited",GetTerm("Accounts Payable") & " Vendor Edited","Minor",0,Description
End If

If Orig_vendorNotes <> vendorNotes Then
	Description = GetTerm("Accounts Payable") & " Vendor notes changed from " & Orig_vendorNotes & " to " & vendorNotes
	CreateAuditLogEntry GetTerm("Accounts Payable") & " Vendor Edited",GetTerm("Accounts Payable") & " Vendor Edited","Minor",0,Description
End If




Response.Redirect("main.asp")

%>