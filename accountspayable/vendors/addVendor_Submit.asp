<!--#include file="../../inc/subsandfuncs.asp"-->
<!--#include file="../../inc/InSightFuncs.asp"-->
<%

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

SQL = "INSERT INTO AP_Vendor (vendorAPIKey, vendorCompanyName, vendorPrimaryContactName, vendorPrimaryContactEmail, "
SQL = SQL & " vendorTechnicalContactName, vendorTechnicalContactEmail, vendorAddress, vendorAddress2, vendorCity, vendorState, vendorZip, vendorPhone, vendorFax, Website, AccountNumber, Notes)"
SQL = SQL &  " VALUES (" 
SQL = SQL & "'" & vendorAPIKey & "','" & vendorCompanyName & "','" & vendorPrimaryContactName & "','" & vendorPrimaryContactEmail & "','" & vendorTechnicalContactName & "','" & vendorTechnicalContactEmail & "',"
SQL = SQL & "'" & vendorAddress & "','" & vendorAddress2 & "','" & vendorCity & "','" & vendorState & "', "
SQL = SQL & "'" & vendorZip & "','" & vendorPhone & "','" & vendorFax & "','" & vendorWebsite & "','" & vendorAccountNumber & "','" & vendorNotes & "')"

Set cnn8 = Server.CreateObject("ADODB.Connection")
cnn8.open (Session("ClientCnnString"))

Set rs8 = Server.CreateObject("ADODB.Recordset")
rs8.CursorLocation = 3 
Set rs8 = cnn8.Execute(SQL)
set rs8 = Nothing

Description = GetUserDisplayNameByUserNo(Session("Userno")) & " added the " & GetTerm("Accounts Payable") & " Vendor : " & vendorCompanyName & " with API Key " & vendorAPIKey
CreateAuditLogEntry GetTerm("Accounts Payable") & " vendor Added",GetTerm("Accounts Payable") & " vendor Added","Minor",0,Description

Response.Redirect("main.asp")

%>















