<!--#include file="../../inc/SubsAndFuncs.asp"-->
<!--#include file="../../inc/InSightFuncs.asp"-->

<%

txtAccountNumber = Request.Form("txtAccountNumber")
txtCompanyName = Request.Form("txtCompanyName")
txtLastPriceChangeDate = Request.Form("txtLastPriceChangeDate")

txtBillToContactFirstName = Request.Form("txtBillToContactFirstName")
txtBillToContactLastName = Request.Form("txtBillToContactLastName")
txtBillToCompanyName = Request.Form("txtBillToCompanyName")
txtBillToAddressLine1 = Request.Form("txtBillToAddressLine1")
txtBillToAddressLine2 = Request.Form("txtBillToAddressLine2")
txtBillToCity = Request.Form("txtBillToCity")
txtBillToState = Request.Form("txtBillToState")
txtBillToZipCode = Request.Form("txtBillToZipCode")
txtBillToCountry = Request.Form("txtBillToCountry")
txtBillToPhoneNumber = Request.Form("txtBillToPhoneNumber")
txtBillToEmailAddress = Request.Form("txtBillToEmailAddress")

txtShipToContactFirstName = Request.Form("txtShipToContactFirstName")
txtShipToContactLastName = Request.Form("txtShipToContactLastName")
txtShipToCompanyName = Request.Form("txtShipToCompanyName")
txtShipToAddressLine1 = Request.Form("txtShipToAddressLine1")
txtShipToAddressLine2 = Request.Form("txtShipToAddressLine2")
txtShipToCity = Request.Form("txtShipToCity")
txtShipToState = Request.Form("txtShipToState")
txtShipToZipCode = Request.Form("txtShipToZipCode")
txtShipToCountry = Request.Form("txtShipToCountry")
txtShipToPhoneNumber = Request.Form("txtShipToPhoneNumber")
txtShipToEmailAddress = Request.Form("txtShipToEmailAddress")


'*******************************************************************************************************************
'FIX ANY ENTRIES THAT MAY CONTAIN SINGLE QUOTES FOR SQL INSERT
'*******************************************************************************************************************

txtAccountNumber = Replace(txtAccountNumber,"'","''")
txtCompanyName = Replace(txtCompanyName,"'","''")
txtLastPriceChangeDate = Replace(txtLastPriceChangeDate,"'","''")

txtBillToContactFirstName = Replace(txtBillToContactFirstName,"'","''")
txtBillToContactLastName = Replace(txtBillToContactLastName,"'","''")
txtBillToCompanyName = Replace(txtBillToCompanyName,"'","''")
txtBillToAddressLine1 = Replace(txtBillToAddressLine1,"'","''")
txtBillToAddressLine2 = Replace(txtBillToAddressLine2,"'","''")
txtBillToCity = Replace(txtBillToCity,"'","''")
txtBillToEmailAddress = Replace(txtBillToEmailAddress,"'","''")

txtShipToContactFirstName = Replace(txtShipToContactFirstName,"'","''")
txtShipToContactLastName = Replace(txtShipToContactLastName,"'","''")
txtShipToCompanyName = Replace(txtShipToCompanyName,"'","''")
txtShipToAddressLine1 = Replace(txtShipToAddressLine1,"'","''")
txtShipToAddressLine2 = Replace(txtShipToAddressLine2,"'","''")
txtShipToCity = Replace(txtShipToCity,"'","''")
txtShipToEmailAddress = Replace(txtShipToEmailAddress,"'","''")


'***************************************************************************************************************************************************************
'First make entry into main customer table, AR_Customer
'***************************************************************************************************************************************************************

ContactName = txtBillToContactFirstName & " " & txtBillToContactLastName 
CityStateZip = txtBillToCity & ", " & txtBillToState & " " & txtBillToZipCode

SQLAR_Customer = "INSERT INTO AR_Customer (CustNum, Name, AcctStatus, ContactFirstName, ContactLastName, Contact, Addr1, Addr2, City, [State], Zip, Country, Phone, CityStateZip, LastPriceChangeDate)"
SQLAR_Customer = SQLAR_Customer &  " VALUES (" 
SQLAR_Customer = SQLAR_Customer & "'" & txtAccountNumber & "','"  & txtCompanyName & "','A','" & txtBillToContactFirstName & "','" & txtBillToContactLastName & "','" & ContactName & "', "
SQLAR_Customer = SQLAR_Customer & "'" & txtBillToAddressLine1 & "','"  & txtBillToAddressLine2 & "','"  & txtBillToCity & "','"  & txtBillToState & "' ,"
SQLAR_Customer = SQLAR_Customer & "'" & txtBillToCountry & "','" & txtBillToZipCode & "','" & txtBillToPhoneNumber & "','" & CityStateZip & "','" & txtLastPriceChangeDate & "')"

'Response.write(SQLAR_Customer)

Set cnnAR_Customer = Server.CreateObject("ADODB.Connection")
cnnAR_Customer.open (Session("ClientCnnString"))

Set rsAR_Customer = Server.CreateObject("ADODB.Recordset")
rsAR_Customer.CursorLocation = 3 
Set rsAR_Customer = cnnAR_Customer.Execute(SQLAR_Customer)


Description = GetUserDisplayNameByUserNo(Session("Userno")) & " added a new customer account, " & txtAccountNumber & ", with company name " & txtCompanyName 
CreateAuditLogEntry "New customer added","New customer added","Major",0,Description



'***************************************************************************************************************************************************************
'Now Insert Billing Info as Default Bill To in AR_CustomerBillTo
'***************************************************************************************************************************************************************

BillToContactName = txtBillToContactFirstName & " " & txtBillToContactLastName 
SQLAR_Customer = "INSERT INTO AR_CustomerBillTo (CustNum, BillName, ContactFirstName, ContactLastName, Contact, Addr1, Addr2, City, [State], Zip, Country, Phone, Email, DefaultBillTo)"
SQLAR_Customer = SQLAR_Customer &  " VALUES (" 
SQLAR_Customer = SQLAR_Customer & "'" & txtAccountNumber & "','" & txtBillToCompanyName & "','" & txtBillToContactFirstName & "','" & txtBillToContactLastName & "','" & BillToContactName & "', "
SQLAR_Customer = SQLAR_Customer & "'" & txtBillToAddressLine1 & "','"  & txtBillToAddressLine2 & "','"  & txtBillToCity & "','"  & txtBillToState & "' ,"
SQLAR_Customer = SQLAR_Customer & "'" & txtBillToCountry & "','" & txtBillToZipCode & "','" & txtBillToPhoneNumber & "','" & txtBillToEmailAddress & "',1)"

 
rsAR_Customer.CursorLocation = 3 
Set rsAR_Customer = cnnAR_Customer.Execute(SQLAR_Customer)




'***************************************************************************************************************************************************************
'Now Insert Shiping Info as Default Ship To in AR_CustomerShipTo
'***************************************************************************************************************************************************************

ShipToContactName = txtShipToContactFirstName & " " & txtShipToContactLastName

SQLAR_Customer = "INSERT INTO AR_CustomerShipTo (CustNum, ShipName, ContactFirstName, ContactLastName, Contact, Addr1, Addr2, City, [State], Zip, Country, Phone, Email, DefaultShipTo)"
SQLAR_Customer = SQLAR_Customer &  " VALUES (" 
SQLAR_Customer = SQLAR_Customer & "'" & txtAccountNumber & "','"  & txtShipToCompanyName & "','" & txtShipToContactFirstName & "','" & txtShipToContactLastName & "','" & ShipToContactName & "', "
SQLAR_Customer = SQLAR_Customer & "'" & txtShipToAddressLine1 & "','"  & txtShipToAddressLine2 & "','"  & txtShipToCity & "','"  & txtShipToState & "' ,"
SQLAR_Customer = SQLAR_Customer & "'" & txtShipToCountry & "','" & txtShipToZipCode & "','" & txtShipToPhoneNumber & "','" & txtShipToEmailAddress & "',1)"

 
rsAR_Customer.CursorLocation = 3 
Set rsAR_Customer = cnnAR_Customer.Execute(SQLAR_Customer)




'Response.Redirect("editCustomer.asp?i=" & ProspectIntRecID)
Response.Redirect("main.asp")
%>
