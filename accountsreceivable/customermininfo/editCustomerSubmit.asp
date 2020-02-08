<!--#include file="../../inc/SubsAndFuncs.asp"-->
<!--#include file="../../inc/InSightFuncs.asp"-->

<%

AccountNumber = Request.Form("txtAccountNumber")
CompanyName = Request.Form("txtCompanyName")

Orig_AccountNumber = Request.Form("txtCustIDOriginal") 
Orig_CompanyName = Request.Form("txtCompanyNameOriginal") 	

'*******************************************************************************************************************
'FIX ANY ENTRIES THAT MAY CONTAIN SINGLE QUOTES FOR SQL INSERT
'*******************************************************************************************************************

AccountNumber = Replace(AccountNumber,"'","''")
CompanyName = Replace(CompanyName,"'","''")
LastPriceChangeDate = Replace(LastPriceChangeDate,"'","''")

'***************************************************************************************************************************************************************
'First update main customer table, AR_Customer
'***************************************************************************************************************************************************************

Set cnnAR_Customer = Server.CreateObject("ADODB.Connection")
cnnAR_Customer.open (Session("ClientCnnString"))
Set rsAR_Customer = Server.CreateObject("ADODB.Recordset")

SQLAR_Customer = "UPDATE AR_Customer SET CustNum = '" & AccountNumber & "', Name = '" & CompanyName & "' WHERE CustNum = '" & Orig_AccountNumber & "'"

rsAR_Customer.CursorLocation = 3 
Set rsAR_Customer = cnnAR_Customer.Execute(SQLAR_Customer)

Response.Write("SQLAR_Customer: " & SQLAR_Customer & "<br><br>")


'***************************************************************************************************************************************************************
'If the Account Number has changed, we need to update AR_CustomerBillTo and AR_CustomerShipTo locations
'***************************************************************************************************************************************************************

Description = GetUserDisplayNameByUserNo(Session("Userno")) & " edited an existing AR customer account, " & AccountNumber & ", with company name " & CompanyName  & " by " & GetUserDisplayNameByUserNo(Session("UserNo"))
CreateAuditLogEntry "Existing customer edited","Existing customer edited","Major",0,Description

If Orig_AccountNumber <> AccountNumber Then

	SQLAR_Customer = "UPDATE AR_CustomerBillTo SET CustNum = '" & AccountNumber & "' WHERE CustNum = '" & Orig_AccountNumber & "'"
	rsAR_Customer.CursorLocation = 3 
	Set rsAR_Customer = cnnAR_Customer.Execute(SQLAR_Customer)
	
	Response.Write("SQLAR_Customer2: " & SQLAR_Customer & "<br><Br>")

	SQLAR_Customer = "UPDATE AR_CustomerShipTo SET CustNum = '" & AccountNumber & "' WHERE CustNum = '" & Orig_AccountNumber & "'"
	rsAR_Customer.CursorLocation = 3 
	Set rsAR_Customer = cnnAR_Customer.Execute(SQLAR_Customer)
	
	Response.Write("SQLAR_Customer3: " & SQLAR_Customer & "<br><Br>")

End If


'***************************************************************************************************************************************************************
'If the Company Name has changed, we need to update AR_CustomerBillTo (default billing location only)
'***************************************************************************************************************************************************************
	
If Orig_CompanyName <> CompanyName Then

	SQLAR_Customer = "UPDATE AR_CustomerBillTo SET BillName = '" & CompanyName & "' WHERE CustNum = '" & AccountNumber & "' AND DefaultBillTo = 1"
	rsAR_Customer.CursorLocation = 3 
	Set rsAR_Customer = cnnAR_Customer.Execute(SQLAR_Customer)

	Response.Write("SQLAR_Customer4: " & SQLAR_Customer & "<br><Br>")
	
	Description =  GetTerm("Accounts Receivable") & " AR customer main company name changed from " & Orig_CompanyName & " to " & CompanyName & " for the customer account " & Orig_AccountNumber & ", " & CompanyName & " by " & GetUserDisplayNameByUserNo(Session("UserNo")) 
	CreateAuditLogEntry GetTerm("Accounts Receivable") & " AR customer main company name change ",GetTerm("Accounts Receivable"),"Minor",0,Description
	
	Description =  GetTerm("Accounts Receivable") & " AR customer bill to location company name changed from " & Orig_CompanyName & " to " & CompanyName & " for customer account " & Orig_AccountNumber & ", " & CompanyName & " by " & GetUserDisplayNameByUserNo(Session("UserNo"))
	CreateAuditLogEntry GetTerm("Accounts Receivable") & " AR customer bill to location company name change ",GetTerm("Accounts Receivable"),"Minor",0,Description
	
End If

Response.Redirect("editViewCustomerDetail.asp?cid=" & AccountNumber)

%>
