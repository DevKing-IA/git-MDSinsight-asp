<!--#include file="../../inc/subsandfuncs.asp"-->
<!--#include file="../../inc/InSightFuncs.asp"-->

<%
InternalRecordIdentifier = Request.Form("txtInternalRecordIdentifier")

'*******************************************************************
'Lookup the record as it exists now so we can fillin the audit trail

SQL = "SELECT * FROM AR_Terms where InternalRecordIdentifier = " & InternalRecordIdentifier
	
Set cnncust = Server.CreateObject("ADODB.Connection")
cnncust.open (Session("ClientCnnString"))
Set rscust = Server.CreateObject("ADODB.Recordset")
rscust.CursorLocation = 3 
Set rscust = cnncust.Execute(SQL)
	
If not rscust.EOF Then	
	Orig_TermDescription = rscust("Description")	
End If

'***********************************************************************
'End Lookup the record as it exists now so we can fillin the audit trail
'***********************************************************************

TermDescription = Request.Form("txtTermDescription")

firstTermsPercent = Request.Form("txtfirstTermsPercent")
firstTermsPeriod = Request.Form("txtfirstTermsPeriod")
secondTermsPeriod = Request.Form("txtsecondTermsPeriod")
TermsType = Request.Form("txtTermsType")
CreditCardBill = Request.Form("txtCreditCardBill")


SQL = "UPDATE AR_Terms SET "
SQL = SQL &  "Description = '" & TermDescription & "', "
SQL = SQL &  "firstTermsPercent = '" & firstTermsPercent & "', "
SQL = SQL &  "firstTermsPeriod = '" & firstTermsPeriod & "', "
SQL = SQL &  "secondTermsPeriod = '" & secondTermsPeriod & "', "
SQL = SQL &  "TermsType = '" & TermsType & "', "
SQL = SQL &  "CreditCardBill = '" & CreditCardBill & "' "
SQL = SQL &  " WHERE InternalRecordIdentifier = " & InternalRecordIdentifier

'Response.Write("<br>" & SQL & "<br>")

Set rscust = cnncust.Execute(SQL)
set rscust = Nothing


Description = ""
If Orig_TermDescription  <> TermDescription  Then
	Description = Description & "Accounts Receivable module Term changed from " & Orig_TermDescription  & " to " & TermDescription  
End If

CreateAuditLogEntry "Accounts Receivable module term edited","Accounts Receivable module term edited","Minor",0,Description

Response.Redirect("main.asp")

%>















