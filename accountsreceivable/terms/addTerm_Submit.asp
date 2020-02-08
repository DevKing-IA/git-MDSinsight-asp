<!--#include file="../../inc/subsandfuncs.asp"-->
<!--#include file="../../inc/InSightFuncs.asp"-->
<%

TermDescription = Request.Form("txtTermDescription")

firstTermsPercent = Request.Form("txtfirstTermsPercent")
firstTermsPeriod = Request.Form("txtfirstTermsPeriod")
secondTermsPeriod = Request.Form("txtsecondTermsPeriod")
TermsType = Request.Form("txtTermsType")
CreditCardBill = Request.Form("txtCreditCardBill")

SQL = "INSERT INTO AR_Terms (Description, firstTermsPercent, firstTermsPeriod, secondTermsPeriod, TermsType, CreditCardBill)"
SQL = SQL &  " VALUES (" 
SQL = SQL & "'"  & TermDescription & "',"
SQL = SQL & "'"  & firstTermsPercent & "',"
SQL = SQL & "'"  & firstTermsPeriod & "',"
SQL = SQL & "'"  & secondTermsPeriod & "',"
SQL = SQL & "'"  & TermsType & "',"
SQL = SQL & "'"  & CreditCardBill & "')"


'Response.Write("<br>" & SQL & "<br>")

Set cnncust = Server.CreateObject("ADODB.Connection")
cnncust.open (Session("ClientCnnString"))

Set rscust = Server.CreateObject("ADODB.Recordset")
rscust.CursorLocation = 3 

Set rscust = cnncust.Execute(SQL)
set rscust = Nothing

Description = GetUserDisplayNameByUserNo(Session("Userno")) & " added the accounts receivable module term: " & TermDescription 
CreateAuditLogEntry "Accounts Receivable module" & " term added","Accounts Receivable module term added","Minor",0,Description

Response.Redirect("main.asp")

%>















