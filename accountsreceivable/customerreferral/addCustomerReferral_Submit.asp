<!--#include file="../../inc/subsandfuncs.asp"-->
<!--#include file="../../inc/InSightFuncs.asp"-->
<%

ReferralName = Request.Form("txtReferralName")
RefDescription = Request.Form("txtCustDescription")
RefDescription2 = Request.Form("txtCustDescription2")

SQL = "INSERT INTO AR_CustomerReferral (ReferralName, Description, Description2)"
SQL = SQL &  " VALUES (" 
SQL = SQL & "'"  & ReferralName & "',"
SQL = SQL & "'"  & RefDescription & "',"
SQL = SQL & "'"  & RefDescription2 & "')"


'Response.Write("<br>" & SQL & "<br>")

Set cnncust = Server.CreateObject("ADODB.Connection")
cnncust.open (Session("ClientCnnString"))

Set rscust = Server.CreateObject("ADODB.Recordset")
rscust.CursorLocation = 3 

Set rscust = cnncust.Execute(SQL)
set rscust = Nothing

Description = GetUserDisplayNameByUserNo(Session("Userno")) & " added the accounts receivable module customer referral: " & ReferralName 
CreateAuditLogEntry "Accounts Receivable module" & " customer referral added","Accounts Receivable module customer referral added","Minor",0,Description

Response.Redirect("main.asp")

%>















