<!--#include file="../../inc/subsandfuncs.asp"-->
<!--#include file="../../inc/InSightFuncs.asp"-->
<%

CustomerClassDesc = Request.Form("txtCustomerClassDesc")
CustomerClassCode = Request.Form("txtCustomerClassCode")

SQL = "INSERT INTO AR_CustomerClass (ClassDescription, ClassCode)"
SQL = SQL &  " VALUES (" 
SQL = SQL & "'" & CustomerClassDesc & "','" & CustomerClassCode & "')"

Set cnn8 = Server.CreateObject("ADODB.Connection")
cnn8.open (Session("ClientCnnString"))

Set rs8 = Server.CreateObject("ADODB.Recordset")
rs8.CursorLocation = 3 
Set rs8 = cnn8.Execute(SQL)
set rs8 = Nothing

Description = GetUserDisplayNameByUserNo(Session("Userno")) & " added the " & GetTerm("Accounts Receivable") & " customer class code : " & ClassCode & " - " & ClassDescription
CreateAuditLogEntry GetTerm("Accounts Receivable") & " customer class code added",GetTerm("Accounts Receivable") & " customer class code added","Minor",0,Description

Response.Redirect("main.asp")

%>















