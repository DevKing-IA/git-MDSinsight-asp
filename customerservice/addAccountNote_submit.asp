<!--#include file="../inc/subsandfuncs.asp"-->
<!--#include file="../inc/InsightFuncs.asp"-->


<%

AccountNumber = Request.Form("txtAccount")
AccountNote = Request.Form("txtAccountNote")
Sticky = Request.Form("chkSticky")
If Sticky = "on" then Sticky = 1 Else Sticky = 0
ExpirationDate = Request.Form("txtExpirationDate")

'Replace all vbCrLf with <BR>s
AccountNote = Replace(AccountNote , vbCrLf, "<BR>")

SQL = "INSERT INTO tblCustomerNotes (CustNum,Note,UserNo,Sequence,Sticky,ExpirationDate) "
SQL = SQL &  " VALUES (" 
SQL = SQL & "'"  & AccountNumber & "'"
SQL = SQL & ",'"  & AccountNote & "'"
SQL = SQL & ","  & Session("UserNo") & ",0," & Sticky & ", '"
SQL = SQL & ExpirationDate & "')"
	
Set cnn8 = Server.CreateObject("ADODB.Connection")
cnn8.open (Session("ClientCnnString"))

Set rs8 = Server.CreateObject("ADODB.Recordset")
rs8.CursorLocation = 3 
Set rs8 = cnn8.Execute(SQL)
set rs8 = Nothing

Description = ""
Description = Description & "A new note was added to account # "  & AccountNumber
Description = Description & "     The text of the note is as follows: "  & AccountNote
 
CreateAuditLogEntry "Account Note Added","Account Note Added","Minor",0,Description

Response.Redirect("main.asp#home")

%>















