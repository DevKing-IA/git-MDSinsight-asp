<!--#include file="../inc/subsandfuncs.asp"-->
<!--#include file="../inc/InSightFuncs.asp"-->
<%

ContactTitle = Request.Form("txtContactTitle")

'check if fields are not empty
If ContactTitle<>"" Then
	ContactTitle = Hacker_Filter2(ContactTitle)
End If

SQL = "INSERT INTO PR_ContactTitles (ContactTitle)"
SQL = SQL &  " VALUES (" 
SQL = SQL & "'"  & ContactTitle &  "')"

Set cnn8 = Server.CreateObject("ADODB.Connection")
cnn8.open (Session("ClientCnnString"))

Set rs8 = Server.CreateObject("ADODB.Recordset")
rs8.CursorLocation = 3 
Response.Write(SQL)
Set rs8 = cnn8.Execute(SQL)
set rs8 = Nothing

Description = GetUserDisplayNameByUserNo(Session("Userno")) & " added the " & GetTerm("Prospecting") & " contact title: " & ContactTitle 
CreateAuditLogEntry GetTerm("Prospecting") & " Contact Title Added",GetTerm("Prospecting") & " Contact Title Added","Minor",0,Description

'Response.Redirect("viewProspectDetail.asp?i=" & InternalRecordIdentifier)
%>

