<!--#include file="../../../inc/subsandfuncs.asp"-->
<!--#include file="../../../inc/InSightFuncs.asp"-->

<%
InternalRecordIdentifier = Request.Form("txtInternalRecordIdentifier")

'*******************************************************************
'Lookup the record as it exists now so we can fillin the audit trail

SQL = "SELECT * FROM PR_ContactTitles where InternalRecordIdentifier = " & InternalRecordIdentifier
	
Set cnn8 = Server.CreateObject("ADODB.Connection")
cnn8.open (Session("ClientCnnString"))
Set rs = Server.CreateObject("ADODB.Recordset")
rs.CursorLocation = 3 
Set rs = cnn8.Execute(SQL)
	
If not rs.EOF Then
	Orig_ContactTitle = rs("ContactTitle")
End If

set rs = Nothing
cnn8.close
set cnn8 = Nothing
'***********************************************************************
'End Lookup the record as it exists now so we can fillin the audit trail
'***********************************************************************

ContactTitle = Request.Form("txtContactTitle")

SQL = "UPDATE PR_ContactTitles SET "
SQL = SQL &  "ContactTitle = '" & ContactTitle & "' "
SQL = SQL &  " WHERE InternalRecordIdentifier = " & InternalRecordIdentifier


Set cnn8 = Server.CreateObject("ADODB.Connection")
cnn8.open (Session("ClientCnnString"))

Set rs8 = Server.CreateObject("ADODB.Recordset")
rs8.CursorLocation = 3 
Set rs8 = cnn8.Execute(SQL)
set rs8 = Nothing


Description = ""
If Orig_ContactTitle  <> ContactTitle  Then
	Description = Description & GetTerm("Accounts Receivable") & " contact title changed from " & Orig_ContactTitle & " to " & ContactTitle
End If

CreateAuditLogEntry GetTerm("Accounts Receivable") & " contact title edited",GetTerm("Accounts Receivable") & " contact title edited","Minor",0,Description

Response.Redirect("main.asp")

%>















