<!--#include file="../../inc/subsandfuncs.asp"-->
<!--#include file="../../inc/InSightFuncs.asp"-->

<%
InternalRecordIdentifier = Request.Form("txtInternalRecordIdentifier")

'*******************************************************************
'Lookup the record as it exists now so we can fillin the audit trail

SQL = "SELECT * FROM AR_CustomerClass where InternalRecordIdentifier = " & InternalRecordIdentifier
	
Set cnn8 = Server.CreateObject("ADODB.Connection")
cnn8.open (Session("ClientCnnString"))
Set rs = Server.CreateObject("ADODB.Recordset")
rs.CursorLocation = 3 
Set rs = cnn8.Execute(SQL)
	
If not rs.EOF Then
	Orig_ClassCode = rs("ClassCode")
	Orig_ClassDescription = rs("ClassDescription")
End If

set rs = Nothing
cnn8.close
set cnn8 = Nothing
'***********************************************************************
'End Lookup the record as it exists now so we can fillin the audit trail
'***********************************************************************

CustomerClassDesc = Request.Form("txtCustomerClassDesc")
CustomerClassCode = Request.Form("txtCustomerClassCode")

SQL = "UPDATE AR_Customer SET ClassCode = '" & CustomerClassCode & "' WHERE ClassCode = '" & Orig_ClassCode & "'"

Set cnn8 = Server.CreateObject("ADODB.Connection")
cnn8.open (Session("ClientCnnString"))

Set rs8 = Server.CreateObject("ADODB.Recordset")
rs8.CursorLocation = 3 
Set rs8 = cnn8.Execute(SQL)
set rs8 = Nothing


SQL = "UPDATE AR_CustomerClass SET "
SQL = SQL &  "ClassDescription = '" & CustomerClassDesc & "',ClassCode = '" & CustomerClassCode & "' "
SQL = SQL &  " WHERE InternalRecordIdentifier = " & InternalRecordIdentifier

Set cnn8 = Server.CreateObject("ADODB.Connection")
cnn8.open (Session("ClientCnnString"))

Set rs8 = Server.CreateObject("ADODB.Recordset")
rs8.CursorLocation = 3 
Set rs8 = cnn8.Execute(SQL)
set rs8 = Nothing


Description = ""


If Orig_ClassDescription <> CustomerClassDesc AND Orig_ClassCode <> CustomerClassCode Then

	Description = Description & GetTerm("Accounts Receivable") & " customer class description changed from " & Orig_ClassDescription & " to " & CustomerClassDesc & " and the class code changed from " & Orig_ClassCode & " to " & CustomerClassCode 
Else

	If Orig_ClassDescription <> CustomerClassDesc Then
		Description = Description & GetTerm("Accounts Receivable") & " customer class description changed from " & Orig_ClassDescription & " to " & CustomerClassDesc 
	End If
	
	If Orig_ClassCode <> CustomerClassCode Then
		Description = Description & GetTerm("Accounts Receivable") & " customer class code changed from " & Orig_ClassCode & " to " & CustomerClassCode 
	End If
	
End If


CreateAuditLogEntry GetTerm("Accounts Receivable") & " customer class code edited",GetTerm("Accounts Receivable") & " customer class code edited","Minor",0,Description

Response.Redirect("main.asp")

%>















