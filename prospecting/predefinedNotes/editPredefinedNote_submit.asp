<!--#include file="../../inc/subsandfuncs.asp"-->
<!--#include file="../../inc/InSightFuncs.asp"-->

<%
InternalRecordIdentifier = Request.Form("txtInternalRecordIdentifier")

'*******************************************************************
'Lookup the record as it exists now so we can fillin the audit trail

SQL = "SELECT * FROM PR_PredefinedNotes where InternalRecordIdentifier = " & InternalRecordIdentifier
	
Set cnn8 = Server.CreateObject("ADODB.Connection")
cnn8.open (Session("ClientCnnString"))
Set rs = Server.CreateObject("ADODB.Recordset")
rs.CursorLocation = 3 
Set rs = cnn8.Execute(SQL)
	
If not rs.EOF Then
	Orig_PredefinedNote = rs("PredefinedNote")
End If

set rs = Nothing
cnn8.close
set cnn8 = Nothing
'***********************************************************************
'End Lookup the record as it exists now so we can fillin the audit trail
'***********************************************************************

PredefinedNote = Request.Form("txtPredefinedNote")

SQL = "UPDATE PR_PredefinedNotes SET "
SQL = SQL &  "PredefinedNote = '" & PredefinedNote & "' "
SQL = SQL &  " WHERE InternalRecordIdentifier = " & InternalRecordIdentifier


Set cnn8 = Server.CreateObject("ADODB.Connection")
cnn8.open (Session("ClientCnnString"))

Set rs8 = Server.CreateObject("ADODB.Recordset")
rs8.CursorLocation = 3 
Set rs8 = cnn8.Execute(SQL)
set rs8 = Nothing


Description = ""
If Orig_PredefinedNote  <> PredefinedNote  Then
	Description = Description & GetTerm("Prospecting") & " predefined note changed from " & Orig_PredefinedNote & " to " & PredefinedNote
End If

CreateAuditLogEntry GetTerm("Prospecting") & " predefined note edited",GetTerm("Prospecting") & " predefined note edited","Minor",0,Description

Response.Redirect("main.asp")

%>















