<!--#include file="../../inc/subsandfuncs.asp"-->
<!--#include file="../../inc/InSightFuncs.asp"-->

<%
InternalRecordIdentifier = Request.Form("txtInternalRecordIdentifier")

'*******************************************************************
'Lookup the record as it exists now so we can fillin the audit trail

SQL = "SELECT * FROM PR_NoteTypes where InternalRecordIdentifier = " & InternalRecordIdentifier
	
Set cnn8 = Server.CreateObject("ADODB.Connection")
cnn8.open (Session("ClientCnnString"))
Set rs = Server.CreateObject("ADODB.Recordset")
rs.CursorLocation = 3 
Set rs = cnn8.Execute(SQL)
	
If not rs.EOF Then
	Orig_NoteType = rs("NoteType")
End If

set rs = Nothing
cnn8.close
set cnn8 = Nothing
'***********************************************************************
'End Lookup the record as it exists now so we can fillin the audit trail
'***********************************************************************

NoteType = Request.Form("txtNoteType")

SQL = "UPDATE PR_NoteTypes SET "
SQL = SQL &  "NoteType = '" & NoteType & "' "
SQL = SQL &  " WHERE InternalRecordIdentifier = " & InternalRecordIdentifier


Set cnn8 = Server.CreateObject("ADODB.Connection")
cnn8.open (Session("ClientCnnString"))

Set rs8 = Server.CreateObject("ADODB.Recordset")
rs8.CursorLocation = 3 
Set rs8 = cnn8.Execute(SQL)
set rs8 = Nothing


Description = ""
If Orig_NoteType  <> NoteType  Then
	Description = Description & GetTerm("Prospecting") & " note type changed from " & Orig_NoteType & " to " & NoteType
End If

CreateAuditLogEntry GetTerm("Prospecting") & " note type edited",GetTerm("Prospecting") & " note type edited","Minor",0,Description

Response.Redirect("main.asp")

%>















