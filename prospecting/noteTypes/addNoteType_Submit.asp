<!--#include file="../../inc/subsandfuncs.asp"-->
<!--#include file="../../inc/InSightFuncs.asp"-->
<%

NoteType = Request.Form("txtNoteType")

SQL = "INSERT INTO PR_NoteTypes (NoteType)"
SQL = SQL &  " VALUES (" 
SQL = SQL & "'"  & NoteType & "')"

Set cnn8 = Server.CreateObject("ADODB.Connection")
cnn8.open (Session("ClientCnnString"))

Set rs8 = Server.CreateObject("ADODB.Recordset")
rs8.CursorLocation = 3 
Set rs8 = cnn8.Execute(SQL)
set rs8 = Nothing

Description = GetUserDisplayNameByUserNo(Session("Userno")) & " added the " & GetTerm("Prospecting") & " note type: " & NoteType
CreateAuditLogEntry GetTerm("Prospecting") & " note type added",GetTerm("Prospecting") & " note type added","Minor",0,Description

Response.Redirect("main.asp")

%>















