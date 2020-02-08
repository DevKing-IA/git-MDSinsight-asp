<!--#include file="../../inc/subsandfuncs.asp"-->
<!--#include file="../../inc/InSightFuncs.asp"-->
<%

PredefinedNote = Request.Form("txtPredefinedNote")

SQL = "INSERT INTO PR_PredefinedNotes (PredefinedNote)"
SQL = SQL &  " VALUES (" 
SQL = SQL & "'"  & PredefinedNote & "')"

Set cnn8 = Server.CreateObject("ADODB.Connection")
cnn8.open (Session("ClientCnnString"))

Set rs8 = Server.CreateObject("ADODB.Recordset")
rs8.CursorLocation = 3 
Set rs8 = cnn8.Execute(SQL)
set rs8 = Nothing

Description = GetUserDisplayNameByUserNo(Session("Userno")) & " added the " & GetTerm("Prospecting") & " predefined note: " & PredefinedNote
CreateAuditLogEntry GetTerm("Prospecting") & " predefined note added",GetTerm("Prospecting") & " predefined note added","Minor",0,Description

Response.Redirect("main.asp")

%>















