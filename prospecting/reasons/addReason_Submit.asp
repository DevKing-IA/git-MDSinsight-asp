<!--#include file="../../inc/subsandfuncs.asp"-->
<!--#include file="../../inc/InSightFuncs.asp"-->
<%

Reason = Request.Form("txtReason")
ReasonType = Request.Form("selReasonType")

SQL = "INSERT INTO PR_Reasons (Reason, ReasonType)"
SQL = SQL &  " VALUES (" 
SQL = SQL & "'" & Reason & "','" & ReasonType & "')"

Set cnn8 = Server.CreateObject("ADODB.Connection")
cnn8.open (Session("ClientCnnString"))

Set rs8 = Server.CreateObject("ADODB.Recordset")
rs8.CursorLocation = 3 
Set rs8 = cnn8.Execute(SQL)
set rs8 = Nothing

Description = GetUserDisplayNameByUserNo(Session("Userno")) & " added the " & GetTerm("Prospecting") & " reason: " & Reason & " of type: " & ReasonType
CreateAuditLogEntry GetTerm("Prospecting") & " reason added",GetTerm("Prospecting") & " reason added","Minor",0,Description

Response.Redirect("main.asp")

%>















