<!--#include file="../../inc/subsandfuncs.asp"-->
<!--#include file="../../inc/InSightFuncs.asp"-->
<!--#include file="../../inc/InsightFuncs_BizIntel.asp"-->
<%

Reason = Request.Form("txtMCSReason")
Reason = Replace(Reason, "'", "''")

SQL = "INSERT INTO BI_MCSReasons (Reason)"
SQL = SQL &  " VALUES (" 
SQL = SQL & "'" & Reason & "')"

Set cnn8 = Server.CreateObject("ADODB.Connection")
cnn8.open (Session("ClientCnnString"))

Set rs8 = Server.CreateObject("ADODB.Recordset")
rs8.CursorLocation = 3 
Set rs8 = cnn8.Execute(SQL)
set rs8 = Nothing

Description = GetUserDisplayNameByUserNo(Session("Userno")) & " added the " & GetTerm("Business Intelligence") & " MCS reason: " & Reason & "."
CreateAuditLogEntry GetTerm("Business Intelligence") & " MCS reason added",GetTerm("Business Intelligence") & " MCS reason added","Minor",0,Description

Response.Redirect("main.asp")

%>















