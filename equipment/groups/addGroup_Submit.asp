<!--#include file="../../inc/subsandfuncs.asp"-->
<!--#include file="../../inc/InSightFuncs.asp"-->
<%

EquipmentGroup = Request.Form("txtGroup")
EquipmentGroup = Replace(EquipmentGroup, "'", "''")

SQL = "INSERT INTO EQ_Groups (GroupName, RecordSource)"
SQL = SQL &  " VALUES (" 
SQL = SQL & "'"  & EquipmentGroup & "','Insight')"

Set cnn8 = Server.CreateObject("ADODB.Connection")
cnn8.open (Session("ClientCnnString"))

Set rs8 = Server.CreateObject("ADODB.Recordset")
rs8.CursorLocation = 3 
Set rs8 = cnn8.Execute(SQL)
set rs8 = Nothing

Description = GetUserDisplayNameByUserNo(Session("Userno")) & " added the " & GetTerm("Equipment") & " Group: " & EquipmentGroup
CreateAuditLogEntry GetTerm("Equipment") & " Group Added",GetTerm("Equipment") & " Group Added","Minor",0,Description

Response.Redirect("main.asp")

%>















