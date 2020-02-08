<!--#include file="../../inc/subsandfuncs.asp"-->
<!--#include file="../../inc/InSightFuncs.asp"-->
<%

Condition = Request.Form("txtCondition")
Condition = Replace(Condition, "'", "''")

ConditionDescription = Request.Form("txtConditionDescription")
ConditionDescription = Replace(ConditionDescription, "'", "''")

SQL = "INSERT INTO EQ_Condition (Condition,Description,RecordSource)"
SQL = SQL &  " VALUES (" 
SQL = SQL & "'"  & Condition & "','"  & ConditionDescription & "','Insight')"

Set cnn8 = Server.CreateObject("ADODB.Connection")
cnn8.open (Session("ClientCnnString"))

Set rs8 = Server.CreateObject("ADODB.Recordset")
rs8.CursorLocation = 3 
Response.Write(SQL)
Set rs8 = cnn8.Execute(SQL)
set rs8 = Nothing

Description = GetUserDisplayNameByUserNo(Session("Userno")) & " added the " & GetTerm("Equipment") & " condition: " & Condition 
CreateAuditLogEntry GetTerm("Equipment") & " Condition Added",GetTerm("Equipment") & " Condition Added","Minor",0,Description

Response.Redirect("main.asp")

%>















