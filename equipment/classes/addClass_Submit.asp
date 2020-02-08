<!--#include file="../../inc/subsandfuncs.asp"-->
<!--#include file="../../inc/InSightFuncs.asp"-->
<%

EquipmentClass = Request.Form("txtClass")
EquipmentClass = Replace(EquipmentClass, "'", "''")
InsightAssetTagPrefix = Request.Form("txtInsightAssetTagPrefix")

SQL = "INSERT INTO EQ_Classes (Class, InsightAssetTagPrefix, RecordSource)"
SQL = SQL &  " VALUES (" 
SQL = SQL & "'" & EquipmentClass & "','" & InsightAssetTagPrefix & "','Insight')"

Set cnn8 = Server.CreateObject("ADODB.Connection")
cnn8.open (Session("ClientCnnString"))

Set rs8 = Server.CreateObject("ADODB.Recordset")
rs8.CursorLocation = 3 
Set rs8 = cnn8.Execute(SQL)
set rs8 = Nothing

Description = GetUserDisplayNameByUserNo(Session("Userno")) & " added the " & GetTerm("Equipment") & " class: " & EquipmentClass & " and Asset Tag Prefix, " & InsightAssetTagPrefix

CreateAuditLogEntry GetTerm("Equipment") & " Class Added",GetTerm("Equipment") & " Class Added","Minor",0,Description

Response.Redirect("main.asp")

%>















