<!--#include file="../../inc/subsandfuncs.asp"-->
<!--#include file="../../inc/InSightFuncs.asp"-->
<!--#include file="../../inc/InSightFuncs_Equipment.asp"-->
<%

Brand = Request.Form("txtBrand")
ManufacturerIntRecID = Request.Form("selManufacturerIntRecID")
InsightAssetTagPrefix = Request.Form("txtInsightAssetTagPrefix")

SQL = "INSERT INTO EQ_Brands (ManufacIntRecID,Brand,InsightAssetTagPrefix,RecordSource)"
SQL = SQL &  " VALUES (" 
SQL = SQL & "'"  & ManufacturerIntRecID & "','" & Brand & "','" & InsightAssetTagPrefix & "','Insight')"

Set cnn8 = Server.CreateObject("ADODB.Connection")
cnn8.open (Session("ClientCnnString"))

Set rs8 = Server.CreateObject("ADODB.Recordset")
rs8.CursorLocation = 3 
Response.Write(SQL)
Set rs8 = cnn8.Execute(SQL)
set rs8 = Nothing

Description = GetUserDisplayNameByUserNo(Session("Userno")) & " added the " & GetTerm("Equipment") & " Brand: " & Brand & " with Manufacturer, " & GetManufacturerNameByIntRecID(ManufacturerIntRecID) & " and Asset Tag Prefix, " & InsightAssetTagPrefix
CreateAuditLogEntry GetTerm("Equipment") & " Brand Added",GetTerm("Equipment") & " Brand Added","Minor",0,Description

Response.Redirect("main.asp")

%>















