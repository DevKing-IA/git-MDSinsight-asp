<!--#include file="../../inc/subsandfuncs.asp"-->
<!--#include file="../../inc/InSightFuncs.asp"-->
<!--#include file="../../inc/InSightFuncs_Equipment.asp"-->

<%

Model = Request.Form("txtModel")
BrandIntRecID = Request.Form("selBrandIntRecID")
GroupIntRecID = Request.Form("selGroupIntRecID")
ClassIntRecID = Request.Form("selClassIntRecID")
DefaultRentalPrice = Request.Form("txtDefaultRentalPrice")
DefaultCost = Request.Form("txtDefaultCost")
ReplacementCost = Request.Form("txtReplacementCost")
BackendSystemCode = Request.Form("txtBackendSystemCode")
InsightAssetTagPrefix = Request.Form("txtInsightAssetTagPrefix")

If DefaultRentalPrice = "" Then
	DefaultRentalPrice = 0
End If
If DefaultCost = "" Then
	DefaultCost = 0
End If
If ReplacementCost = "" Then
	ReplacementCost = 0
End If

ManufacturerIntRecID = GetManufacturerIntRecIDByBrandIntRecID(BrandIntRecID)

SQL = "INSERT INTO EQ_Models (Model, BrandIntRecID, ManufacIntRecID, GroupIntRecID, "
SQL = SQL & " ClassIntRecID, DefaultRentalPrice, DefaultCostPrice, ReplacementCost, RecordSource, BackendSystemCode, InsightAssetTagPrefix)"
SQL = SQL &  " VALUES (" 
SQL = SQL & "'"  & Model & "'," & BrandIntRecID & "," & ManufacturerIntRecID & "," & GroupIntRecID & ", "
SQL = SQL & ClassIntRecID & "," & DefaultRentalPrice & "," & DefaultCost & "," & ReplacementCost & ",'Insight','" & BackendSystemCode & "','" & InsightAssetTagPrefix & "')"

Set cnn8 = Server.CreateObject("ADODB.Connection")
cnn8.open (Session("ClientCnnString"))

Set rs8 = Server.CreateObject("ADODB.Recordset")
rs8.CursorLocation = 3 
Response.Write(SQL)
Set rs8 = cnn8.Execute(SQL)
set rs8 = Nothing

Description = GetUserDisplayNameByUserNo(Session("Userno")) & " added the " & GetTerm("Equipment") & " model: " & Model  & " and Asset Tag Prefix, " & InsightAssetTagPrefix

CreateAuditLogEntry GetTerm("Equipment") & " Equipment Model Added",GetTerm("Equipment") & " Equipment Model Added","Minor",0,Description

Response.Redirect("main.asp")

%>















