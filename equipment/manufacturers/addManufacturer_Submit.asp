<!--#include file="../../inc/subsandfuncs.asp"-->
<!--#include file="../../inc/InSightFuncs.asp"-->
<%

manufacturerName = Request.Form("txtManufacturerName")
manufacturerAddress = Request.Form("txtAddressLine1")
manufacturerAddress2 = Request.Form("txtAddressLine2")
manufacturerCity = Request.Form("txtCity")
manufacturerState = Request.Form("txtState")
manufacturerZip = Request.Form("txtZipCode")
manufacturerPhone = Request.Form("txtPhoneNumber")
manufacturerFax = Request.Form("txtFaxNumber")
manufacturerEmail = Request.Form("txtEmailAddress")
InsightAssetTagPrefix = Request.Form("txtInsightAssetTagPrefix")


SQL = "INSERT INTO EQ_Manufacturers (ManufacturerName, Address1, Address2, City, State, Zip, Phone, Fax, Email, RecordSource, InsightAssetTagPrefix) "
SQL = SQL &  " VALUES (" 
SQL = SQL & "'" & manufacturerName & "','" & manufacturerAddress & "','" & manufacturerAddress2 & "','" & manufacturerCity & "','" & manufacturerState & "',"
SQL = SQL & "'" & manufacturerZip & "','" & manufacturerPhone & "','" & manufacturerFax & "','" & manufacturerEmail & "','Insight','" & InsightAssetTagPrefix & "')"

Set cnn8 = Server.CreateObject("ADODB.Connection")
cnn8.open (Session("ClientCnnString"))

Set rs8 = Server.CreateObject("ADODB.Recordset")
rs8.CursorLocation = 3 
Set rs8 = cnn8.Execute(SQL)
set rs8 = Nothing

Description = GetUserDisplayNameByUserNo(Session("Userno")) & " added the " & GetTerm("Equipment") & " Manufacturer : " & manufacturerName & " and Asset Tag Prefix, " & InsightAssetTagPrefix

CreateAuditLogEntry GetTerm("Equipment") & " Manufacturer Added",GetTerm("Equipment") & " Manufacturer Added","Minor",0,Description

Response.Redirect("main.asp")

%>















