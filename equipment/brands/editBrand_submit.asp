<!--#include file="../../inc/subsandfuncs.asp"-->
<!--#include file="../../inc/InSightFuncs.asp"-->
<!--#include file="../../inc/InSightFuncs_Equipment.asp"-->

<%
InternalRecordIdentifier = Request.Form("txtInternalRecordIdentifier")
Brand = Request.Form("txtBrand")
ManufacturerIntRecID = Request.Form("selManufacturerIntRecID")
InsightAssetTagPrefix = Request.Form("txtInsightAssetTagPrefix")

'*******************************************************************
'Lookup the record as it exists now so we can fillin the audit trail

SQL = "SELECT * FROM EQ_Brands where InternalRecordIdentifier = " & InternalRecordIdentifier
	
Set cnn8 = Server.CreateObject("ADODB.Connection")
cnn8.open (Session("ClientCnnString"))
Set rs = Server.CreateObject("ADODB.Recordset")
rs.CursorLocation = 3 
Set rs = cnn8.Execute(SQL)
	
If not rs.EOF Then
	Orig_Brand = rs("Brand")
	Orig_ManufacturerIntRecID = rs("ManufacIntRecID")
	Orig_InsightAssetTagPrefix = rs("InsightAssetTagPrefix")
	Orig_RecordSource = rs("RecordSource")
End If

set rs = Nothing
cnn8.close
set cnn8 = Nothing
'***********************************************************************
'End Lookup the record as it exists now so we can fillin the audit trail
'***********************************************************************

SQL = "UPDATE EQ_Brands SET "
SQL = SQL &  "ManufacIntRecID = " & ManufacturerIntRecID & ", "
SQL = SQL &  "Brand = '" & Brand & "', "
SQL = SQL &  "InsightAssetTagPrefix = '" & InsightAssetTagPrefix & "', "
SQL = SQL &  "RecordSource = 'Insight' "
SQL = SQL &  " WHERE InternalRecordIdentifier = " & InternalRecordIdentifier


Set cnn8 = Server.CreateObject("ADODB.Connection")
cnn8.open (Session("ClientCnnString"))

Set rs8 = Server.CreateObject("ADODB.Recordset")
rs8.CursorLocation = 3 
Set rs8 = cnn8.Execute(SQL)
set rs8 = Nothing


Description = ""
If Orig_Brand  <> Brand  Then
	Description = Description & GetTerm("Equipment") & " Brand: " & Brand & " with Manufacturer, " & GetManufacturerNameByIntRecID(ManufacturerIntRecID) & " changed from " & Orig_Brand & " to " & Brand
End If
If Orig_ManufacturerIntRecID <> ManufacturerIntRecID Then
	Description = Description & GetTerm("Equipment") & " brand, " & Brand & ", changed manufacturer from " & Orig_ManufacturerIntRecID & " to " & ManufacturerIntRecID 
End If
If Orig_InsightAssetTagPrefix <> InsightAssetTagPrefix Then
	Description = Description & GetTerm("Equipment") & " brand, " & Brand & ", changed the Insight Asset Tag Prefix from " & Orig_InsightAssetTagPrefix & " to " & InsightAssetTagPrefix
End If
If Orig_RecordSource <> RecordSource Then
	Description = Description & GetTerm("Equipment") & " brand, " & Brand & ", changed Record Source from " & Orig_RecordSource & " to " & RecordSource
End If

CreateAuditLogEntry GetTerm("Equipment") & " Brand Edited",GetTerm("Equipment") & " Brand Edited","Minor",0,Description


Response.Redirect("main.asp")

%>















