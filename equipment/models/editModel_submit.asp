<!--#include file="../../inc/subsandfuncs.asp"-->
<!--#include file="../../inc/InSightFuncs.asp"-->
<!--#include file="../../inc/InSightFuncs_Equipment.asp"-->
<%
InternalRecordIdentifier = Request.Form("txtInternalRecordIdentifier")
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

'*******************************************************************
'Lookup the record as it exists now so we can fillin the audit trail

SQL = "SELECT * FROM EQ_Models where InternalRecordIdentifier = " & InternalRecordIdentifier
	
Set cnn8 = Server.CreateObject("ADODB.Connection")
cnn8.open (Session("ClientCnnString"))
Set rs = Server.CreateObject("ADODB.Recordset")
rs.CursorLocation = 3 
Set rs = cnn8.Execute(SQL)
	
If not rs.EOF Then
	Orig_Model = rs("Model")
	Orig_BrandIntRecID = rs("BrandIntRecID")
	Orig_GroupIntRecID= rs("GroupIntRecID")
	Orig_ClassIntRecID = rs("ClassIntRecID")
	Orig_ManufacIntRecID = rs("ManufacIntRecID")
	Orig_DefaultRentalPrice = rs("DefaultRentalPrice")
	Orig_DefaultCostPrice = rs("DefaultCostPrice")	
	Orig_ReplacementCost = rs("ReplacementCost")
	Orig_BackendSystemCode = rs("BackendSystemCode")
	Orig_RecordSource = rs("RecordSource")
	Orig_InsightAssetTagPrefix = rs("InsightAssetTagPrefix")
End If

set rs = Nothing
cnn8.close
set cnn8 = Nothing
'***********************************************************************
'End Lookup the record as it exists now so we can fillin the audit trail
'***********************************************************************

SQL = "UPDATE EQ_Models SET "
SQL = SQL &  "Model = '" & Model & "', "
SQL = SQL &  "BrandIntRecID = " & BrandIntRecID & ", "
SQL = SQL &  "GroupIntRecID = " & GroupIntRecID & ", "
SQL = SQL &  "ClassIntRecID = " & ClassIntRecID & ", "
SQL = SQL &  "ManufacIntRecID = " & ManufacturerIntRecID & ", "
SQL = SQL &  "DefaultRentalPrice = " & DefaultRentalPrice & ", "
SQL = SQL &  "DefaultCostPrice = " & DefaultCost & ", "
SQL = SQL &  "ReplacementCost = " & ReplacementCost & ", "
SQL = SQL &  "BackendSystemCode = '" & BackendSystemCode & "', "
SQL = SQL &  "InsightAssetTagPrefix = '" & InsightAssetTagPrefix & "', "
SQL = SQL &  "RecordSource = 'Insight' "
SQL = SQL &  " WHERE InternalRecordIdentifier = " & InternalRecordIdentifier

Response.write(SQL)

Set cnn8 = Server.CreateObject("ADODB.Connection")
cnn8.open (Session("ClientCnnString"))

Set rs8 = Server.CreateObject("ADODB.Recordset")
rs8.CursorLocation = 3 
Set rs8 = cnn8.Execute(SQL)
set rs8 = Nothing


Description = ""
If Orig_Model  <> Model  Then
	Description = Description & GetTerm("Equipment") & " model name changed from " & Orig_Model & " to " & Model
End If
If Orig_BrandIntRecID <> BrandIntRecID Then
	Description = Description & GetTerm("Equipment") & " brand changed from " & GetBrandNameByIntRecID(Orig_BrandIntRecID) & " to " & GetBrandNameByIntRecID(BrandIntRecID)
End If
If Orig_GroupIntRecID <> GroupIntRecID Then
	Description = Description & GetTerm("Equipment") & " Group changed from " & GetGroupNameByIntRecID(Orig_GroupIntRecID) & " to " & GetGroupNameByIntRecID(GroupIntRecID)
End If
If Orig_ClassIntRecID <> ClassIntRecID Then
	Description = Description & GetTerm("Equipment") & " class changed from " & GetClassNameByIntRecID(Orig_ClassIntRecID) & " to " & GetClassNameByIntRecID(ClassIntRecID)
End If
If Orig_ManufacIntRecID <> ManufacturerIntRecID Then
	Description = Description & GetTerm("Equipment") & " manufacturer changed from " & GetManufacturerNameByIntRecID(Orig_ManufacIntRecID) & " to " & GetManufacturerNameByIntRecID(ManufacturerIntRecID)
End If
If Orig_DefaultRentalPrice <> DefaultRentalPrice Then
	Description = Description & GetTerm("Equipment") & " default rental price changed from " & Orig_DefaultRentalPrice & " to " & DefaultRentalPrice
End If
If Orig_DefaultCostPrice <> DefaultCost Then
	Description = Description & GetTerm("Equipment") & " default cost changed from " & Orig_DefaultCostPrice & " to " & DefaultCost
End If
If Orig_ReplacementCost <> ReplacementCost Then
	Description = Description & GetTerm("Equipment") & " replacement cost changed from " & Orig_ReplacementCost & " to " & ReplacementCost
End If
If Orig_BackendSystemCode <> BackendSystemCode Then
	Description = Description & GetTerm("Equipment") & " backend system code changed from " & Orig_BackendSystemCode & " to " & BackendSystemCode
End If
If Orig_RecordSource <> RecordSource Then
	Description = Description & GetTerm("Equipment") & " record source changed from " & Orig_RecordSource & " to " & RecordSource
End If
If Orig_InsightAssetTagPrefix <> InsightAssetTagPrefix Then
	Description = Description & GetTerm("Equipment") & " model, " & Model & ", changed the Insight Asset Tag Prefix from " & Orig_InsightAssetTagPrefix & " to " & InsightAssetTagPrefix
End If


CreateAuditLogEntry GetTerm("Equipment") & " Model Edited",GetTerm("Equipment") & " Model Edited","Minor",0,Description

Response.Redirect("main.asp")

%>















