<!--#include file="../../inc/header.asp"-->
<!--#include file="../../inc/InSightFuncs_Equipment.asp"-->

<% 

InternalRecordIdentifier = Request.Form("txtInternalRecordIdentifier")

'*******************************************************************
'Lookup the record as it exists now so we can fill in the audit trail

SQL = "SELECT * FROM EQ_Equipment where InternalRecordIdentifier = " & InternalRecordIdentifier
	
Set cnn8 = Server.CreateObject("ADODB.Connection")
cnn8.open (Session("ClientCnnString"))
Set rs = Server.CreateObject("ADODB.Recordset")
rs.CursorLocation = 3 
Set rs = cnn8.Execute(SQL)
	
If not rs.EOF Then
	Orig_ModelIntRecID = rs("ModelIntRecID")
	Orig_StatusCodeIntRecID = rs("StatusCodeIntRecID")
	Orig_SerialNumber = rs("SerialNumber")
	Orig_AssetTag1 = rs("AssetTag1")
	Orig_AssetTag2 = rs("AssetTag2")
	Orig_AssetTag3 = rs("AssetTag3")
	Orig_AssetTag4 = rs("AssetTag4")
	Orig_AcquisitionCodeIntRecID = rs("AcquisitionCodeIntRecID")
	Orig_PurchasedFromVendorID = rs("PurchasedFromVendorID")
	Orig_PurchasedViaPONumber = rs("PurchasedViaPONumber")
	Orig_PurchaseDate = rs("PurchaseDate")
	Orig_PurchaseCost = rs("PurchaseCost")
	Orig_ReplacementCost = rs("ReplacementCost")
	Orig_AcquiredConditionIntRecID = rs("AquiredConditionIntRecID")
	Orig_CurrentConditionIntRecID = rs("CurrentConditionIntRecID")
	Orig_WarrentyStartDate = rs("WarrentyStartDate")
	Orig_WarrentyEndDate = rs("WarrentyEndDate")
	Orig_Comments = rs("Comments")
	Orig_Color = rs("Color")
	Orig_Size= rs("Size")
End If

set rs = Nothing
cnn8.close
set cnn8 = Nothing
'*************************************************************************
'End Lookup the record as it exists now so we can fill in the audit trail
'*************************************************************************

ModelIntRecID = Request.Form("selModelIntRecID")
StatusCodeIntRecID = Request.Form("selStatusCodeIntRecID")
SerialNumber = Request.Form("txtSerialNumber")
AssetTag1 = Request.Form("txtAssetTag1")
AssetTag2 = Request.Form("txtAssetTag2")
AssetTag3 = Request.Form("txtAssetTag3")
AssetTag4 = Request.Form("txtAssetTag4")
AcquisitionCodeIntRecID = Request.Form("selAcquisitionCodeIntRecID")
PurchasedFromVendorID = Request.Form("selVendorIntRecID")
PurchasedViaPONumber = Request.Form("txtPurchasedViaPONumber")
PurchaseDate = Request.Form("txtPurchaseDate")
PurchaseCost = Request.Form("txtPurchaseCost")
ReplacementCost = Request.Form("txtReplacementCost")
AcquiredConditionIntRecID = Request.Form("selAcquiredConditionIntRecID")
CurrentConditionIntRecID = Request.Form("selCurrentConditionIntRecID")
WarrentyStartDate = Request.Form("txtWarrentyStartDate")
WarrentyEndDate = Request.Form("txtWarrentyEndDate")
Comments = Request.Form("txtComments")
Color = Request.Form("txtColor")
Size = Request.Form("txtSize")


Comments = Replace(Comments, "'", "''")
Size = Replace(Size, "'", "''")


SQL = "UPDATE EQ_Equipment SET "
SQL = SQL &  "ModelIntRecID = " & ModelIntRecID & ", "

If StatusCodeIntRecID <> "" Then
	SQL = SQL &  "StatusCodeIntRecID = " & StatusCodeIntRecID & ", "
End If

SQL = SQL &  "SerialNumber = '" & SerialNumber & "', "
SQL = SQL &  "AssetTag1 = '" & AssetTag1 & "', "
SQL = SQL &  "AssetTag2 = '" & AssetTag2 & "', "
SQL = SQL &  "AssetTag3 = '" & AssetTag3 & "', "
SQL = SQL &  "AssetTag4 = '" & AssetTag4 & "', "

If AcquisitionCodeIntRecID <> "" Then
	SQL = SQL &  "AcquisitionCodeIntRecID = " & AcquisitionCodeIntRecID & ", "
End If

SQL = SQL &  "PurchasedViaPONumber = '" & PurchasedViaPONumber & "', "
SQL = SQL &  "PurchaseDate = '" & PurchaseDate & "', "

If PurchaseCost <> "" Then
	SQL = SQL &  "PurchaseCost = " & PurchaseCost & ", "
End If

If ReplacementCost <> "" Then
	SQL = SQL &  "ReplacementCost = " & ReplacementCost & ", "
End If

If AcquiredConditionIntRecID <> "" Then
	SQL = SQL &  "AquiredConditionIntRecID = " & AcquiredConditionIntRecID & ", "
End If

If CurrentConditionIntRecID <> "" Then
	SQL = SQL &  "CurrentConditionIntRecID = " & CurrentConditionIntRecID & ", "
End If

SQL = SQL &  "WarrentyStartDate = '" & WarrentyStartDate & "', "
SQL = SQL &  "WarrentyEndDate = '" & WarrentyEndDate & "', "
SQL = SQL &  "Comments = '" & Comments & "', "
SQL = SQL &  "Size= '" & Size & "' "
SQL = SQL &  " WHERE InternalRecordIdentifier = " & InternalRecordIdentifier

Response.Write("<br><br><br>" & SQL)

Set cnn8 = Server.CreateObject("ADODB.Connection")
cnn8.open (Session("ClientCnnString"))  

Set rs8 = Server.CreateObject("ADODB.Recordset")
rs8.CursorLocation = 3 
Set rs8 = cnn8.Execute(SQL)
set rs8 = Nothing


Description = ""
EQ_Desc = ""

If Orig_Model  <> Model  Then
	Description = Description & GetTerm("Equipment") & " ID, " & InternalRecordIdentifier & ", model name changed from " & GetModelNameByIntRecID(Orig_Model) & " to " & GetModelNameByIntRecID(Model)
	EQ_Desc = "Model name changed from " & GetModelNameByIntRecID(Orig_Model) & " to " & GetModelNameByIntRecID(Model)
	Record_EQ_Activity InternalRecordIdentifier,EQ_Desc,Session("UserNo")
End If

If Orig_StatusCodeIntRecID <> StatusCodeIntRecID Then
	Description = Description & GetTerm("Equipment") & " ID, " & InternalRecordIdentifier & ", status code changed from " & GetStatusCodeNameByIntRecID(Orig_StatusCodeIntRecID) & " to " & GetStatusCodeNameByIntRecID(StatusCodeIntRecID)
	EQ_Desc = "Status code changed from " & GetStatusCodeNameByIntRecID(Orig_StatusCodeIntRecID) & " to " & GetStatusCodeNameByIntRecID(StatusCodeIntRecID)
	Record_EQ_Activity InternalRecordIdentifier,EQ_Desc,Session("UserNo")
End If

If Orig_SerialNumber <> SerialNumber Then
	Description = Description & GetTerm("Equipment") & " ID, " & InternalRecordIdentifier & ", serial number changed from " & Orig_SerialNumber & " to " & SerialNumber
	EQ_Desc = "Serial number changed from " & Orig_SerialNumber & " to " & SerialNumber
	Record_EQ_Activity InternalRecordIdentifier,EQ_Desc,Session("UserNo")	
End If

If Orig_AssetTag1 <> AssetTag1 Then
	Description = Description & GetTerm("Equipment") & " ID, " & InternalRecordIdentifier & ", asset tag 1 changed from " & Orig_AssetTag1 & " to " & AssetTag1
	EQ_Desc = "Asset tag 1 changed from " & Orig_AssetTag1 & " to " & AssetTag1
	Record_EQ_Activity InternalRecordIdentifier,EQ_Desc,Session("UserNo")	
End If

If Orig_AssetTag2 <> AssetTag2 Then
	Description = Description & GetTerm("Equipment") & " ID, " & InternalRecordIdentifier & ", asset tag 2 changed from " & Orig_AssetTag2 & " to " & AssetTag2
	EQ_Desc = "Asset tag 2 changed from " & Orig_AssetTag2 & " to " & AssetTag2
	Record_EQ_Activity InternalRecordIdentifier,EQ_Desc,Session("UserNo")	
End If

If Orig_AssetTag3 <> AssetTag3 Then
	Description = Description & GetTerm("Equipment") & " ID, " & InternalRecordIdentifier & ", asset tag 3 changed from " & Orig_AssetTag3 & " to " & AssetTag3
	EQ_Desc = "Asset tag 3 changed from " & Orig_AssetTag3 & " to " & AssetTag3
	Record_EQ_Activity InternalRecordIdentifier,EQ_Desc,Session("UserNo")	
End If

If Orig_AssetTag4 <> AssetTag4 Then
	Description = Description & GetTerm("Equipment") & " ID, " & InternalRecordIdentifier & ", asset tag 4 changed from " & Orig_AssetTag4 & " to " & AssetTag4
	EQ_Desc = "Asset tag 4 changed from " & Orig_AssetTag4 & " to " & AssetTag4
	Record_EQ_Activity InternalRecordIdentifier,EQ_Desc,Session("UserNo")	
End If

If Orig_AcquisitionCodeIntRecID <> AcquisitionCodeIntRecID Then
	Description = Description & GetTerm("Equipment") & " ID, " & InternalRecordIdentifier & ", acquisition type changed from " & GetAcquisitionCodeByIntRecID(Orig_AcquisitionCodeIntRecID) & " to " & GetAcquisitionCodeByIntRecID(AcquisitionCodeIntRecID)
	EQ_Desc = "Acquisition type changed from " & GetAcquisitionCodeByIntRecID(Orig_AcquisitionCodeIntRecID) & " to " & GetAcquisitionCodeByIntRecID(AcquisitionCodeIntRecID)
	Record_EQ_Activity InternalRecordIdentifier,EQ_Desc,Session("UserNo")		
End If

If Orig_PurchasedFromVendorID <> PurchasedFromVendorID Then
	Description = Description & GetTerm("Equipment") & " ID, " & InternalRecordIdentifier & ", vendor changed from " & GetEquipVendorNameByVendorID(Orig_PurchasedFromVendorID) & " to " & GetEquipVendorNameByVendorID(PurchasedFromVendorID)
	EQ_Desc = "Vendor changed from " & GetEquipVendorNameByVendorID(Orig_PurchasedFromVendorID) & " to " & GetEquipVendorNameByVendorID(PurchasedFromVendorID)
	Record_EQ_Activity InternalRecordIdentifier,EQ_Desc,Session("UserNo")
End If

If Orig_PurchasedViaPONumber <> PurchasedViaPONumber Then
	Description = Description & GetTerm("Equipment") & " ID, " & InternalRecordIdentifier & ", purchase PO Number changed from " & Orig_PurchasedViaPONumber & " to " & PurchasedViaPONumber
	EQ_Desc = "Purchase PO Number changed from " & Orig_PurchasedViaPONumber & " to " & PurchasedViaPONumber
	Record_EQ_Activity InternalRecordIdentifier,EQ_Desc,Session("UserNo")
End If

If Orig_PurchaseDate <> PurchaseDate Then
	Description = Description & GetTerm("Equipment") & " ID, " & InternalRecordIdentifier & ", purchase date changed from " & Orig_PurchaseDate & " to " & PurchaseDate
	EQ_Desc = "Purchase date changed from " & Orig_PurchaseDate & " to " & PurchaseDate
	Record_EQ_Activity InternalRecordIdentifier,EQ_Desc,Session("UserNo")
End If

If cDbl(Orig_PurchaseCost) <> cDbl(PurchaseCost) Then
	Description = Description & GetTerm("Equipment") & " ID, " & InternalRecordIdentifier & ", purchase cost changed from " & Orig_PurchaseCost & " to " & PurchaseCost
	EQ_Desc = "Purchase cost changed from " & Orig_PurchaseCost & " to " & PurchaseCost
	Record_EQ_Activity InternalRecordIdentifier,EQ_Desc,Session("UserNo")	
End If

If cDbl(Orig_ReplacementCost) <> cDbl(ReplacementCost) Then
	Description = Description & GetTerm("Equipment") & " ID, " & InternalRecordIdentifier & ", replacement cost changed from " & Orig_ReplacementCost & " to " & ReplacementCost
	EQ_Desc = "Replacement cost changed from " & Orig_ReplacementCost & " to " & ReplacementCost
	Record_EQ_Activity InternalRecordIdentifier,EQ_Desc,Session("UserNo")	
End If

If Orig_AcquiredConditionIntRecID <> AcquiredConditionIntRecID Then
	Description = Description & GetTerm("Equipment") & " ID, " & InternalRecordIdentifier & ", Acquired condition changed from " & GetConditionNameByIntRecID(Orig_AcquiredConditionIntRecID) & " to " & GetConditionNameByIntRecID(AcquiredConditionIntRecID)
	EQ_Desc = "Acquired condition changed from " & GetConditionNameByIntRecID(Orig_AcquiredConditionIntRecID) & " to " & GetConditionNameByIntRecID(AcquiredConditionIntRecID)
	Record_EQ_Activity InternalRecordIdentifier,EQ_Desc,Session("UserNo")
End If

If Orig_CurrentConditionIntRecID <> CurrentConditionIntRecID Then
	Description = Description & GetTerm("Equipment") & " ID, " & InternalRecordIdentifier & ", Current condition changed from " & GetConditionNameByIntRecID(Orig_CurrentConditionIntRecID) & " to " & GetConditionNameByIntRecID(CurrentConditionIntRecID)
	EQ_Desc = "Current condition changed from " & GetConditionNameByIntRecID(Orig_CurrentConditionIntRecID) & " to " & GetConditionNameByIntRecID(CurrentConditionIntRecID)
	Record_EQ_Activity InternalRecordIdentifier,EQ_Desc,Session("UserNo")
End If

If Orig_WarrentyStartDate <> WarrentyStartDate Then
	Description = Description & GetTerm("Equipment") & " ID, " & InternalRecordIdentifier & ", warranty start date changed from " & Orig_WarrentyStartDate & " to " & WarrentyStartDate
	EQ_Desc = "Warranty start date changed from " & Orig_WarrentyStartDate & " to " & WarrentyStartDate
	Record_EQ_Activity InternalRecordIdentifier,EQ_Desc,Session("UserNo")
End If

If Orig_WarrentyEndDate <> WarrentyEndDate Then
	Description = Description & GetTerm("Equipment") & " ID, " & InternalRecordIdentifier & ", warranty end date changed from " & Orig_WarrentyEndDate & " to " & WarrentyEndDate
	EQ_Desc = "Warranty end date changed from " & Orig_WarrentyEndDate & " to " & WarrentyEndDate
	Record_EQ_Activity InternalRecordIdentifier,EQ_Desc,Session("UserNo")
End If

If Orig_Comments <> Comments Then
	Description = Description & GetTerm("Equipment") & " ID, " & InternalRecordIdentifier & ", comments changed from " & Orig_Comments & " to " & Comments
	EQ_Desc = "Comments changed from " & Orig_Comments & " to " & Comments
	Record_EQ_Activity InternalRecordIdentifier,EQ_Desc,Session("UserNo")
End If

If Orig_Color <> Color Then
	Description = Description & GetTerm("Equipment") & " ID, " & InternalRecordIdentifier & ", color changed from " & Orig_Color & " to " & Color
	EQ_Desc = "Color changed from " & Orig_Color & " to " & Color
	Record_EQ_Activity InternalRecordIdentifier,EQ_Desc,Session("UserNo")
End If

If Orig_Size <> Size Then
	Description = Description & GetTerm("Equipment") & " ID, " & InternalRecordIdentifier & ", size changed from " & Orig_Size & " to " & Size
	EQ_Desc = "Size changed from " & Orig_Size & " to " & Size
	Record_EQ_Activity InternalRecordIdentifier,EQ_Desc,Session("UserNo")
End If


CreateAuditLogEntry GetTerm("Equipment") & " Existing Equipment Edited",GetTerm("Equipment") & " Existing Equipment Edited","Minor",0,Description


'*****************************************************************************************************************************
'Code to update equipment_list_CLIENTID.json if necessary, so equipment search works
'*****************************************************************************************************************************

If (Orig_ModelIntRecID <> ModelIntRecID) OR (Orig_SerialNumber <> SerialNumber) OR (Orig_AssetTag1 <> AssetTag1) Then

	ClientKeyForFileNames = MUV_READ("ClientKeyForFileNames")
	
	Const ForReading = 1
	Const ForWriting = 2
	
	Set objFSO = CreateObject("Scripting.FileSystemObject")
	Set objFile = objFSO.OpenTextFile(Server.MapPath("..\..\..\") & "\clientfiles\"  & ClientKeyForFileNames & "\autocomplete\equipment_list_" & ClientKeyForFileNames & ".json", ForReading)
	
	strCurrentText = objFile.ReadAll
	
	'Response.Write(strCurrentText & "<br><br>")
	
	objFile.Close
	
	strOriginalText = "{""name"":""" & GetModelNameByIntRecID(Orig_ModelIntRecID) & " --- " & Orig_SerialNumber & " --- " & Orig_AssetTag1 & """, ""code"":""" & InternalRecordIdentifier & """}"
	
	strUpdatedText = "{""name"":""" & GetModelNameByIntRecID(ModelIntRecID) & " --- " & SerialNumber & " --- " & AssetTag1 & """, ""code"":""" & InternalRecordIdentifier & """}"
	
	strNewText = Replace(strCurrentText, strOriginalText, strUpdatedText)
	
	
	Set objFile = objFSO.OpenTextFile(Server.MapPath("..\..\..\") & "\clientfiles\"  & ClientKeyForFileNames & "\autocomplete\equipment_list_" & ClientKeyForFileNames & ".json", ForWriting)
	objFile.WriteLine strNewText
	objFile.Close
	
End If

'*****************************************************************************************************************************
'End Code to update equipment_list_CLIENTID.json
'*****************************************************************************************************************************


Response.Redirect("findEquipment.asp")

%>