<!--#include file="../../inc/header.asp"-->
<!--#include file="../../inc/InSightFuncs_Equipment.asp"-->

<% 

ModelIntRecID = Request.Form("selModelIntRecID")
StatusCodeIntRecID = Request.Form("selStatusCodeIntRecID")
SerialNumber = Request.Form("txtSerialNumber")
AssetTag1 = Request.Form("txtAssetTag1")
AssetTag2 = Request.Form("txtAssetTag2")
AssetTag3 = Request.Form("txtAssetTag3")
AssetTag4 = Request.Form("txtAssetTag4")
AcquisitionIntRecID = Request.Form("selAcquisitionCodeIntRecID")
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



SQL = "INSERT INTO EQ_Equipment ("

SQL = SQL &  " ModelIntRecID, "

If StatusCodeIntRecID <> "" Then
	SQL = SQL &  " StatusCodeIntRecID, "
End If


SQL = SQL &  " SerialNumber, "
SQL = SQL &  " AssetTag1, "
SQL = SQL &  " AssetTag2, "
SQL = SQL &  " AssetTag3, "
SQL = SQL &  " AssetTag4, "


If PurchasedFromVendorID <> "" Then
	SQL = SQL &  " PurchasedFromVendorID,"
End If

SQL = SQL &  " PurchasedViaPONumber, "
SQL = SQL &  " PurchaseDate, "

If PurchaseCost <> "" Then
	SQL = SQL &  " PurchaseCost, "
End If

If ReplacementCost <> "" Then
	SQL = SQL &  " ReplacementCost, "
End If

If AcquiredConditionIntRecID <> "" Then
	SQL = SQL & " AquiredConditionIntRecID,"
End If

If CurrentConditionIntRecID <> "" Then
	SQL = SQL &  " CurrentConditionIntRecID,"
End If

If AcquisitionIntRecID <> "" Then
	SQL = SQL &  " AcquisitionCodeIntRecID ,"
End If


SQL = SQL &  " WarrentyStartDate, "
SQL = SQL &  " WarrentyEndDate, "
SQL = SQL &  " Comments, "
SQL = SQL &  " Color, "
SQL = SQL &  " Size)"

SQL = SQL &  " VALUES (" 

SQL = SQL & ModelIntRecID & "," 

If StatusCodeIntRecID <> "" Then
	SQL = SQL & StatusCodeIntRecID & ","
End If

SQL = SQL & "'" & SerialNumber & "'," 
SQL = SQL & "'" & AssetTag1 & "',"
SQL = SQL & "'" & AssetTag2 & "',"
SQL = SQL & "'" & AssetTag3 & "',"
SQL = SQL & "'" & AssetTag4 & "',"


If PurchasedFromVendorID <> "" Then
	SQL = SQL & PurchasedFromVendorID & ","
End If

SQL = SQL & "'" & PurchasedViaPONumber & "',"
SQL = SQL & "'" & PurchaseDate & "',"

If PurchaseCost <> "" Then
	SQL = SQL & PurchaseCost & ","
End If

If ReplacementCost <> "" Then
	SQL = SQL & ReplacementCost & ","
End If

If AcquiredConditionIntRecID <> "" Then
	SQL = SQL & AcquiredConditionIntRecID & ","
End If

If CurrentConditionIntRecID <> "" Then
	SQL = SQL & CurrentConditionIntRecID & ","
End If

If AcquisitionIntRecID <> "" Then
	SQL = SQL & AcquisitionIntRecID & ","
End If

SQL = SQL & "'" & WarrentyStartDate & "',"
SQL = SQL & "'" & WarrentyEndDate & "',"
SQL = SQL & "'" & Comments & "',"
SQL = SQL & "'" & Color & "',"
SQL = SQL & "'" & Size & "')"


Response.Write("<br><br><br>" & SQL)

Set cnn8 = Server.CreateObject("ADODB.Connection")
cnn8.open (Session("ClientCnnString"))  

Set rs8 = Server.CreateObject("ADODB.Recordset")
rs8.CursorLocation = 3 
Set rs8 = cnn8.Execute(SQL)
set rs8 = Nothing



'*************************************************************************
'Lookup the record just inserted to obtain InternalRecordIdentifier 
'*************************************************************************

SQL = "SELECT * FROM EQ_Equipment ORDER BY InternalRecordIdentifier DESC"
	
Set cnn8 = Server.CreateObject("ADODB.Connection")
cnn8.open (Session("ClientCnnString"))
Set rs = Server.CreateObject("ADODB.Recordset")
rs.CursorLocation = 3 
Set rs = cnn8.Execute(SQL)
	
If not rs.EOF Then
	InternalRecordIdentifier = rs("InternalRecordIdentifier")
End If

set rs = Nothing
cnn8.close
set cnn8 = Nothing
'*************************************************************************
'End Lookup InternalRecordIdentifier 
'*************************************************************************


Description = ""
EQ_Desc = ""

If Model <> "" Then
	Description = Description & GetTerm("Equipment") & " ID, " & InternalRecordIdentifier & ", was created with model " & GetModelNameByIntRecID(Model)
	EQ_Desc = "Model name " & GetModelNameByIntRecID(Model) & " assigned to new equipment piece."
	Record_EQ_Activity InternalRecordIdentifier,EQ_Desc,Session("UserNo")
End If

If StatusCodeIntRecID <> "" Then
	Description = Description & GetTerm("Equipment") & " ID, " & InternalRecordIdentifier & ", was created with status code " & GetStatusCodeNameByIntRecID(StatusCodeIntRecID)
	EQ_Desc = "Status code " & GetStatusCodeNameByIntRecID(StatusCodeIntRecID) & " assigned to new equipment piece."
	Record_EQ_Activity InternalRecordIdentifier,EQ_Desc,Session("UserNo")
End If

If SerialNumber <> "" Then
	Description = Description & GetTerm("Equipment") & " ID, " & InternalRecordIdentifier & ", was created with serial number " & SerialNumber
	EQ_Desc = "Serial number " & SerialNumber & " assigned to new equipment piece."
	Record_EQ_Activity InternalRecordIdentifier,EQ_Desc,Session("UserNo")	
End If

If AssetTag1 <> "" Then
	Description = Description & GetTerm("Equipment") & " ID, " & InternalRecordIdentifier & ", was created with asset tag 1 " & AssetTag1
	EQ_Desc = "Asset tag 1 " & AssetTag1 & " assigned to new equipment piece."
	Record_EQ_Activity InternalRecordIdentifier,EQ_Desc,Session("UserNo")	
End If

If AssetTag2 <> "" Then
	Description = Description & GetTerm("Equipment") & " ID, " & InternalRecordIdentifier & ", was created with asset tag 2 " & AssetTag2
	EQ_Desc = "Asset tag 2 " & AssetTag2 & " assigned to new equipment piece."
	Record_EQ_Activity InternalRecordIdentifier,EQ_Desc,Session("UserNo")	
End If

If AssetTag3 <> "" Then
	Description = Description & GetTerm("Equipment") & " ID, " & InternalRecordIdentifier & ", was created with asset tag 3 " & AssetTag3
	EQ_Desc = "Asset tag 3 " & AssetTag3 & " assigned to new equipment piece."
	Record_EQ_Activity InternalRecordIdentifier,EQ_Desc,Session("UserNo")	
End If

If AssetTag4 <> "" Then
	Description = Description & GetTerm("Equipment") & " ID, " & InternalRecordIdentifier & ", was created with asset tag 4 " & AssetTag4
	EQ_Desc = "Asset tag 4 " & AssetTag4 & " assigned to new equipment piece."
	Record_EQ_Activity InternalRecordIdentifier,EQ_Desc,Session("UserNo")	
End If

If PurchasedFromVendorID <> "" Then
	Description = Description & GetTerm("Equipment") & " ID, " & InternalRecordIdentifier & ", was created with vendor " & GetEquipVendorNameByVendorID(PurchasedFromVendorID)
	EQ_Desc = "Vendor " & GetEquipVendorNameByVendorID(PurchasedFromVendorID) & " assigned to new equipment piece."
	Record_EQ_Activity InternalRecordIdentifier,EQ_Desc,Session("UserNo")
End If

If PurchasedViaPONumber <> "" Then
	Description = Description & GetTerm("Equipment") & " ID, " & InternalRecordIdentifier & ", was created with purchase PO Number " & PurchasedViaPONumber
	EQ_Desc = "Purchase PO Number " & PurchasedViaPONumber & " assigned to new equipment piece."
	Record_EQ_Activity InternalRecordIdentifier,EQ_Desc,Session("UserNo")
End If

If PurchaseDate <> "" Then
	Description = Description & GetTerm("Equipment") & " ID, " & InternalRecordIdentifier & ", was created with purchase date " & PurchaseDate
	EQ_Desc = "Purchase date " & PurchaseDate & " assigned to new equipment piece."
	Record_EQ_Activity InternalRecordIdentifier,EQ_Desc,Session("UserNo")
End If

If PurchaseCost <> "" Then
	Description = Description & GetTerm("Equipment") & " ID, " & InternalRecordIdentifier & ", was created with purchase cost " & PurchaseCost
	EQ_Desc = "Purchase cost " & PurchaseCost & " assigned to new equipment piece."
	Record_EQ_Activity InternalRecordIdentifier,EQ_Desc,Session("UserNo")	
End If

If ReplacementCost <> "" Then
	Description = Description & GetTerm("Equipment") & " ID, " & InternalRecordIdentifier & ", was created with replacement cost " & ReplacementCost
	EQ_Desc = "Replacement cost " & ReplacementCost & " assigned to new equipment piece."
	Record_EQ_Activity InternalRecordIdentifier,EQ_Desc,Session("UserNo")	
End If

If AcquiredConditionIntRecID <> "" Then
	Description = Description & GetTerm("Equipment") & " ID, " & InternalRecordIdentifier & ", was created with Acquired condition " & GetConditionNameByIntRecID(AcquiredConditionIntRecID)
	EQ_Desc = "Acquired condition " & GetConditionNameByIntRecID(AcquiredConditionIntRecID) & " assigned to new equipment piece."
	Record_EQ_Activity InternalRecordIdentifier,EQ_Desc,Session("UserNo")
End If

If CurrentConditionIntRecID <> "" Then
	Description = Description & GetTerm("Equipment") & " ID, " & InternalRecordIdentifier & ", was created with Current condition " & GetConditionNameByIntRecID(CurrentConditionIntRecID)
	EQ_Desc = "Current condition " & GetConditionNameByIntRecID(CurrentConditionIntRecID) & " assigned to new equipment piece."
	Record_EQ_Activity InternalRecordIdentifier,EQ_Desc,Session("UserNo")
End If

If AcquisitionCodeIntRecID <> "" Then
	Description = Description & GetTerm("Equipment") & " ID, " & InternalRecordIdentifier & ", was created with Acquisition Code " & GetAcquisitionCodeByIntRecID(AcquisitionCodeIntRecID)
	EQ_Desc = "Acquisition Code " & GetAcquisitionCodeByIntRecID(AcquisitionCodeIntRecID) & " assigned to new equipment piece."
	Record_EQ_Activity InternalRecordIdentifier,EQ_Desc,Session("UserNo")
End If


If WarrentyStartDate <> "" Then
	Description = Description & GetTerm("Equipment") & " ID, " & InternalRecordIdentifier & ", was created with warranty start date " & WarrentyStartDate
	EQ_Desc = "Warranty start date " & WarrentyStartDate & " assigned to new equipment piece."
	Record_EQ_Activity InternalRecordIdentifier,EQ_Desc,Session("UserNo")
End If

If WarrentyEndDate <> "" Then
	Description = Description & GetTerm("Equipment") & " ID, " & InternalRecordIdentifier & ", was created with warranty end date " & WarrentyEndDate
	EQ_Desc = "Warranty end date " & WarrentyEndDate & " assigned to new equipment piece."
	Record_EQ_Activity InternalRecordIdentifier,EQ_Desc,Session("UserNo")
End If

If Comments <> "" Then
	Description = Description & GetTerm("Equipment") & " ID, " & InternalRecordIdentifier & ", was created with comments " & Comments
	EQ_Desc = "Comments " & Comments & " assigned to new equipment piece."
	Record_EQ_Activity InternalRecordIdentifier,EQ_Desc,Session("UserNo")
End If

If Color <> "" Then
	Description = Description & GetTerm("Equipment") & " ID, " & InternalRecordIdentifier & ", was created with color " & Color
	EQ_Desc = "Color " & Color & " assigned to new equipment piece."
	Record_EQ_Activity InternalRecordIdentifier,EQ_Desc,Session("UserNo")
End If

If Size <> "" Then
	Description = Description & GetTerm("Equipment") & " ID, " & InternalRecordIdentifier & ", was created with size " & Size
	EQ_Desc = "Size " & Size & " assigned to new equipment piece."
	Record_EQ_Activity InternalRecordIdentifier,EQ_Desc,Session("UserNo")
End If


CreateAuditLogEntry GetTerm("Equipment") & " New Equipment Piece Added",GetTerm("Equipment") & " New Equipment Piece Added","Minor",0,Description


'*****************************************************************************************************************************
'Code to update equipment_list_CLIENTID.json if necessary, so equipment search works
'*****************************************************************************************************************************


	ClientKeyForFileNames = MUV_READ("ClientKeyForFileNames")
	
	Const ForReading = 1
	Const ForWriting = 2
	
	Set objFSO = CreateObject("Scripting.FileSystemObject")
	Set objFile = objFSO.OpenTextFile(Server.MapPath("..\..\..\") & "\clientfiles\"  & ClientKeyForFileNames & "\autocomplete\equipment_list_" & ClientKeyForFileNames & ".json", ForReading)
	
	strCurrentText = objFile.ReadAll
	
	objFile.Close
	
	strUpdatedText = "},{""name"":""" & GetModelNameByIntRecID(ModelIntRecID) & " --- " & SerialNumber & " --- " & AssetTag1 & """, ""code"":""" & InternalRecordIdentifier & """}]"
	
	strNewText = Replace(strCurrentText, "}]", strUpdatedText)
	
	Set objFile = objFSO.OpenTextFile(Server.MapPath("..\..\..\") & "\clientfiles\"  & ClientKeyForFileNames & "\autocomplete\equipment_list_" & ClientKeyForFileNames & ".json", ForWriting)
	objFile.WriteLine strNewText
	objFile.Close
	


'*****************************************************************************************************************************
'End Code to update equipment_list_CLIENTID.json
'*****************************************************************************************************************************



Response.Redirect("addEquipment.asp")
%>