<!--#include file="../../inc/subsandfuncs.asp"-->
<!--#include file="../../inc/InSightFuncs.asp"-->

<%
InternalRecordIdentifier = Request.Form("txtInternalRecordIdentifier")
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

'*******************************************************************
'Lookup the record as it exists now so we can fillin the audit trail

SQL = "SELECT * FROM EQ_Manufacturers where InternalRecordIdentifier = " & InternalRecordIdentifier
	
Set cnn8 = Server.CreateObject("ADODB.Connection")
cnn8.open (Session("ClientCnnString"))
Set rs = Server.CreateObject("ADODB.Recordset")
rs.CursorLocation = 3 
Set rs = cnn8.Execute(SQL)
	
If not rs.EOF Then
	Orig_ManufacturerName = rs("ManufacturerName")
	Orig_Address = rs("Address1")
	Orig_Address2 = rs("Address2")
	Orig_City = rs("City")
	Orig_State = rs("State")
	Orig_Zip = rs("Zip")
	Orig_Phone = rs("Phone")
	Orig_Fax = rs("Fax")
	Orig_Email = rs("Email")	
	Orig_RecordSource = rs("RecordSource")
	Orig_InsightAssetTagPrefix = rs("InsightAssetTagPrefix")
End If

set rs = Nothing
cnn8.close
set cnn8 = Nothing
'***********************************************************************
'End Lookup the record as it exists now so we can fillin the audit trail
'***********************************************************************

SQL = "UPDATE EQ_Manufacturers SET "
SQL = SQL &  "ManufacturerName = '" & manufacturerName & "',Address1 = '" & manufacturerAddress & "', "
SQL = SQL &  "Address2 = '" & manufacturerAddress2 & "',City = '" & manufacturerCity & "', "
SQL = SQL &  "State = '" & manufacturerState & "',Zip = '" & manufacturerZip & "', "
SQL = SQL &  "Phone = '" & manufacturerPhone & "',Fax = '" & manufacturerFax & "', InsightAssetTagPrefix = '" & InsightAssetTagPrefix & "', "
SQL = SQL &  "Email = '" & manufacturerEmail & "', RecordSource = 'Insight' WHERE InternalRecordIdentifier = " & InternalRecordIdentifier


Set cnn8 = Server.CreateObject("ADODB.Connection")
cnn8.open (Session("ClientCnnString"))

Set rs8 = Server.CreateObject("ADODB.Recordset")
rs8.CursorLocation = 3 
Set rs8 = cnn8.Execute(SQL)
set rs8 = Nothing


Description = ""


If Orig_ManufacturerName <> manufacturerName Then
	Description = GetTerm("Equipment") & " Manufacturer Company Name changed from " & Orig_ManufacturerName & " to " & manufacturerName 
	CreateAuditLogEntry GetTerm("Equipment") & " Manufacturer Edited",GetTerm("Equipment") & " Manufacturer Edited","Minor",0,Description
End If

If Orig_Address <> ManufacturerAddress Then
	Description = GetTerm("Equipment") & " Manufacturer, " & manufacturerName & ", Manufacturer Address 1 changed from " & Orig_Address & " to " & ManufacturerAddress 
	CreateAuditLogEntry GetTerm("Equipment") & " Manufacturer Edited",GetTerm("Equipment") & " Manufacturer Edited","Minor",0,Description
End If

If Orig_Address2 <> ManufacturerAddress2 Then
	Description = GetTerm("Equipment") & " Manufacturer, " & manufacturerName & ", Manufacturer Address 2 changed from " & Orig_Address2 & " to " & ManufacturerAddress2 
	CreateAuditLogEntry GetTerm("Equipment") & " Manufacturer Edited",GetTerm("Equipment") & " Manufacturer Edited","Minor",0,Description
End If

If Orig_City <> ManufacturerCity Then
	Description = GetTerm("Equipment") & " Manufacturer, " & manufacturerName & ", Manufacturer City changed from " & Orig_City & " to " & ManufacturerCity 
	CreateAuditLogEntry GetTerm("Equipment") & " Manufacturer Edited",GetTerm("Equipment") & " Manufacturer Edited","Minor",0,Description
End If

If Orig_State <> ManufacturerState Then
	Description = GetTerm("Equipment") & " Manufacturer, " & manufacturerName & ", Manufacturer State changed from " & Orig_State & " to " & ManufacturerState 
	CreateAuditLogEntry GetTerm("Equipment") & " Manufacturer Edited",GetTerm("Equipment") & " Manufacturer Edited","Minor",0,Description
End If

If Orig_Zip <> ManufacturerZip Then
	Description = GetTerm("Equipment") & " Manufacturer, " & manufacturerName & ", Manufacturer Zip Code changed from " & Orig_Zip & " to " & ManufacturerZip 
	CreateAuditLogEntry GetTerm("Equipment") & " Manufacturer Edited",GetTerm("Equipment") & " Manufacturer Edited","Minor",0,Description
End If

If Orig_Phone <> ManufacturerPhone Then
	Description = GetTerm("Equipment") & " Manufacturer, " & manufacturerName & ", Manufacturer Phone Number changed from " & Orig_Phone & " to " & ManufacturerPhone 
	CreateAuditLogEntry GetTerm("Equipment") & " Manufacturer Edited",GetTerm("Equipment") & " Manufacturer Edited","Minor",0,Description
End If

If Orig_Fax <> ManufacturerFax Then
	Description = GetTerm("Equipment") & " Manufacturer, " & manufacturerName & ", Manufacturer Fax Number changed from " & Orig_Fax & " to " & ManufacturerFax 
	CreateAuditLogEntry GetTerm("Equipment") & " Manufacturer Edited",GetTerm("Equipment") & " Manufacturer Edited","Minor",0,Description
End If

If Orig_Email <> ManufacturerEmail Then
	Description = GetTerm("Equipment") & " Manufacturer, " & manufacturerName & ", Manufacturer Primary Contact Email changed from " & Orig_Email & " to " & ManufacturerEmail 
	CreateAuditLogEntry GetTerm("Equipment") & " Manufacturer Edited",GetTerm("Equipment") & " Manufacturer Edited","Minor",0,Description
End If

If Orig_RecordSource <> RecordSource Then
	Description = GetTerm("Equipment") & " Manufacturer, " & manufacturerName & ", Manufacturer Record Source changed from " & Orig_RecordSource & " to " & RecordSource
	CreateAuditLogEntry GetTerm("Equipment") & " Manufacturer Edited",GetTerm("Equipment") & " Manufacturer Edited","Minor",0,Description
End If

If Orig_InsightAssetTagPrefix <> InsightAssetTagPrefix Then
	Description = Description & GetTerm("Equipment") & " Manufacturer, " & manufacturerName & ", changed the Insight Asset Tag Prefix from " & Orig_InsightAssetTagPrefix & " to " & InsightAssetTagPrefix
End If


Response.Redirect("main.asp")

%>