<!--#include file="../../inc/subsandfuncs.asp"-->
<!--#include file="../../inc/InSightFuncs.asp"-->

<%
InternalRecordIdentifier = Request.Form("txtInternalRecordIdentifier")

'*******************************************************************
'Lookup the record as it exists now so we can fillin the audit trail

SQL = "SELECT * FROM PR_Competitors where InternalRecordIdentifier = " & InternalRecordIdentifier
	
Set cnn8 = Server.CreateObject("ADODB.Connection")
cnn8.open (Session("ClientCnnString"))
Set rs = Server.CreateObject("ADODB.Recordset")
rs.CursorLocation = 3 
Set rs = cnn8.Execute(SQL)
	
If not rs.EOF Then
	Orig_Competitor = rs("CompetitorName")
	Orig_CompetitorAddressInfo = rs("AddressInformation")
	Orig_BottledWater = rs("BottledWater")
	Orig_FilteredWater = rs("FilteredWater")
	Orig_OCS = rs("OCS")
	Orig_OCS_Supply = rs("OCS_Supply")
	Orig_OfficeSupplies = rs("OfficeSupplies")
	Orig_Vending = rs("Vending")
	Orig_Micromarket = rs("Micromarket")
	Orig_Pantry = rs("Pantry")	
End If

set rs = Nothing
cnn8.close
set cnn8 = Nothing
'***********************************************************************
'End Lookup the record as it exists now so we can fillin the audit trail
'***********************************************************************

competitorName 			= Request.Form("txtCompetitorName")
competitorAddressInfo	= Request.Form("txtCompetitorAddressInfo")
chkOfficeCoffee			= Request.Form("chkOfficeCoffee")
chkOfficeCoffeeSupply	= Request.Form("chkOfficeCoffeeSupply")
chkBottledWater			= Request.Form("chkBottledWater")
chkFilteredWater		= Request.Form("chkFilteredWater")
chkVending				= Request.Form("chkVending")
chkMicromarket			= Request.Form("chkMicromarket")
chkPantry 				= Request.Form("chkPantry")
chkOfficeSupplies 		= Request.Form("chkOfficeSupplies")

If chkOfficeCoffee <> "" Then
	chkOfficeCoffee = 1
Else
	chkOfficeCoffee = 0
End If


If chkOfficeCoffeeSupply <> "" Then
	chkOfficeCoffeeSupply = 1
Else
	chkOfficeCoffeeSupply = 0
End If


If chkBottledWater <> "" Then
	chkBottledWater = 1
Else
	chkBottledWater = 0
End If


If chkFilteredWater <> "" Then
	chkFilteredWater = 1
Else
	chkFilteredWater = 0
End If


If chkVending <> "" Then
	chkVending = 1
Else
	chkVending = 0
End If


If chkMicromarket <> "" Then
	chkMicromarket = 1
Else
	chkMicromarket = 0
End If


If chkPantry <> "" Then
	chkPantry = 1
Else
	chkPantry = 0
End If


If chkOfficeSupplies <> "" Then
	chkOfficeSupplies = 1
Else
	chkOfficeSupplies = 0
End If


SQL = "UPDATE PR_Competitors SET "
SQL = SQL &  "CompetitorName = '" & CompetitorName & "', "
SQL = SQL &  "AddressInformation= '" & competitorAddressInfo & "', "
SQL = SQL &  "BottledWater= " & chkBottledWater & ", "
SQL = SQL &  "FilteredWater= " & chkFilteredWater & ", "
SQL = SQL &  "OCS= " & chkOfficeCoffee & ", "
SQL = SQL &  "OCS_Supply= " & chkOfficeCoffeeSupply & ", "
SQL = SQL &  "OfficeSupplies= " & chkOfficeSupplies & ", "
SQL = SQL &  "Vending= " & chkVending & ", "
SQL = SQL &  "Micromarket= " & chkMicromarket & ", "
SQL = SQL &  "Pantry= " & chkPantry 
SQL = SQL &  " WHERE InternalRecordIdentifier = " & InternalRecordIdentifier


Set cnn8 = Server.CreateObject("ADODB.Connection")
cnn8.open (Session("ClientCnnString"))

Set rs8 = Server.CreateObject("ADODB.Recordset")
rs8.CursorLocation = 3 
Set rs8 = cnn8.Execute(SQL)
set rs8 = Nothing


Description = ""
If Orig_Competitor  <> CompetitorName  Then
	Description = Description & GetTerm("Prospecting") & " Competitor changed from " & Orig_Competitor & " to " & CompetitorName
End If
If Orig_CompetitorAddressInfo  <> competitorAddressInfo Then
	Description = Description & GetTerm("Prospecting") & " Competitor address changed from " & Orig_CompetitorAddressInfo  & " to " & competitorAddressInfo & " for competitor " & CompetitorName
End If


If Orig_BottledWaterNote = 1 Then Orig_BottledWaterNote = "True" else Orig_BottledWaterNote = "False"
If BottledWaterNote = 1 Then BottledWaterNote = "True" else BottledWaterNote = "False"
If Orig_BottledWaterNote <> BottledWaterNote Then
	Description = Description & GetTerm("Prospecting") & " industry, bottled water, was changed from " & Orig_BottledWaterNote & " to " & BottledWaterNote & " for competitor " & CompetitorName
End If


If Orig_FilteredWaterNote = 1 Then Orig_FilteredWaterNote = "True" else Orig_FilteredWaterNote = "False"
If FilteredWaterNote = 1 Then FilteredWaterNote = "True" else FilteredWaterNote = "False"
If Orig_FilteredWaterNote <> FilteredWaterNote Then
	Description = Description & GetTerm("Prospecting") & " industry, filtered water, was changed from " & Orig_FilteredWaterNote & " to " & FilteredWaterNote & " for competitor " & CompetitorName
End If


If Orig_OCSNote = 1 Then Orig_OCSNote = "True" else Orig_OCSNote = "False"
If OCSNote = 1 Then OCSNote = "True" else OCSNote = "False"
If Orig_OCSNote <> OCSNote Then
	Description = Description & GetTerm("Prospecting") & " industry, office coffee, was changed from " & Orig_OCSNote & " to " & OCSNote & " for competitor " & CompetitorName
End If


If Orig_OCS_SupplyNote = 1 Then Orig_OCS_SupplyNote = "True" else Orig_OCS_SupplyNote = "False"
If OCS_SupplyNote = 1 Then OCS_SupplyNote = "True" else OCS_SupplyNote = "False"
If Orig_OCS_SupplyNote <> OCS_SupplyNote Then
	Description = Description & GetTerm("Prospecting") & " industry, office coffee supply, was changed from " & Orig_OCS_SupplyNote & " to " & OCS_SupplyNote & " for competitor " & CompetitorName
End If


If Orig_OfficeSuppliesNote = 1 Then Orig_OfficeSuppliesNote = "True" else Orig_OfficeSuppliesNote = "False"
If OfficeSuppliesNote = 1 Then OfficeSuppliesNote= "True" else OfficeSuppliesNote= "False"
If Orig_OfficeSuppliesNote <> OfficeSuppliesNote Then
	Description = Description & GetTerm("Prospecting") & " industry, office supplies, was changed from " & Orig_OfficeSuppliesNote & " to " & OfficeSuppliesNote  & " for competitor " & CompetitorName
End If


If Orig_VendingNote = 1 Then Orig_VendingNote = "True" else Orig_VendingNote = "False"
If VendingNote = 1 Then VendingNote = "True" else VendingNote = "False"
If Orig_VendingNote <> VendingNote Then
	Description = Description & GetTerm("Prospecting") & " industry, vending, was changed from " & Orig_VendingNote & " to " & VendingNote & " for competitor " & CompetitorName
End If


If Orig_MicromarketNote = 1 Then Orig_MicromarketNote = "True" else Orig_MicromarketNote = "False"
If MicromarketNote = 1 Then MicromarketNote = "True" else MicromarketNote = "False"
If Orig_MicromarketNote <> MicromarketNote Then
	Description = Description & GetTerm("Prospecting") & " industry, micromarket, was changed from " & Orig_MicromarketNote & " to " & MicromarketNote & " for competitor " & CompetitorName
End If


If Orig_PantryNote = 1 Then Orig_PantryNote = "True" else Orig_PantryNote = "False"
If PantryNote = 1 Then PantryNote = "True" else PantryNote = "False"
If Orig_PantryNote <> PantryNote Then
	Description = Description & GetTerm("Prospecting") & " industry, pantry, was changed from " & Orig_PantryNote & " to " & PantryNote & " for competitor " & CompetitorName
End If


CreateAuditLogEntry GetTerm("Prospecting") & " Competitor edited",GetTerm("Prospecting") & " Competitor edited","Minor",0,Description

Response.Redirect("main.asp")

%>















