<%

UnitUPCDataFilter = Request.Form("optUnitUPCData")
CaseUPCDataFilter = Request.Form("optCaseUPCData")
InventoriedItemDataFilter = Request.Form("optInventoriedItem")
PickableItemDataFilter = Request.Form("optPickableItem")
ProductCategoriesForInventoryReport = Request.Form("chkCategoryNum")

Response.Write("UnitUPCDataFilter " & UnitUPCDataFilter & "<br>")
Response.Write("CaseUPCDataFilter " & CaseUPCDataFilter & "<br>")
Response.Write("InventoriedItemDataFilter " & InventoriedItemDataFilter & "<br>")
Response.Write("PickableItemDataFilter " & PickableItemDataFilter & "<br>")
Response.Write("ProductCategoriesForInventoryReport " & ProductCategoriesForInventoryReport & "<br>")

If Right(ProductCategoriesForInventoryReport,1) = "," Then 
	ProductCategoriesForInventoryReport = left(ProductCategoriesForInventoryReport,Len(ProductCategoriesForInventoryReport)-1)
End If

ProductCategoriesArrayForCustomize = ""
ProductCategoriesArrayForCustomize = Split(ProductCategoriesForInventoryReport,",")

For z = 0 to UBound(ProductCategoriesArrayForCustomize)
	If z = 0 Then
		ProductCategoriesForInventoryReport = Trim(ProductCategoriesArrayForCustomize(z))
	Else
		ProductCategoriesForInventoryReport = ProductCategoriesForInventoryReport & "," & Trim(ProductCategoriesArrayForCustomize(z))
	End If
Next	




SQL = "SELECT * FROM Settings_Reports WHERE ReportNumber = 1600 AND UserNo = " & Session("userNo")

Set cnn8 = Server.CreateObject("ADODB.Connection")
cnn8.open (Session("ClientCnnString"))
Set rs = Server.CreateObject("ADODB.Recordset")
Set rs= cnn8.Execute(SQL)

'Rec does not exist yet, make it quick but empty, update it later
If rs.EOF Then
	SQL = "INSERT INTO Settings_Reports (ReportNumber, UserNo) VALUES (1600, " & Session("userNo") & ")"
	rs.Close
	Set rs= cnn8.Execute(SQL)
End If

'Now update the table with the values


SQL = "UPDATE Settings_Reports Set ReportSpecificData1 = '" & UnitUPCDataFilter & "', "
SQL = SQL & "ReportSpecificData2 = '" & CaseUPCDataFilter & "', "
SQL = SQL & "ReportSpecificData3 = '" & InventoriedItemDataFilter & "', " 
SQL = SQL & "ReportSpecificData4 = '" & PickableItemDataFilter & "', " 
SQL = SQL & "ReportSpecificData5 = '" & ProductCategoriesForInventoryReport & "' " 
SQL = SQL & "WHERE ReportNumber = 1600 AND UserNo = " & Session("userNo")

Set rs= cnn8.Execute(SQL)

cnn8.Close

Set rs = Nothing
Set cnn8 = Nothing

Response.Redirect ("ProductInventoryReport.asp")
%>

 
