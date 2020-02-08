<%		

	Set cnn_Settings_InventoryControl = Server.CreateObject("ADODB.Connection")
	cnn_Settings_InventoryControl.open (MUV_READ("ClientCnnString"))
	Set rs_Settings_InventoryControl = Server.CreateObject("ADODB.Recordset")
	rs_Settings_InventoryControl.CursorLocation = 3 
		
	'*******************************************************************************************************
	'Check to see if we have product backups for comparison
	'If we do, get the name of the most recent backup to compare IC_Product too
	'*******************************************************************************************************
	
	ICProductBackupTableName = ""
	SQLTableListing = ""
	SQLTableListingArray = ""
	
	SQL_Settings_InventoryControl = "SELECT * FROM INFORMATION_SCHEMA.TABLES WHERE TABLE_NAME LIKE 'IC_Product_%'"
	
	Set rs_Settings_InventoryControl = cnn_Settings_InventoryControl.Execute(SQL_Settings_InventoryControl)
	
	NoInventoryControlBackups = 0
	
	If rs_Settings_InventoryControl.EOF Then
		
		NoInventoryControlBackups = 1
		
	Else
		
		Do While NOT rs_Settings_InventoryControl.EOF
			SQLTableListing = SQLTableListing & rs_Settings_InventoryControl("TABLE_NAME") & ","
			rs_Settings_InventoryControl.MoveNext
		Loop
		
		If Right(SQLTableListing,1) = "," Then SQLTableListing = Left(SQLTableListing, Len(SQLTableListing) -1)
		
		SQLTableListingArray = Split(SQLTableListing,",")
		
		SQLTableListingArray = sortArray(SQLTableListingArray)
	
		If UBound(SQLTableListingArray) > 0 Then
			ICProductBackupTableName = SQLTableListingArray(UBound(SQLTableListingArray)-1)
		Else
			ICProductBackupTableName = SQLTableListingArray(UBound(SQLTableListingArray))
		End If
		
	End If
	'*******************************************************************************************************
	
	Set rs_Settings_InventoryControl = Nothing
	cnn_Settings_InventoryControl.Close
	Set cnn_Settings_InventoryControl = Nothing
	
	'*********************************************************************************************************************
	'GET THE COUNT OF TOTAL ROWS ADDED TO THE PRODUCTS TABLE TODAY, ALSO BUILD STRING OF PRODUCT ADDED
	'*********************************************************************************************************************	
	Set cnnProductChanges = Server.CreateObject("ADODB.Connection")
	cnnProductChanges.open (MUV_READ("ClientCnnString"))
	Set rsProductChanges  = Server.CreateObject("ADODB.Recordset")
	rsProductChanges.CursorLocation = 3 
	
	ICProductRowsAdded  = 0
	
	SQLChangeCounts = "SELECT Count(*) as totalDifferenceCount FROM "
	SQLChangeCounts = SQLChangeCounts & " (SELECT prodsku FROM IC_Product EXCEPT SELECT prodsku FROM " & ICProductBackupTableName & ") as T"
	rsProductChanges.Open SQLChangeCounts, cnnProductChanges
	
	'Response.Write("SQLChangeCounts:" & SQLChangeCounts& "<br>")
	
	ICProductRowsAdded = rsProductChanges("totalDifferenceCount")	
	
	Set rsProductChanges = Nothing
	cnnProductChanges.Close
	Set cnnProductChanges = Nothing
	
	addedProductsString = ""
	
	If ICProductRowsAdded > 0 Then
	
		'**************************************************************
		'BUILD COMMA SEPARATED STRING OF PRODUCT SKUS ADDED TODAY
		'**************************************************************
	
		SQLChangeCounts = "SELECT prodSKU, prodDescription FROM IC_Product WHERE prodSKU IN "
		SQLChangeCounts = SQLChangeCounts & " (SELECT  prodSKU "
		SQLChangeCounts = SQLChangeCounts & " FROM IC_Product "
		SQLChangeCounts = SQLChangeCounts & " EXCEPT "
		SQLChangeCounts = SQLChangeCounts & " SELECT prodSKU "
		SQLChangeCounts = SQLChangeCounts & " FROM  " & ICProductBackupTableName & ") "
	
		Set cnnProductChanges = Server.CreateObject("ADODB.Connection")
		cnnProductChanges.open (MUV_READ("ClientCnnString"))
		Set rsProductChanges = Server.CreateObject("ADODB.Recordset")
		rsProductChanges.CursorLocation = 3 
		rsProductChanges.Open SQLChangeCounts, cnnProductChanges
		
		If Not rsProductChanges.EOF Then
			
			Do While Not rsProductChanges.EOF
			
				'***************************************************************
				'CREATE AN ARRAY OF ADDED PRODUCTS
				'THIS ARRAY WILL SERVE AS A COMPARISON FOR LOWER CHANGES SO WE
				'DONT'T REPORT ADDED PRODUCTS AS BEING CHANGED
				'***************************************************************
				
				addedProductsString = addedProductsString &  "'" & rsProductChanges("prodSKU") & "',"

				rsProductChanges.Movenext
					
			Loop
		End If
		
		If Right(addedProductsString,1) = "," Then addedProductsString = Left(addedProductsString, Len(addedProductsString) -1)
		
	End If
	
	If Len(addedProductsString) = 0 Then
		addedProductsString = "'NO PRODUCTS ADDED'"
	End If
	
	'*********************************************************************************************************************
	
	
	'*********************************************************************************************************************
	'GET THE COUNT OF TOTAL ROWS DELETED FROM THE PRODUCTS TABLE TODAY
	'*********************************************************************************************************************
	Set cnnProductChanges = Server.CreateObject("ADODB.Connection")
	cnnProductChanges.open (MUV_READ("ClientCnnString"))
	Set rsProductChanges  = Server.CreateObject("ADODB.Recordset")
	rsProductChanges.CursorLocation = 3 
	
	ICProductRowsDeleted = 0 
	
	SQLChangeCounts = "SELECT Count(*) as totalDifferenceCount FROM "
	SQLChangeCounts = SQLChangeCounts & " (SELECT prodsku FROM " & ICProductBackupTableName & " EXCEPT SELECT prodsku FROM IC_Product) as T"
	rsProductChanges.Open SQLChangeCounts, cnnProductChanges
	
	ICProductRowsDeleted = rsProductChanges("totalDifferenceCount")
	
	Set rsProductChanges = Nothing
	cnnProductChanges.Close
	Set cnnProductChanges = Nothing
	'*********************************************************************************************************************
	

	'*********************************************************************************************************************
	'GET THE COUNT OF TOTAL PRODUCTS WITH MODIFIED PRODUCT UNIT DESCRIPTIONS TODAY
	'*********************************************************************************************************************	
	Set cnnProductChanges = Server.CreateObject("ADODB.Connection")
	cnnProductChanges.open (MUV_READ("ClientCnnString"))
	Set rsProductChanges  = Server.CreateObject("ADODB.Recordset")
	rsProductChanges.CursorLocation = 3 
	
	ICProductUnitDescriptionRowChanged = 0 
	
	SQLChangeCounts = "SELECT COUNT(prodSKU) as totalDifferenceCount FROM IC_Product "
	SQLChangeCounts = SQLChangeCounts & " WHERE prodSKU IN "
	SQLChangeCounts = SQLChangeCounts & " (SELECT prodSKU from "
	SQLChangeCounts = SQLChangeCounts & " (SELECT prodSKU,prodDescription FROM IC_Product "
	SQLChangeCounts = SQLChangeCounts & " EXCEPT "
	SQLChangeCounts = SQLChangeCounts & " SELECT prodSKU,prodDescription FROM " & ICProductBackupTableName & ") as t) "			
	SQLChangeCounts = SQLChangeCounts & " AND prodSKU NOT IN (" & addedProductsString & ")"
	
	'Response.Write(SQLChangeCounts)
	
	rsProductChanges.Open SQLChangeCounts, cnnProductChanges
	
	ICProductUnitDescriptionRowChanged = rsProductChanges("totalDifferenceCount")	
	
	Set rsProductChanges = Nothing
	cnnProductChanges.Close
	Set cnnProductChanges = Nothing
	'*********************************************************************************************************************
	

	'*********************************************************************************************************************
	'GET THE COUNT OF TOTAL PRODUCTS WITH MODIFIED PRODUCT CASE DESCRIPTIONS TODAY
	'*********************************************************************************************************************	
	Set cnnProductChanges = Server.CreateObject("ADODB.Connection")
	cnnProductChanges.open (MUV_READ("ClientCnnString"))
	Set rsProductChanges  = Server.CreateObject("ADODB.Recordset")
	rsProductChanges.CursorLocation = 3 
	
	ICProductCaseDescriptionRowsChanged = 0 
	
	SQLChangeCounts = "SELECT COUNT(prodSKU) as totalDifferenceCount FROM IC_Product "
	SQLChangeCounts = SQLChangeCounts & " WHERE prodSKU IN "
	SQLChangeCounts = SQLChangeCounts & " (SELECT prodSKU from "
	SQLChangeCounts = SQLChangeCounts & " (SELECT prodSKU,prodCaseDescription FROM IC_Product "
	SQLChangeCounts = SQLChangeCounts & " EXCEPT "
	SQLChangeCounts = SQLChangeCounts & " SELECT prodSKU,prodCaseDescription FROM " & ICProductBackupTableName & ") as t) "			
	SQLChangeCounts = SQLChangeCounts & " AND prodSKU NOT IN (" & addedProductsString & ")"
	
	rsProductChanges.Open SQLChangeCounts, cnnProductChanges
	
	ICProductCaseDescriptionRowsChanged = rsProductChanges("totalDifferenceCount")	
	
	Set rsProductChanges = Nothing
	cnnProductChanges.Close
	Set cnnProductChanges = Nothing

	'*********************************************************************************************************************
	

	'*********************************************************************************************************************
	'GET THE COUNT OF TOTAL PRODUCTS WITH MODIFIED PRODUCT UNIT COSTS TODAY
	'*********************************************************************************************************************	
	Set cnnProductChanges = Server.CreateObject("ADODB.Connection")
	cnnProductChanges.open (MUV_READ("ClientCnnString"))
	Set rsProductChanges  = Server.CreateObject("ADODB.Recordset")
	rsProductChanges.CursorLocation = 3 
	
	ICProductUnitCostRowsChanged = 0 
	
	SQLChangeCounts = "SELECT COUNT(prodSKU) as totalDifferenceCount FROM IC_Product "
	SQLChangeCounts = SQLChangeCounts & " WHERE prodSKU IN "
	SQLChangeCounts = SQLChangeCounts & " (SELECT prodSKU from "
	SQLChangeCounts = SQLChangeCounts & " (SELECT prodSKU,prodUnitCost FROM IC_Product "
	SQLChangeCounts = SQLChangeCounts & " EXCEPT "
	SQLChangeCounts = SQLChangeCounts & " SELECT prodSKU,prodUnitCost FROM " & ICProductBackupTableName & ") as t) "			
	SQLChangeCounts = SQLChangeCounts & " AND prodSKU NOT IN (" & addedProductsString & ")"
			
	rsProductChanges.Open SQLChangeCounts, cnnProductChanges
	
	ICProductUnitCostRowsChanged = rsProductChanges("totalDifferenceCount")	
	
	Set rsProductChanges = Nothing
	cnnProductChanges.Close
	Set cnnProductChanges = Nothing		
	'*********************************************************************************************************************

	

	'*********************************************************************************************************************
	'GET THE COUNT OF TOTAL PRODUCTS WITH MODIFIED PRODUCT UNIT PRICES TODAY
	'*********************************************************************************************************************	
	Set cnnProductChanges = Server.CreateObject("ADODB.Connection")
	cnnProductChanges.open (MUV_READ("ClientCnnString"))
	Set rsProductChanges  = Server.CreateObject("ADODB.Recordset")
	rsProductChanges.CursorLocation = 3 
	
	ICProductUnitPriceRowsChanged = 0 
	
	SQLChangeCounts = "SELECT COUNT(prodSKU) as totalDifferenceCount FROM IC_Product "
	SQLChangeCounts = SQLChangeCounts & " WHERE prodSKU IN "
	SQLChangeCounts = SQLChangeCounts & " (SELECT prodSKU from "
	SQLChangeCounts = SQLChangeCounts & " (SELECT prodSKU,prodPriceLvl1 FROM IC_Product "
	SQLChangeCounts = SQLChangeCounts & " EXCEPT "
	SQLChangeCounts = SQLChangeCounts & " SELECT prodSKU,prodPriceLvl1 FROM " & ICProductBackupTableName & ") as t) "			
	SQLChangeCounts = SQLChangeCounts & " AND prodSKU NOT IN (" & addedProductsString & ")"
	
	rsProductChanges.Open SQLChangeCounts, cnnProductChanges
	
	ICProductUnitPriceRowsChanged = rsProductChanges("totalDifferenceCount")	
	
	Set rsProductChanges = Nothing
	cnnProductChanges.Close
	Set cnnProductChanges = Nothing		
	'*********************************************************************************************************************


	'*********************************************************************************************************************
	'GET THE COUNT OF TOTAL PRODUCTS WITH MODIFIED PRODUCT CASE PRICES TODAY
	'*********************************************************************************************************************	
	Set cnnProductChanges = Server.CreateObject("ADODB.Connection")
	cnnProductChanges.open (MUV_READ("ClientCnnString"))
	Set rsProductChanges  = Server.CreateObject("ADODB.Recordset")
	rsProductChanges.CursorLocation = 3 
	
	ICProductCasePricingRowsChanged = 0 
	
	SQLChangeCounts = "SELECT COUNT(prodSKU) as totalDifferenceCount FROM IC_Product "
	SQLChangeCounts = SQLChangeCounts & " WHERE prodSKU IN "
	SQLChangeCounts = SQLChangeCounts & " (SELECT prodSKU from "
	SQLChangeCounts = SQLChangeCounts & " (SELECT prodSKU,prodCasePricing FROM IC_Product "
	SQLChangeCounts = SQLChangeCounts & " EXCEPT "
	SQLChangeCounts = SQLChangeCounts & " SELECT prodSKU,prodCasePricing FROM " & ICProductBackupTableName & ") as t) "			
	SQLChangeCounts = SQLChangeCounts & " AND prodSKU NOT IN (" & addedProductsString & ")"
	
	rsProductChanges.Open SQLChangeCounts, cnnProductChanges
	
	ICProductCasePricingRowsChanged = rsProductChanges("totalDifferenceCount")	
	
	Set rsProductChanges = Nothing
	cnnProductChanges.Close
	Set cnnProductChanges = Nothing			
	'*********************************************************************************************************************



	'*********************************************************************************************************************
	'GET THE COUNT OF TOTAL PRODUCTS WITH MODIFIED PRODUCT UNIT BINS TODAY
	'*********************************************************************************************************************	
	Set cnnProductChanges = Server.CreateObject("ADODB.Connection")
	cnnProductChanges.open (MUV_READ("ClientCnnString"))
	Set rsProductChanges  = Server.CreateObject("ADODB.Recordset")
	rsProductChanges.CursorLocation = 3 
	
	ICProductBinNoRowsChanged = 0 
	
	SQLChangeCounts = "SELECT COUNT(prodSKU) as totalDifferenceCount FROM IC_Product "
	SQLChangeCounts = SQLChangeCounts & " WHERE prodSKU IN "
	SQLChangeCounts = SQLChangeCounts & " (SELECT prodSKU from "
	SQLChangeCounts = SQLChangeCounts & " (SELECT prodSKU,prodUnitBin FROM IC_Product "
	SQLChangeCounts = SQLChangeCounts & " EXCEPT "
	SQLChangeCounts = SQLChangeCounts & " SELECT prodSKU,prodUnitBin FROM " & ICProductBackupTableName & ") as t) "			
	SQLChangeCounts = SQLChangeCounts & " AND prodSKU NOT IN (" & addedProductsString & ")"
		
	rsProductChanges.Open SQLChangeCounts, cnnProductChanges
	
	ICProductBinNoRowsChanged = rsProductChanges("totalDifferenceCount")	
	
	Set rsProductChanges = Nothing
	cnnProductChanges.Close
	Set cnnProductChanges = Nothing				
	'*********************************************************************************************************************


	'*********************************************************************************************************************
	'GET THE COUNT OF TOTAL PRODUCTS WITH MODIFIED PRODUCT CASE BINS TODAY
	'*********************************************************************************************************************	
	Set cnnProductChanges = Server.CreateObject("ADODB.Connection")
	cnnProductChanges.open (MUV_READ("ClientCnnString"))
	Set rsProductChanges  = Server.CreateObject("ADODB.Recordset")
	rsProductChanges.CursorLocation = 3 
	
	ICProductBinCaseRowsChanged = 0 
	
	SQLChangeCounts = "SELECT COUNT(prodSKU) as totalDifferenceCount FROM IC_Product "
	SQLChangeCounts = SQLChangeCounts & " WHERE prodSKU IN "
	SQLChangeCounts = SQLChangeCounts & " (SELECT prodSKU from "
	SQLChangeCounts = SQLChangeCounts & " (SELECT prodSKU,prodCaseBin FROM IC_Product "
	SQLChangeCounts = SQLChangeCounts & " EXCEPT "
	SQLChangeCounts = SQLChangeCounts & " SELECT prodSKU,prodCaseBin FROM " & ICProductBackupTableName & ") as t) "			
	SQLChangeCounts = SQLChangeCounts & " AND prodSKU NOT IN (" & addedProductsString & ")"
	
	rsProductChanges.Open SQLChangeCounts, cnnProductChanges
	
	ICProductBinCaseRowsChanged = rsProductChanges("totalDifferenceCount")	
	
	Set rsProductChanges = Nothing
	cnnProductChanges.Close
	Set cnnProductChanges = Nothing				
	'*********************************************************************************************************************


	'*********************************************************************************************************************
	'GET THE COUNT OF TOTAL PRODUCTS WITH MODIFIED PRODUCT PERPETUAL FLAGS TODAY
	'*********************************************************************************************************************	
	Set cnnProductChanges = Server.CreateObject("ADODB.Connection")
	cnnProductChanges.open (MUV_READ("ClientCnnString"))
	Set rsProductChanges  = Server.CreateObject("ADODB.Recordset")
	rsProductChanges.CursorLocation = 3 
	
	ICProductPerpetualFlagRowsChanged = 0 
	
	SQLChangeCounts = "SELECT COUNT(prodSKU) as totalDifferenceCount FROM IC_Product "
	SQLChangeCounts = SQLChangeCounts & " WHERE prodSKU IN "
	SQLChangeCounts = SQLChangeCounts & " (SELECT prodSKU from "
	SQLChangeCounts = SQLChangeCounts & " (SELECT prodSKU,DisplayOnWeb FROM IC_Product "
	SQLChangeCounts = SQLChangeCounts & " EXCEPT "
	SQLChangeCounts = SQLChangeCounts & " SELECT prodSKU,DisplayOnWeb FROM " & ICProductBackupTableName & ") as t) "			
	SQLChangeCounts = SQLChangeCounts & " AND prodSKU NOT IN (" & addedProductsString & ")"
	
	rsProductChanges.Open SQLChangeCounts, cnnProductChanges
	
	ICProductPerpetualFlagRowsChanged = rsProductChanges("totalDifferenceCount")	
	
	Set rsProductChanges = Nothing
	cnnProductChanges.Close
	Set cnnProductChanges = Nothing				
	'*********************************************************************************************************************


	'*********************************************************************************************************************
	'GET THE COUNT OF TOTAL PRODUCTS WITH MODIFIED PRODUCT INVENTORIED FLAGS TODAY
	'*********************************************************************************************************************	
	Set cnnProductChanges = Server.CreateObject("ADODB.Connection")
	cnnProductChanges.open (MUV_READ("ClientCnnString"))
	Set rsProductChanges  = Server.CreateObject("ADODB.Recordset")
	rsProductChanges.CursorLocation = 3 
	
	ICProductInventoriedRowsChanged = 0 
	
	SQLChangeCounts = "SELECT COUNT(prodSKU) as totalDifferenceCount FROM IC_Product "
	SQLChangeCounts = SQLChangeCounts & " WHERE prodSKU IN "
	SQLChangeCounts = SQLChangeCounts & " (SELECT prodSKU from "
	SQLChangeCounts = SQLChangeCounts & " (SELECT prodSKU,prodInventoriedItem FROM IC_Product "
	SQLChangeCounts = SQLChangeCounts & " EXCEPT "
	SQLChangeCounts = SQLChangeCounts & " SELECT prodSKU,prodInventoriedItem FROM " & ICProductBackupTableName & ") as t) "			
	SQLChangeCounts = SQLChangeCounts & " AND prodSKU NOT IN (" & addedProductsString & ")"
	
	rsProductChanges.Open SQLChangeCounts, cnnProductChanges
	
	ICProductInventoriedRowsChanged = rsProductChanges("totalDifferenceCount")	
	
	Set rsProductChanges = Nothing
	cnnProductChanges.Close
	Set cnnProductChanges = Nothing						
	'*********************************************************************************************************************


	'*********************************************************************************************************************
	'GET THE COUNT OF TOTAL PRODUCTS WITH MODIFIED PRODUCT PICKABLE FLAGS TODAY
	'*********************************************************************************************************************	
	Set cnnProductChanges = Server.CreateObject("ADODB.Connection")
	cnnProductChanges.open (MUV_READ("ClientCnnString"))
	Set rsProductChanges  = Server.CreateObject("ADODB.Recordset")
	rsProductChanges.CursorLocation = 3 
	
	ICProductPickableRowsChanged = 0 
	
	SQLChangeCounts = "SELECT COUNT(prodSKU) as totalDifferenceCount FROM IC_Product "
	SQLChangeCounts = SQLChangeCounts & " WHERE prodSKU IN "
	SQLChangeCounts = SQLChangeCounts & " (SELECT prodSKU from "
	SQLChangeCounts = SQLChangeCounts & " (SELECT prodSKU,prodPickableItem FROM IC_Product "
	SQLChangeCounts = SQLChangeCounts & " EXCEPT "
	SQLChangeCounts = SQLChangeCounts & " SELECT prodSKU,prodPickableItem FROM " & ICProductBackupTableName & ") as t) "			
	SQLChangeCounts = SQLChangeCounts & " AND prodSKU NOT IN (" & addedProductsString & ")"
	
	rsProductChanges.Open SQLChangeCounts, cnnProductChanges
	
	ICProductPickableRowsChanged = rsProductChanges("totalDifferenceCount")	
	
	Set rsProductChanges = Nothing
	cnnProductChanges.Close
	Set cnnProductChanges = Nothing			
	'*********************************************************************************************************************
	
	'Response.Write(SQLChangeCounts & "<br>")
	
%>
