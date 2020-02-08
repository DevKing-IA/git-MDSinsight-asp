<%	
	Set cnnCheckICProduct = Server.CreateObject("ADODB.Connection")
	cnnCheckICProduct.open (Session("ClientCnnString"))
	Set rsCheckICProduct = Server.CreateObject("ADODB.Recordset")
	rsCheckICProduct.CursorLocation = 3 
	
	SQL_CheckICProduct = "SELECT COL_LENGTH('IC_Product', 'prodUnitBin') AS IsItThere"
	Set rsCheckICProduct = cnnCheckICProduct.Execute(SQL_CheckICProduct)
	If IsNull(rsCheckICProduct("IsItThere")) Then
		SQL_CheckICProduct = "ALTER TABLE IC_Product ADD prodUnitBin varchar(255) NULL"
		Set rsCheckICProduct = cnnCheckICProduct.Execute(SQL_CheckICProduct)
	End If
	
	SQL_CheckICProduct = "SELECT COL_LENGTH('IC_Product', 'prodCaseBin') AS IsItThere"
	Set rsCheckICProduct = cnnCheckICProduct.Execute(SQL_CheckICProduct)
	If IsNull(rsCheckICProduct("IsItThere")) Then
		SQL_CheckICProduct = "ALTER TABLE IC_Product ADD prodCaseBin varchar(255) NULL"
		Set rsCheckICProduct = cnnCheckICProduct.Execute(SQL_CheckICProduct)
	End If

	SQL_CheckICProduct = "SELECT COL_LENGTH('IC_Product', 'QtyOnHand_Units') AS IsItThere"
	Set rsCheckICProduct = cnnCheckICProduct.Execute(SQL_CheckICProduct)
	If IsNull(rsCheckICProduct("IsItThere")) Then
		SQL_CheckICProduct = "ALTER TABLE IC_Product ADD QtyOnHand_Units int NULL"
		Set rsCheckICProduct = cnnCheckICProduct.Execute(SQL_CheckICProduct)
	End If

	SQL_CheckICProduct = "SELECT COL_LENGTH('IC_Product', 'QtyOnHand_LastUpdated') AS IsItThere"
	Set rsCheckICProduct = cnnCheckICProduct.Execute(SQL_CheckICProduct)
	If IsNull(rsCheckICProduct("IsItThere")) Then
		SQL_CheckICProduct = "ALTER TABLE IC_Product ADD QtyOnHand_LastUpdated datetime NULL"
		Set rsCheckICProduct = cnnCheckICProduct.Execute(SQL_CheckICProduct)
		SQL_CheckICProduct = "ALTER TABLE [IC_Product] ADD CONSTRAINT [DF_IC_Product_QOHUpdate]  DEFAULT (getdate()) FOR [QtyOnHand_LastUpdated]"
		Set rsCheckICProduct = cnnCheckICProduct.Execute(SQL_CheckICProduct)	
		SQL_CheckICProduct = "UPDATE IC_Product SET QtyOnHand_LastUpdated = getdate()"
		Set rsCheckICProduct = cnnCheckICProduct.Execute(SQL_CheckICProduct)	
	End If

	SQL_CheckICProduct = "SELECT COL_LENGTH('IC_Product', 'DisplayOnWeb') AS IsItThere"
	Set rsCheckICProduct = cnnCheckICProduct.Execute(SQL_CheckICProduct)
	If IsNull(rsCheckICProduct("IsItThere")) Then
		SQL_CheckICProduct = "ALTER TABLE IC_Product ADD DisplayOnWeb int NOT NULL DEFAULT 1"
		Set rsCheckICProduct = cnnCheckICProduct.Execute(SQL_CheckICProduct)
		SQL_CheckICProduct = "UPDATE IC_Product SET DisplayOnWeb = 1"
		Set rsCheckICProduct = cnnCheckICProduct.Execute(SQL_CheckICProduct)	
	End If
	
	SQL_CheckICProduct = "SELECT COL_LENGTH('IC_Product', 'WebShortDescription') AS IsItThere"
	Set rsCheckICProduct = cnnCheckICProduct.Execute(SQL_CheckICProduct)
	If IsNull(rsCheckICProduct("IsItThere")) Then
		SQL_CheckICProduct = "ALTER TABLE IC_Product ADD WebShortDescription varchar(1000) NULL"
		Set rsCheckICProduct = cnnCheckICProduct.Execute(SQL_CheckICProduct)
	End If
	
	SQL_CheckICProduct = "SELECT COL_LENGTH('IC_Product', 'WebLongDescription') AS IsItThere"
	Set rsCheckICProduct = cnnCheckICProduct.Execute(SQL_CheckICProduct)
	If IsNull(rsCheckICProduct("IsItThere")) Then
		SQL_CheckICProduct = "ALTER TABLE IC_Product ADD WebLongDescription varchar(8000) NULL"
		Set rsCheckICProduct = cnnCheckICProduct.Execute(SQL_CheckICProduct)
	End If
	
	SQL_CheckICProduct = "SELECT COL_LENGTH('IC_Product', 'OutOfStock') AS IsItThere"
	Set rsCheckICProduct = cnnCheckICProduct.Execute(SQL_CheckICProduct)
	If IsNull(rsCheckICProduct("IsItThere")) Then
		SQL_CheckICProduct = "ALTER TABLE IC_Product ADD OutOfStock bit NOT NULL DEFAULT 0"
		Set rsCheckICProduct = cnnCheckICProduct.Execute(SQL_CheckICProduct)
	End If

	SQL_CheckICProduct = "SELECT COL_LENGTH('IC_Product', 'OutOfStockMessage') AS IsItThere"
	Set rsCheckICProduct = cnnCheckICProduct.Execute(SQL_CheckICProduct)
	If IsNull(rsCheckICProduct("IsItThere")) Then
		SQL_CheckICProduct = "ALTER TABLE IC_Product ADD OutOfStockMessage varchar(1000) NULL"
		Set rsCheckICProduct = cnnCheckICProduct.Execute(SQL_CheckICProduct)
	End If
	
	SQL_CheckICProduct = "SELECT COL_LENGTH('IC_Product', 'HideIfNotLoggedIn') AS IsItThere"
	Set rsCheckICProduct = cnnCheckICProduct.Execute(SQL_CheckICProduct)
	If IsNull(rsCheckICProduct("IsItThere")) Then
		SQL_CheckICProduct = "ALTER TABLE IC_Product ADD HideIfNotLoggedIn bit NULL"
		Set rsCheckICProduct = cnnCheckICProduct.Execute(SQL_CheckICProduct)
		SQL_CheckICProduct = "UPDATE IC_Product SET HideIfNotLoggedIn = 1"
		Set rsCheckICProduct = cnnCheckICProduct.Execute(SQL_CheckICProduct)		
	End If
	
	SQL_CheckICProduct = "SELECT COL_LENGTH('IC_Product', 'Taxable') AS IsItThere"
	Set rsCheckICProduct = cnnCheckICProduct.Execute(SQL_CheckICProduct)
	If IsNull(rsCheckICProduct("IsItThere")) Then
		SQL_CheckICProduct = "ALTER TABLE IC_Product ADD Taxable varchar(50) NULL"
		Set rsCheckICProduct = cnnCheckICProduct.Execute(SQL_CheckICProduct)
	End If
	
	
' This one is a DROP
	SQLCheckICProduct = "SELECT COL_LENGTH('IC_Product', 'WebUnitShortDescription') AS IsItThere"
	Set rsCheckICProduct  = cnnCheckICProduct.Execute(SQLCheckICProduct)
	If NOT IsNull(rsCheckICProduct("IsItThere")) Then
		SQLCheckICProduct = "ALTER TABLE IC_Product DROP COLUMN WebUnitShortDescription"
		Set rsCheckICProduct = cnnCheckICProduct.Execute(SQLCheckICProduct)
	End If
	
' This one is a DROP
	SQLCheckICProduct = "SELECT COL_LENGTH('IC_Product', 'WebCaseShortDescription') AS IsItThere"
	Set rsCheckICProduct  = cnnCheckICProduct.Execute(SQLCheckICProduct)
	If NOT IsNull(rsCheckICProduct("IsItThere")) Then
		SQLCheckICProduct = "ALTER TABLE IC_Product DROP COLUMN WebCaseShortDescription"
		Set rsCheckICProduct = cnnCheckICProduct.Execute(SQLCheckICProduct)
	End If

' This one is a DROP
	SQLCheckICProduct = "SELECT COL_LENGTH('IC_Product', 'WebUnitLongDescription') AS IsItThere"
	Set rsCheckICProduct  = cnnCheckICProduct.Execute(SQLCheckICProduct)
	If NOT IsNull(rsCheckICProduct("IsItThere")) Then
		SQLCheckICProduct = "ALTER TABLE IC_Product DROP COLUMN WebUnitLongDescription"
		Set rsCheckICProduct = cnnCheckICProduct.Execute(SQLCheckICProduct)
	End If
	
' This one is a DROP
	SQLCheckICProduct = "SELECT COL_LENGTH('IC_Product', 'WebCaseLongDescription') AS IsItThere"
	Set rsCheckICProduct  = cnnCheckICProduct.Execute(SQLCheckICProduct)
	If NOT IsNull(rsCheckICProduct("IsItThere")) Then
		SQLCheckICProduct = "ALTER TABLE IC_Product DROP COLUMN WebCaseLongDescription"
		Set rsCheckICProduct = cnnCheckICProduct.Execute(SQLCheckICProduct)
	End If


	SQL_CheckICProduct = "SELECT COL_LENGTH('IC_Product', 'ProductIsAFilter') AS IsItThere"
	Set rsCheckICProduct = cnnCheckICProduct.Execute(SQL_CheckICProduct)
	If IsNull(rsCheckICProduct("IsItThere")) Then
		SQL_CheckICProduct = "ALTER TABLE IC_Product ADD ProductIsAFilter int NULL"
		Set rsCheckICProduct = cnnCheckICProduct.Execute(SQL_CheckICProduct)
		SQL_CheckICProduct = "UPDATE IC_Product SET ProductIsAFilter = 0"
		Set rsCheckICProduct = cnnCheckICProduct.Execute(SQL_CheckICProduct)		
	End If

	Set rsCheckICProduct = Nothing
	cnnCheckICProduct.Close
	Set cnnCheckICProduct = Nothing
%>