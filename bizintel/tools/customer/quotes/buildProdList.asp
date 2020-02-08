<%

Set objConnection = CreateObject("ADODB.Connection")
Set objRecordSet = CreateObject("ADODB.Recordset")
objRecordSet.CursorLocation = 3
objConnection.Open (Session("ClientCnnString"))
objSQL = "SELECT * FROM sysobjects Where Name= 'zPRC_AccountQuotedItems_" & trim(Session("Userno")) & "' AND xType= 'U'"
objRecordSet.Open objSQL,objConnection

If objRecordset.RecordCount = 0  Then

    'Response.Write("The table is not in the database. Create the table.")

	Set cnnBuildProdList  = Server.CreateObject("ADODB.Connection")
	cnnBuildProdList.open (Session("ClientCnnString"))
	Set rsBuildProdList = Server.CreateObject("ADODB.Recordset")
	rsBuildProdList.CursorLocation = 3 
	Set rsBuildProdList2 = Server.CreateObject("ADODB.Recordset")
	rsBuildProdList2.CursorLocation = 3 
	
	'Before we build anything we need to see if the indexes exist in the Product
	'table and if not, create them to make this process faster
	SQLBuildProdList = "SELECT * FROM sys.indexes WHERE name = 'ProductIndex1' AND object_id = OBJECT_ID('Product')" 
	Set rsBuildProdList = cnnBuildProdList.Execute(SQLBuildProdList)
	If rsBuildProdList.Eof Then ' Build It
		SQLBuildProdList = "CREATE NONCLUSTERED INDEX [ProductIndex1] ON [Product] "
		SQLBuildProdList = SQLBuildProdList & "( "
		SQLBuildProdList = SQLBuildProdList & "[Category] ASC "
		SQLBuildProdList = SQLBuildProdList & ")WITH (PAD_INDEX = OFF, STATISTICS_NORECOMPUTE = OFF, SORT_IN_TEMPDB = OFF, DROP_EXISTING = OFF, ONLINE = OFF, ALLOW_ROW_LOCKS = ON, ALLOW_PAGE_LOCKS = ON) ON [PRIMARY]"
		'Response.Write(SQLBuildProdList & "<BR>")
		Set rsBuildProdList = cnnBuildProdList.Execute(SQLBuildProdList)
	End IF
	
	SQLBuildProdList = "SELECT * FROM sys.indexes WHERE name = 'ProductIndex2' AND object_id = OBJECT_ID('Product')" 
	Set rsBuildProdList = cnnBuildProdList.Execute(SQLBuildProdList)
	If rsBuildProdList.Eof Then ' Build It
		SQLBuildProdList = "CREATE NONCLUSTERED INDEX [ProductIndex2] ON [Product] "
		SQLBuildProdList = SQLBuildProdList & "( "
		SQLBuildProdList = SQLBuildProdList & "[PartNo] ASC, [Description] ASC, [ListPriceLevel1] ASC, [UnitCost] ASC, [CaseConversionFactor] ASC, [CasePricing] ASC, [CaseDescription] ASC "
		SQLBuildProdList = SQLBuildProdList & ")WITH (PAD_INDEX = OFF, STATISTICS_NORECOMPUTE = OFF, SORT_IN_TEMPDB = OFF, DROP_EXISTING = OFF, ONLINE = OFF, ALLOW_ROW_LOCKS = ON, ALLOW_PAGE_LOCKS = ON) ON [PRIMARY]"
		'Response.Write(SQLBuildProdList & "<BR>")
		Set rsBuildProdList = cnnBuildProdList.Execute(SQLBuildProdList)
	End IF
	
	'Also need to index the temp quoted items table
	SQLBuildProdList = "SELECT * FROM sys.indexes WHERE name = 'Index1' AND object_id = OBJECT_ID('zPRC_AccountQuotedItems_" & trim(Session("Userno")) & "')" 
	Set rsBuildProdList = cnnBuildProdList.Execute(SQLBuildProdList)
	If rsBuildProdList.Eof Then ' Build It
		SQLBuildProdList = "CREATE NONCLUSTERED INDEX [Index1] ON [zPRC_AccountQuotedItems_" & trim(Session("Userno")) & "] "
		SQLBuildProdList = SQLBuildProdList & "( "
		SQLBuildProdList = SQLBuildProdList & "[prodSKU] ASC, [QuoteType] ASC "
		SQLBuildProdList = SQLBuildProdList & ")WITH (PAD_INDEX = OFF, STATISTICS_NORECOMPUTE = OFF, SORT_IN_TEMPDB = OFF, DROP_EXISTING = OFF, ONLINE = OFF, ALLOW_ROW_LOCKS = ON, ALLOW_PAGE_LOCKS = ON) ON [PRIMARY]"
		'Response.Write(SQLBuildProdList & "<BR>")
		Set rsBuildProdList = cnnBuildProdList.Execute(SQLBuildProdList)
	End IF
	
	
	
	On Error Resume Next ' In caase the table isn't there
	SQLBuildProdList = "DROP TABLE zPRC_AccountQuotedItems_ProdList_" & trim(Session("Userno"))
	Set rsBuildProdList = cnnBuildProdList.Execute(SQLBuildProdList)
	On Error Goto 0
	
	SQLBuildProdList = "CREATE TABLE zPRC_AccountQuotedItems_ProdList_" & trim(Session("Userno"))
	SQLBuildProdList = SQLBuildProdList & "("
	SQLBuildProdList = SQLBuildProdList &	"              [InternalRecordIdentifier] [int] IDENTITY(1,1) NOT NULL, "
	SQLBuildProdList = SQLBuildProdList & "                [prodSKU] [varchar](255) NULL, "
	SQLBuildProdList = SQLBuildProdList & "                [Description] [varchar](255) NULL, "
	SQLBuildProdList = SQLBuildProdList & "                [Category] [int] NULL, "
	SQLBuildProdList = SQLBuildProdList & "                [UM] [varchar](255) NULL, "
	SQLBuildProdList = SQLBuildProdList & "                [Price] [float] NULL, "
	SQLBuildProdList = SQLBuildProdList & "                [Cost] [float] NULL, "
	SQLBuildProdList = SQLBuildProdList & "                [CaseConversionFactor] [int], "
	SQLBuildProdList = SQLBuildProdList & "                [CaseDescription] [varchar](255) NULL, "
	SQLBuildProdList = SQLBuildProdList & "                [DataSource] [varchar](255) NULL "
	SQLBuildProdList = SQLBuildProdList & ")"
	'Response.Write(SQLBuildProdList & "<BR>")
	Set rsBuildProdList = cnnBuildProdList.Execute(SQLBuildProdList)
	
	
	'Create indexes
	SQLBuildProdList = "CREATE NONCLUSTERED INDEX [IX_zPRC_AccountQuotedItems_ProdList_1_1] ON [zPRC_AccountQuotedItems_ProdList_" & trim(Session("Userno")) & "] "
	SQLBuildProdList = SQLBuildProdList & "( "
	SQLBuildProdList = SQLBuildProdList & "[prodSKU] ASC "
	SQLBuildProdList = SQLBuildProdList & ")WITH (PAD_INDEX = OFF, STATISTICS_NORECOMPUTE = OFF, SORT_IN_TEMPDB = OFF, DROP_EXISTING = OFF, ONLINE = OFF, ALLOW_ROW_LOCKS = ON, ALLOW_PAGE_LOCKS = ON) ON [PRIMARY]"
	'Response.Write(SQLBuildProdList & "<BR>")
	Set rsBuildProdList = cnnBuildProdList.Execute(SQLBuildProdList)
	
	
	SQLBuildProdList = "CREATE NONCLUSTERED INDEX [IX_zPRC_AccountQuotedItems_ProdList_1_2] ON [zPRC_AccountQuotedItems_ProdList_" & trim(Session("Userno")) & "] "
	SQLBuildProdList = SQLBuildProdList & "( "
	SQLBuildProdList = SQLBuildProdList & "[prodSKU] ASC, [UM] ASC "
	SQLBuildProdList = SQLBuildProdList & ")WITH (PAD_INDEX = OFF, STATISTICS_NORECOMPUTE = OFF, SORT_IN_TEMPDB = OFF, DROP_EXISTING = OFF, ONLINE = OFF, ALLOW_ROW_LOCKS = ON, ALLOW_PAGE_LOCKS = ON) ON [PRIMARY]"
	'Response.Write(SQLBuildProdList & "<BR>")
	Set rsBuildProdList = cnnBuildProdList.Execute(SQLBuildProdList)
	
	
	
	'Insert all the products which are not currently quoted to this customer
	SQLBuildProdList = "INSERT INTO zPRC_AccountQuotedItems_ProdList_" & trim(Session("Userno")) & " (prodSKU, Description, Category, UM, Price, Cost, CaseConversionFactor, CaseDescription, DataSource) "
	SQLBuildProdList = SQLBuildProdList & "SELECT PartNo, Description, Category, CasePricing, ListPriceLevel1, UnitCost, CaseConversionFactor, CaseDescription, 'TABLE DATA' FROM Product "
	SQLBuildProdList = SQLBuildProdList & "WHERE (Category > 1) AND (Category < 19) AND ({ fn CONCAT(PartNo, CasePricing) } NOT IN "
	SQLBuildProdList = SQLBuildProdList & "(SELECT { fn CONCAT(prodSKU, QuoteType) } AS Expr1 FROM zPRC_AccountQuotedItems_" & trim(Session("Userno")) & "))"
	'Response.Write(SQLBuildProdList & "<BR>")
	Set rsBuildProdList = cnnBuildProdList.Execute(SQLBuildProdList)
	
	
	'Now fill in any missing CASES using the UNIT information
	SQLBuildProdList = "SELECT * FROM zPRC_AccountQuotedItems_ProdList_" & trim(Session("Userno")) & " WHERE UM = 'U' "
	SQLBuildProdList = SQLBuildProdList & "AND prodSKU NOT IN (SELECT prodSKU FROM zPRC_AccountQuotedItems_" & trim(Session("Userno")) & " WHERE QuoteType = 'C')"
	
	'Response.Write(SQLBuildProdList & "<BR>")
	Set rsBuildProdList = cnnBuildProdList.Execute(SQLBuildProdList)
	
	
	If Not rsBuildProdList.EOF Then
		Do While Not rsBuildProdList.EOF
			
			SQLBuildProdList2 = "INSERT INTO zPRC_AccountQuotedItems_ProdList_" & trim(Session("Userno")) & " (prodSKU, Description, Category, UM, Price, Cost, CaseConversionFactor, CaseDescription, DataSource) "
			SQLBuildProdList2 = SQLBuildProdList2 & " VALUES ("
			SQLBuildProdList2 = SQLBuildProdList2 & "'" & rsBuildProdList("prodSKU") & "', "
			SQLBuildProdList2 = SQLBuildProdList2 & "'" & Replace(rsBuildProdList("CaseDescription"),"'","''") & "', "
			SQLBuildProdList2 = SQLBuildProdList2 & rsBuildProdList("Category") & ", "
			SQLBuildProdList2 = SQLBuildProdList2 & "'C', "
			SQLBuildProdList2 = SQLBuildProdList2 & Round(rsBuildProdList("Price") * rsBuildProdList("CaseConversionFactor"),2) & ", "
			SQLBuildProdList2 = SQLBuildProdList2 & Round(rsBuildProdList("Cost") * rsBuildProdList("CaseConversionFactor"),2) & ", "
			SQLBuildProdList2 = SQLBuildProdList2 & rsBuildProdList("CaseConversionFactor") & ", "
			SQLBuildProdList2 = SQLBuildProdList2 & "'" & Replace(rsBuildProdList("CaseDescription"),"'","''") & "',"
			SQLBuildProdList2 = SQLBuildProdList2 & "'GENERATED CASE'"
			SQLBuildProdList2 = SQLBuildProdList2 & ")"
			
			'Response.Write(SQLBuildProdList2 & "<BR>")
			Set rsBuildProdList2 = cnnBuildProdList.Execute(SQLBuildProdList2)
			
			rsBuildProdList.movenext
		Loop
	End If 
	
	'Now do the opposite, fill in missing UNITS using the CASE information we have
	' BUT only do this when the UNIT information is not already present
	
	
	SQLBuildProdList = "SELECT * FROM zPRC_AccountQuotedItems_ProdList_" & trim(Session("Userno")) & " WHERE UM = 'C' AND prodSKU NOT IN "
	SQLBuildProdList = SQLBuildProdList & "(SELECT prodSKU FROM zPRC_AccountQuotedItems_ProdList_" & trim(Session("Userno")) & " WHERE UM = 'U') "
	SQLBuildProdList = SQLBuildProdList & "AND prodSKU NOT IN (SELECT prodSKU FROM zPRC_AccountQuotedItems_" & trim(Session("Userno")) & " WHERE QuoteType = 'U')"
	
	'Response.Write(SQLBuildProdList & "<BR>")
	Set rsBuildProdList = cnnBuildProdList.Execute(SQLBuildProdList)
	
	
	If Not rsBuildProdList.EOF Then
		Do While Not rsBuildProdList.EOF
	
			SQLBuildProdList2 = "INSERT INTO zPRC_AccountQuotedItems_ProdList_" & trim(Session("Userno")) & " (prodSKU, Description, Category, UM, Price, Cost, CaseConversionFactor, CaseDescription, DataSource) "
			SQLBuildProdList2 = SQLBuildProdList2 & " VALUES ("
			SQLBuildProdList2 = SQLBuildProdList2 & "'" & rsBuildProdList("prodSKU") & "', "
			SQLBuildProdList2 = SQLBuildProdList2 & "'" & Replace(rsBuildProdList("Description"),"'","''") & "', "
			SQLBuildProdList2 = SQLBuildProdList2 & rsBuildProdList("Category") & ", "
			SQLBuildProdList2 = SQLBuildProdList2 & "'U', "
			SQLBuildProdList2 = SQLBuildProdList2 & Round(rsBuildProdList("Price") / rsBuildProdList("CaseConversionFactor"),2) & ", "
			SQLBuildProdList2 = SQLBuildProdList2 & Round(rsBuildProdList("Cost") / rsBuildProdList("CaseConversionFactor"),2) & ", "
			SQLBuildProdList2 = SQLBuildProdList2 & rsBuildProdList("CaseConversionFactor") & ", "
			SQLBuildProdList2 = SQLBuildProdList2 & "'" & Replace(rsBuildProdList("CaseDescription"),"'","''") & "', "
			SQLBuildProdList2 = SQLBuildProdList2 & "'GENERATED UNIT'"
			SQLBuildProdList2 = SQLBuildProdList2 & ")"
			
			'Response.Write(SQLBuildProdList2 & "<BR>")
			Set rsBuildProdList2 = cnnBuildProdList.Execute(SQLBuildProdList2)
				
			rsBuildProdList.movenext
		Loop
	End If 
	
	Set rsBuildProdList2 = Nothing
	Set rsBuildProdList = Nothing
	cnnBuildProdList.Close
	Set cnnBuildProdList = Nothing	   
Else
    'Response.Write("The table is in the database.")
End If

%>
