<%	
	Set cnnCheckINInvoiceHistHeader = Server.CreateObject("ADODB.Connection")
	cnnCheckINInvoiceHistHeader.open (Session("ClientCnnString"))
	Set rsCheckINInvoiceHistHeader = Server.CreateObject("ADODB.Recordset")

	Err.Clear
	on error resume next
	Set rsCheckINInvoiceHistHeader = cnnCheckINInvoiceHistHeader.Execute("SELECT TOP 1 * FROM IN_InvoiceHistHeader")
	If Err.Description <> "" Then
		If InStr(Ucase(Err.Description),"INVALID OBJECT NAME") <> 0 Then
			On error goto 0		
			'The table is not there, we need to create it
			
		    SQLCheckINInvoiceHistHeader = "CREATE TABLE [IN_InvoiceHistHeader]( "
		    
		    
			SQLCheckINInvoiceHistHeader = SQLCheckINInvoiceHistHeader & "[InternalRecordIdentifier] [int] IDENTITY(1,1) NOT NULL, "
			SQLCheckINInvoiceHistHeader = SQLCheckINInvoiceHistHeader & "[RecordCreationDateTime] [datetime] NULL, "
			SQLCheckINInvoiceHistHeader = SQLCheckINInvoiceHistHeader & "[InvoiceID] [varchar](255) NULL, "
			SQLCheckINInvoiceHistHeader = SQLCheckINInvoiceHistHeader & "[CustID] [varchar](255) NULL, "
			SQLCheckINInvoiceHistHeader = SQLCheckINInvoiceHistHeader & "[AlternateCustID] [varchar](255) NULL, "
			SQLCheckINInvoiceHistHeader = SQLCheckINInvoiceHistHeader & "[InvoiceCreationDate] [datetime] NULL, "
			SQLCheckINInvoiceHistHeader = SQLCheckINInvoiceHistHeader & "[InvoiceDueDate] [datetime] NULL, "
			SQLCheckINInvoiceHistHeader = SQLCheckINInvoiceHistHeader & "[InvoiceType] [varchar](255) NULL, "
			SQLCheckINInvoiceHistHeader = SQLCheckINInvoiceHistHeader & "[InvoiceGrandTotal] [money] NULL, "
			SQLCheckINInvoiceHistHeader = SQLCheckINInvoiceHistHeader & "[OrderDate] [datetime] NULL, "
			SQLCheckINInvoiceHistHeader = SQLCheckINInvoiceHistHeader & "[Terms] [varchar] (255) NULL, "
			SQLCheckINInvoiceHistHeader = SQLCheckINInvoiceHistHeader & "[ShipVia] [varchar] (255) NULL, "
			SQLCheckINInvoiceHistHeader = SQLCheckINInvoiceHistHeader & "[CustName] [varchar](255) NULL, "
			SQLCheckINInvoiceHistHeader = SQLCheckINInvoiceHistHeader & "[BackendShipToIDIfApplicable] [varchar](255) NULL, "
			SQLCheckINInvoiceHistHeader = SQLCheckINInvoiceHistHeader & "[ShipToName] [varchar](255) NULL, "
			SQLCheckINInvoiceHistHeader = SQLCheckINInvoiceHistHeader & "[ShipToAddr1] [varchar](255) NULL, "
			SQLCheckINInvoiceHistHeader = SQLCheckINInvoiceHistHeader & "[ShipToAddr2] [varchar](255) NULL, "
			SQLCheckINInvoiceHistHeader = SQLCheckINInvoiceHistHeader & "[ShipToCity] [varchar](255) NULL, "
			SQLCheckINInvoiceHistHeader = SQLCheckINInvoiceHistHeader & "[ShipToState] [varchar](255) NULL, "
			SQLCheckINInvoiceHistHeader = SQLCheckINInvoiceHistHeader & "[ShipToPostalCode] [varchar](255) NULL, "
			SQLCheckINInvoiceHistHeader = SQLCheckINInvoiceHistHeader & "[ShipToContact] [varchar](255) NULL, "
			SQLCheckINInvoiceHistHeader = SQLCheckINInvoiceHistHeader & "[ShipToDescription] [varchar](255) NULL, "
			SQLCheckINInvoiceHistHeader = SQLCheckINInvoiceHistHeader & "[BackendBillToIDIfApplicable] [varchar](255) NULL, "
			SQLCheckINInvoiceHistHeader = SQLCheckINInvoiceHistHeader & "[BillToName] [varchar](255) NULL, "
			SQLCheckINInvoiceHistHeader = SQLCheckINInvoiceHistHeader & "[BillToAddr1] [varchar](255) NULL, "
			SQLCheckINInvoiceHistHeader = SQLCheckINInvoiceHistHeader & "[BillToAddr2] [varchar](255) NULL, "
			SQLCheckINInvoiceHistHeader = SQLCheckINInvoiceHistHeader & "[BillToCity] [varchar](255) NULL, "
			SQLCheckINInvoiceHistHeader = SQLCheckINInvoiceHistHeader & "[BillToState] [varchar](255) NULL, "
			SQLCheckINInvoiceHistHeader = SQLCheckINInvoiceHistHeader & "[BillToPostalCode] [varchar](255) NULL, "
			SQLCheckINInvoiceHistHeader = SQLCheckINInvoiceHistHeader & "[BillToContact] [varchar](255) NULL, "
			SQLCheckINInvoiceHistHeader = SQLCheckINInvoiceHistHeader & "[BillToDescription] [varchar](255) NULL, " 
			SQLCheckINInvoiceHistHeader = SQLCheckINInvoiceHistHeader & "[PONumber] [varchar](255) NULL " 
			SQLCheckINInvoiceHistHeader = SQLCheckINInvoiceHistHeader & " ) ON [PRIMARY]"      

			Set rsCheckINInvoiceHistHeader = cnnCheckINInvoiceHistHeader.Execute(SQLCheckINInvoiceHistHeader)

			SQLCheckINInvoiceHistHeader = "ALTER TABLE [IN_InvoiceHistHeader] ADD CONSTRAINT [DF_IN_InvoiceHeader_RecordCreationDateTime]  DEFAULT (getdate()) FOR [RecordCreationDateTime]"
			Set rsCheckINInvoiceHistHeader = cnnCheckINInvoiceHistHeader.Execute(SQLCheckINInvoiceHistHeader)			
	
		End If
	End If

	SQLCheckINInvoiceHistHeader = "SELECT COL_LENGTH('IN_InvoiceHistHeader', 'BackendShipToIDIfApplicable') AS IsItThere"
	Set rsCheckINInvoiceHistHeader = cnnCheckINInvoiceHistHeader.Execute(SQLCheckINInvoiceHistHeader)
	If IsNull(rsCheckINInvoiceHistHeader("IsItThere")) Then
		SQLCheckINInvoiceHistHeader = "ALTER TABLE IN_InvoiceHistHeader ADD BackendShipToIDIfApplicable [varchar] (255) NULL"
		Set rsCheckINInvoiceHistHeader = cnnCheckINInvoiceHistHeader.Execute(SQLCheckINInvoiceHistHeader)
	End If

	SQLCheckINInvoiceHistHeader = "SELECT COL_LENGTH('IN_InvoiceHistHeader', 'BackendBillToIDIfApplicable') AS IsItThere"
	Set rsCheckINInvoiceHistHeader = cnnCheckINInvoiceHistHeader.Execute(SQLCheckINInvoiceHistHeader)
	If IsNull(rsCheckINInvoiceHistHeader("IsItThere")) Then
		SQLCheckINInvoiceHistHeader = "ALTER TABLE IN_InvoiceHistHeader ADD BackendBillToIDIfApplicable [varchar] (255) NULL"
		Set rsCheckINInvoiceHistHeader = cnnCheckINInvoiceHistHeader.Execute(SQLCheckINInvoiceHistHeader)
	End If

	SQLCheckINInvoiceHistHeader = "SELECT COL_LENGTH('IN_InvoiceHistHeader', 'PONumber') AS IsItThere"
	Set rsCheckINInvoiceHistHeader = cnnCheckINInvoiceHistHeader.Execute(SQLCheckINInvoiceHistHeader)
	If IsNull(rsCheckINInvoiceHistHeader("IsItThere")) Then
		SQLCheckINInvoiceHistHeader = "ALTER TABLE IN_InvoiceHistHeader ADD PONumber [varchar] (255) NULL"
		Set rsCheckINInvoiceHistHeader = cnnCheckINInvoiceHistHeader.Execute(SQLCheckINInvoiceHistHeader)
	End If

	SQLCheckINInvoiceHistHeader = "SELECT COL_LENGTH('IN_InvoiceHistHeader', 'Terms') AS IsItThere"
	Set rsCheckINInvoiceHistHeader = cnnCheckINInvoiceHistHeader.Execute(SQLCheckINInvoiceHistHeader)
	If IsNull(rsCheckINInvoiceHistHeader("IsItThere")) Then
		SQLCheckINInvoiceHistHeader = "ALTER TABLE IN_InvoiceHistHeader ADD Terms [varchar] (255) NULL"
		Set rsCheckINInvoiceHistHeader = cnnCheckINInvoiceHistHeader.Execute(SQLCheckINInvoiceHistHeader)
	End If

	SQLCheckINInvoiceHistHeader = "SELECT COL_LENGTH('IN_InvoiceHistHeader', 'OrderDate') AS IsItThere"
	Set rsCheckINInvoiceHistHeader = cnnCheckINInvoiceHistHeader.Execute(SQLCheckINInvoiceHistHeader)
	If IsNull(rsCheckINInvoiceHistHeader("IsItThere")) Then
		SQLCheckINInvoiceHistHeader = "ALTER TABLE IN_InvoiceHistHeader ADD OrderDate [datetime] NULL"
		Set rsCheckINInvoiceHistHeader = cnnCheckINInvoiceHistHeader.Execute(SQLCheckINInvoiceHistHeader)
	End If

	SQLCheckINInvoiceHistHeader = "SELECT COL_LENGTH('IN_InvoiceHistHeader', 'ShipToLongitude') AS IsItThere"
	Set rsCheckINInvoiceHistHeader = cnnCheckINInvoiceHistHeader.Execute(SQLCheckINInvoiceHistHeader)
	If IsNull(rsCheckINInvoiceHistHeader("IsItThere")) Then
		SQLCheckINInvoiceHistHeader = "ALTER TABLE IN_InvoiceHistHeader ADD ShipToLongitude [varchar] (255) NULL"
		Set rsCheckINInvoiceHistHeader = cnnCheckINInvoiceHistHeader.Execute(SQLCheckINInvoiceHistHeader)
	End If

	SQLCheckINInvoiceHistHeader = "SELECT COL_LENGTH('IN_InvoiceHistHeader', 'ShipToLatitude') AS IsItThere"
	Set rsCheckINInvoiceHistHeader = cnnCheckINInvoiceHistHeader.Execute(SQLCheckINInvoiceHistHeader)
	If IsNull(rsCheckINInvoiceHistHeader("IsItThere")) Then
		SQLCheckINInvoiceHistHeader = "ALTER TABLE IN_InvoiceHistHeader ADD ShipToLatitude [varchar] (255) NULL"
		Set rsCheckINInvoiceHistHeader = cnnCheckINInvoiceHistHeader.Execute(SQLCheckINInvoiceHistHeader)
	End If
	
	SQLCheckINInvoiceHistHeader = "SELECT COL_LENGTH('IN_InvoiceHistHeader', 'BillToLongitude') AS IsItThere"
	Set rsCheckINInvoiceHistHeader = cnnCheckINInvoiceHistHeader.Execute(SQLCheckINInvoiceHistHeader)
	If IsNull(rsCheckINInvoiceHistHeader("IsItThere")) Then
		SQLCheckINInvoiceHistHeader = "ALTER TABLE IN_InvoiceHistHeader ADD BillToLongitude [varchar] (255) NULL"
		Set rsCheckINInvoiceHistHeader = cnnCheckINInvoiceHistHeader.Execute(SQLCheckINInvoiceHistHeader)
	End If

	SQLCheckINInvoiceHistHeader = "SELECT COL_LENGTH('IN_InvoiceHistHeader', 'BillToLatitude') AS IsItThere"
	Set rsCheckINInvoiceHistHeader = cnnCheckINInvoiceHistHeader.Execute(SQLCheckINInvoiceHistHeader)
	If IsNull(rsCheckINInvoiceHistHeader("IsItThere")) Then
		SQLCheckINInvoiceHistHeader = "ALTER TABLE IN_InvoiceHistHeader ADD BillToLatitude [varchar] (255) NULL"
		Set rsCheckINInvoiceHistHeader = cnnCheckINInvoiceHistHeader.Execute(SQLCheckINInvoiceHistHeader)
	End If


	'This one is a drop
	SQLCheckINInvoiceHistDetail  = "SELECT COL_LENGTH('IN_InvoiceHistDetail', 'Longitude') AS IsItThere"
	Set rsCheckINInvoiceHistDetail  = cnnCheckINInvoiceHistDetail.Execute(SQLCheckINInvoiceHistDetail  )
	If NOT IsNull(rsCheckINInvoiceHistDetail ("IsItThere")) Then
		SQLCheckINInvoiceHistDetail = "ALTER TABLE IN_InvoiceHistDetail DROP COLUMN Longitude"
		Set rsCheckINInvoiceHistDetail = cnnCheckINInvoiceHistDetail.Execute(SQLCheckINInvoiceHistDetail)
	End If
	
	'This one is a drop
	SQLCheckINInvoiceHistDetail  = "SELECT COL_LENGTH('IN_InvoiceHistDetail', 'Latitude') AS IsItThere"
	Set rsCheckINInvoiceHistDetail  = cnnCheckINInvoiceHistDetail.Execute(SQLCheckINInvoiceHistDetail  )
	If NOT IsNull(rsCheckINInvoiceHistDetail ("IsItThere")) Then
		SQLCheckINInvoiceHistDetail = "ALTER TABLE IN_InvoiceHistDetail DROP COLUMN Latitude"
		Set rsCheckINInvoiceHistDetail = cnnCheckINInvoiceHistDetail.Execute(SQLCheckINInvoiceHistDetail)
	End If

	'This one is a drop
	SQLCheckINInvoiceHistDetail  = "SELECT COL_LENGTH('IN_InvoiceHistDetail', 'SalesTaxID') AS IsItThere"
	Set rsCheckINInvoiceHistDetail  = cnnCheckINInvoiceHistDetail.Execute(SQLCheckINInvoiceHistDetail  )
	If NOT IsNull(rsCheckINInvoiceHistDetail ("IsItThere")) Then
		SQLCheckINInvoiceHistDetail = "ALTER TABLE IN_InvoiceHistDetail DROP COLUMN SalesTaxID"
		Set rsCheckINInvoiceHistDetail = cnnCheckINInvoiceHistDetail.Execute(SQLCheckINInvoiceHistDetail)
	End If

	'This one is a drop
	SQLCheckINInvoiceHistDetail  = "SELECT COL_LENGTH('IN_InvoiceHistDetail', 'SalesTaxDesc') AS IsItThere"
	Set rsCheckINInvoiceHistDetail  = cnnCheckINInvoiceHistDetail.Execute(SQLCheckINInvoiceHistDetail  )
	If NOT IsNull(rsCheckINInvoiceHistDetail ("IsItThere")) Then
		SQLCheckINInvoiceHistDetail = "ALTER TABLE IN_InvoiceHistDetail DROP COLUMN SalesTaxDesc"
		Set rsCheckINInvoiceHistDetail = cnnCheckINInvoiceHistDetail.Execute(SQLCheckINInvoiceHistDetail)
	End If

	set rsCheckINInvoiceHistHeader = nothing
	cnnCheckINInvoiceHistHeader.close
	set cnnCheckINInvoiceHistHeader = nothing
				
%>