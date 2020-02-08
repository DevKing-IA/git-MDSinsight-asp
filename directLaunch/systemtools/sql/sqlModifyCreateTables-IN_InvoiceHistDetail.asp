<%	
	Set cnnCheckINInvoiceHistDetail = Server.CreateObject("ADODB.Connection")
	cnnCheckINInvoiceHistDetail.open (Session("ClientCnnString"))
	Set rsCheckINInvoiceHistDetail = Server.CreateObject("ADODB.Recordset")

	Err.Clear
	on error resume next
	Set rsCheckINInvoiceHistDetail = cnnCheckINInvoiceHistDetail.Execute("SELECT TOP 1 * FROM IN_InvoiceHistDetail")
	If Err.Description <> "" Then
		If InStr(Ucase(Err.Description),"INVALID OBJECT NAME") <> 0 Then
			On error goto 0		
			'The table is not there, we need to create it
			
		    SQLCheckINInvoiceHistDetail = "CREATE TABLE [IN_InvoiceHistDetail]( "
			SQLCheckINInvoiceHistDetail = SQLCheckINInvoiceHistDetail & "[InternalRecordIdentifier] [int] IDENTITY(1,1) NOT NULL, "
			SQLCheckINInvoiceHistDetail = SQLCheckINInvoiceHistDetail & "[RecordCreationDateTime] [datetime] NULL, "
			SQLCheckINInvoiceHistDetail = SQLCheckINInvoiceHistDetail & "[InvoiceID] [varchar](255) NULL, "
			SQLCheckINInvoiceHistDetail = SQLCheckINInvoiceHistDetail & "[CustID] [varchar](255) NULL, "
			SQLCheckINInvoiceHistDetail = SQLCheckINInvoiceHistDetail & "[InvoiceCreationDate] [datetime] NULL, "
			SQLCheckINInvoiceHistDetail = SQLCheckINInvoiceHistDetail & "[LineNumber] [int] NULL, "
			SQLCheckINInvoiceHistDetail = SQLCheckINInvoiceHistDetail & "[prodSKU] [varchar](255) NULL, "
			SQLCheckINInvoiceHistDetail = SQLCheckINInvoiceHistDetail & "[prodDescription] [varchar](8000) NULL, "
			SQLCheckINInvoiceHistDetail = SQLCheckINInvoiceHistDetail & "[prodCategoryDescription] [varchar](255) NULL, "
			SQLCheckINInvoiceHistDetail = SQLCheckINInvoiceHistDetail & "[PricePerUnitSold] [money] NULL, "
			SQLCheckINInvoiceHistDetail = SQLCheckINInvoiceHistDetail & "[CostPerUnitSold] [money] NULL, "
			SQLCheckINInvoiceHistDetail = SQLCheckINInvoiceHistDetail & "[QtyOrdered] [int] NULL, "
			SQLCheckINInvoiceHistDetail = SQLCheckINInvoiceHistDetail & "[QtyShipped] [int] NULL, "
			SQLCheckINInvoiceHistDetail = SQLCheckINInvoiceHistDetail & "[Taxable] [int] NULL, "
			SQLCheckINInvoiceHistDetail = SQLCheckINInvoiceHistDetail & "[UMSold] [varchar](255) NULL, "
			SQLCheckINInvoiceHistDetail = SQLCheckINInvoiceHistDetail & "[LineItemNotes] [varchar](8000) NULL, "
			SQLCheckINInvoiceHistDetail = SQLCheckINInvoiceHistDetail & "[ThisLineNotAProduct] [int] NULL "
			SQLCheckINInvoiceHistDetail = SQLCheckINInvoiceHistDetail & " ) ON [PRIMARY]"      

			Set rsCheckINInvoiceHistDetail = cnnCheckINInvoiceHistDetail.Execute(SQLCheckINInvoiceHistDetail)

			SQLCheckINInvoiceHistDetail = "ALTER TABLE [IN_InvoiceHistDetail] ADD CONSTRAINT [DF_IN_InvoiceDetail_RecordCreationDateTime]  DEFAULT (getdate()) FOR [RecordCreationDateTime]"
			Set rsCheckINInvoiceHistDetail = cnnCheckINInvoiceHistDetail.Execute(SQLCheckINInvoiceHistDetail)			
	
		End If
	End If


	SQLCheckINInvoiceHistDetail  = "SELECT COL_LENGTH('IN_InvoiceHistDetail', 'GL_AR_Account') AS IsItThere"
	Set rsCheckINInvoiceHistDetail  = cnnCheckINInvoiceHistDetail.Execute(SQLCheckINInvoiceHistDetail  )
	If IsNull(rsCheckINInvoiceHistDetail ("IsItThere")) Then
		SQLCheckINInvoiceHistDetail = "ALTER TABLE IN_InvoiceHistDetail ADD GL_AR_Account [varchar](255) NULL "
		Set rsCheckINInvoiceHistDetail = cnnCheckINInvoiceHistDetail.Execute(SQLCheckINInvoiceHistDetail)
	End If

	SQLCheckINInvoiceHistDetail  = "SELECT COL_LENGTH('IN_InvoiceHistDetail', 'GL_Account') AS IsItThere"
	Set rsCheckINInvoiceHistDetail  = cnnCheckINInvoiceHistDetail.Execute(SQLCheckINInvoiceHistDetail  )
	If IsNull(rsCheckINInvoiceHistDetail ("IsItThere")) Then
		SQLCheckINInvoiceHistDetail = "ALTER TABLE IN_InvoiceHistDetail ADD GL_Account [varchar](255) NULL "
		Set rsCheckINInvoiceHistDetail = cnnCheckINInvoiceHistDetail.Execute(SQLCheckINInvoiceHistDetail)
	End If

	SQLCheckINInvoiceHistDetail  = "SELECT COL_LENGTH('IN_InvoiceHistDetail', 'TotalTaxForLine') AS IsItThere"
	Set rsCheckINInvoiceHistDetail  = cnnCheckINInvoiceHistDetail.Execute(SQLCheckINInvoiceHistDetail  )
	If IsNull(rsCheckINInvoiceHistDetail ("IsItThere")) Then
		SQLCheckINInvoiceHistDetail = "ALTER TABLE IN_InvoiceHistDetail ADD TotalTaxForLine [float] NULL "
		Set rsCheckINInvoiceHistDetail = cnnCheckINInvoiceHistDetail.Execute(SQLCheckINInvoiceHistDetail)
	End If
			
	set rsCheckINInvoiceHistDetail = nothing
	cnnCheckINInvoiceHistDetail.close
	set cnnCheckINInvoiceHistDetail = nothing
				
%>