<%	

	Set cnnAR_CustomerQuotedItems = Server.CreateObject("ADODB.Connection")
	cnnAR_CustomerQuotedItems.open (Session("ClientCnnString"))
	Set rsAR_CustomerQuotedItems = Server.CreateObject("ADODB.Recordset")
	rsAR_CustomerQuotedItems.CursorLocation = 3 


	Err.Clear
	on error resume next
	Set rsAR_CustomerQuotedItems = cnnAR_CustomerQuotedItems.Execute("SELECT TOP 1 * FROM AR_CustomerQuotedItems")
	If Err.Description <> "" Then
		If InStr(Ucase(Err.Description),"INVALID OBJECT NAME") <> 0 Then
			On error goto 0		
			'The table is not there, we need to create it
			SQLAR_CustomerQuotedItems = "CREATE TABLE [AR_CustomerQuotedItems]( "
			SQLAR_CustomerQuotedItems = SQLAR_CustomerQuotedItems & " [InternalRecordIdentifier] [int] IDENTITY(2,1) NOT NULL, "
			SQLAR_CustomerQuotedItems = SQLAR_CustomerQuotedItems & " [RecordCreationDateTime] [datetime] NULL CONSTRAINT [DF_AR_CustomerQuotedItems]  DEFAULT (getdate()), "
			SQLAR_CustomerQuotedItems = SQLAR_CustomerQuotedItems & " [CustID] [varchar](255) NULL, "
			SQLAR_CustomerQuotedItems = SQLAR_CustomerQuotedItems & " [prodSKU] [varchar](255) NULL, "
			SQLAR_CustomerQuotedItems = SQLAR_CustomerQuotedItems & " [SeqNo] [int] NULL, "
			SQLAR_CustomerQuotedItems = SQLAR_CustomerQuotedItems & " [Desc] [varchar](255) NULL, "
			SQLAR_CustomerQuotedItems = SQLAR_CustomerQuotedItems & " [UMQty] [int] NULL, "
			SQLAR_CustomerQuotedItems = SQLAR_CustomerQuotedItems & " [UM] [varchar](255) NULL, "
			SQLAR_CustomerQuotedItems = SQLAR_CustomerQuotedItems & " [UserDefined] [bit] NULL, "
			SQLAR_CustomerQuotedItems = SQLAR_CustomerQuotedItems & " [DepositAmount] [money] NULL, "
			SQLAR_CustomerQuotedItems = SQLAR_CustomerQuotedItems & " [StreamW_ShipToLocation] [varchar](255) NULL, "
			SQLAR_CustomerQuotedItems = SQLAR_CustomerQuotedItems & " [StreamW_POS] [varchar](255) NULL, "
			SQLAR_CustomerQuotedItems = SQLAR_CustomerQuotedItems & " [StreamW_ShipPOS] [varchar](255) NULL, "
			SQLAR_CustomerQuotedItems = SQLAR_CustomerQuotedItems & " [MPLEX_Price] [money] NULL, "
			SQLAR_CustomerQuotedItems = SQLAR_CustomerQuotedItems & " [ParLevel] [float] NULL "
			SQLAR_CustomerQuotedItems = SQLAR_CustomerQuotedItems & ") ON [PRIMARY]"
			Set rsAR_CustomerQuotedItems = cnnAR_CustomerQuotedItems.Execute(SQLAR_CustomerQuotedItems)
		End If
	End If
	
	
	set rsAR_CustomerQuotedItems = nothing
	cnnAR_CustomerQuotedItems.close
	set cnnAR_CustomerQuotedItems = nothing


%>