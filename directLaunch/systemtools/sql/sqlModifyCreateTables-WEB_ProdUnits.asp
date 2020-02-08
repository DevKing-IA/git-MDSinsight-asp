<%	

	Set cnnWEB_ProdUnits = Server.CreateObject("ADODB.Connection")
	cnnWEB_ProdUnits.open (Session("ClientCnnString"))
	Set rsWEB_ProdUnits = Server.CreateObject("ADODB.Recordset")
	rsWEB_ProdUnits.CursorLocation = 3 


	Err.Clear
	on error resume next
	Set rsWEB_ProdUnits = cnnWEB_ProdUnits.Execute("SELECT TOP 1 * FROM WEB_ProdUnits")
	If Err.Description <> "" Then
		If InStr(Ucase(Err.Description),"INVALID OBJECT NAME") <> 0 Then
			On error goto 0		
			'The table is not there, we need to create it
			SQLWEB_ProdUnits = "CREATE TABLE [WEB_ProdUnits]( "
			SQLWEB_ProdUnits = SQLWEB_ProdUnits & " [InternalRecordIdentifier] [int] IDENTITY(2,1) NOT NULL, "
			SQLWEB_ProdUnits = SQLWEB_ProdUnits & " [RecordCreationDateTime] [datetime] NULL CONSTRAINT [DF_WEB_ProdUnits]  DEFAULT (getdate()), "
			SQLWEB_ProdUnits = SQLWEB_ProdUnits & " [SKU] [varchar](255) NULL, "
			SQLWEB_ProdUnits = SQLWEB_ProdUnits & " [UM] [varchar](50) NULL, "
			SQLWEB_ProdUnits = SQLWEB_ProdUnits & " [Qty] [int] NULL, "
			SQLWEB_ProdUnits = SQLWEB_ProdUnits & " [ListPrice] [money] NULL, "
			SQLWEB_ProdUnits = SQLWEB_ProdUnits & " [CostPrice] [money] NULL, "
			SQLWEB_ProdUnits = SQLWEB_ProdUnits & " [MplexPrice1] [money] NULL, "
			SQLWEB_ProdUnits = SQLWEB_ProdUnits & " [MplexPrice2] [money] NULL, "
			SQLWEB_ProdUnits = SQLWEB_ProdUnits & " [MplexPrice3] [money] NULL, "
			SQLWEB_ProdUnits = SQLWEB_ProdUnits & " [MplexPrice4] [money] NULL, "
			SQLWEB_ProdUnits = SQLWEB_ProdUnits & " [MplexPrice5] [money] NULL, "
			SQLWEB_ProdUnits = SQLWEB_ProdUnits & " [UM_MPLEX] [varchar](50) NULL, "
			SQLWEB_ProdUnits = SQLWEB_ProdUnits & " [UMQty_MPLEX] [int] NULL "
			SQLWEB_ProdUnits = SQLWEB_ProdUnits & ") ON [PRIMARY]"
			Set rsWEB_ProdUnits = cnnWEB_ProdUnits.Execute(SQLWEB_ProdUnits)
		End If
	End If
	
	
	set rsWEB_ProdUnits = nothing
	cnnWEB_ProdUnits.close
	set cnnWEB_ProdUnits = nothing


%>