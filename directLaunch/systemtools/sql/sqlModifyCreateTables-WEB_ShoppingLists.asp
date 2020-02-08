<%	

	Set cnnWEB_ShoppingLists = Server.CreateObject("ADODB.Connection")
	cnnWEB_ShoppingLists.open (Session("ClientCnnString"))
	Set rsWEB_ShoppingLists = Server.CreateObject("ADODB.Recordset")
	rsWEB_ShoppingLists.CursorLocation = 3 


	Err.Clear
	on error resume next
	Set rsWEB_ShoppingLists = cnnWEB_ShoppingLists.Execute("SELECT TOP 1 * FROM WEB_ShoppingLists")
	If Err.Description <> "" Then
		If InStr(Ucase(Err.Description),"INVALID OBJECT NAME") <> 0 Then
			On error goto 0		
			'The table is not there, we need to create it
			SQLWEB_ShoppingLists = "CREATE TABLE [WEB_ShoppingLists]( "
			SQLWEB_ShoppingLists = SQLWEB_ShoppingLists & " [InternalRecordIdentifier] [int] IDENTITY(2,1) NOT NULL, "
			SQLWEB_ShoppingLists = SQLWEB_ShoppingLists & " [RecordCreationDateTime] [datetime] NULL CONSTRAINT [DF_WEB_ShoppingLists]  DEFAULT (getdate()), "
			SQLWEB_ShoppingLists = SQLWEB_ShoppingLists & " [ProdSKU] [varchar](255) NULL, "
			SQLWEB_ShoppingLists = SQLWEB_ShoppingLists & " [CustID] [varchar](255) NULL, "
			SQLWEB_ShoppingLists = SQLWEB_ShoppingLists & " [UserNo] [int] NULL, "
			SQLWEB_ShoppingLists = SQLWEB_ShoppingLists & " [ListName] [varchar](255) NULL, "
			SQLWEB_ShoppingLists = SQLWEB_ShoppingLists & " [Qty] [int] NULL, "
			SQLWEB_ShoppingLists = SQLWEB_ShoppingLists & " [UM] [varchar](255) NULL "
			SQLWEB_ShoppingLists = SQLWEB_ShoppingLists & ") ON [PRIMARY]"
			Set rsWEB_ShoppingLists = cnnWEB_ShoppingLists.Execute(SQLWEB_ShoppingLists)
		End If
	End If
	
	
	set rsWEB_ShoppingLists = nothing
	cnnWEB_ShoppingLists.close
	set cnnWEB_ShoppingLists = nothing


%>