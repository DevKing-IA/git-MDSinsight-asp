<%	

	Set cnnWEB_HomepageProducts = Server.CreateObject("ADODB.Connection")
	cnnWEB_HomepageProducts.open (Session("ClientCnnString"))
	Set rsWEB_HomepageProducts = Server.CreateObject("ADODB.Recordset")
	rsWEB_HomepageProducts.CursorLocation = 3 


	Err.Clear
	on error resume next
	Set rsWEB_HomepageProducts = cnnWEB_HomepageProducts.Execute("SELECT TOP 1 * FROM WEB_HomepageProducts")
	If Err.Description <> "" Then
		If InStr(Ucase(Err.Description),"INVALID OBJECT NAME") <> 0 Then
			On error goto 0		
			'The table is not there, we need to create it
			SQLWEB_HomepageProducts = "CREATE TABLE [WEB_HomepageProducts]( "
			SQLWEB_HomepageProducts = SQLWEB_HomepageProducts & " [InternalRecordIdentifier] [int] IDENTITY(2,1) NOT NULL, "
			SQLWEB_HomepageProducts = SQLWEB_HomepageProducts & " [RecordCreationDateTime] [datetime] NULL CONSTRAINT [DF_WEB_HomepageProducts]  DEFAULT (getdate()), "
			SQLWEB_HomepageProducts = SQLWEB_HomepageProducts & " [ItemNo] [int] NOT NULL, "
			SQLWEB_HomepageProducts = SQLWEB_HomepageProducts & " [ProdID] [varchar](50) NULL "
			SQLWEB_HomepageProducts = SQLWEB_HomepageProducts & ") ON [PRIMARY]"
			Set rsWEB_HomepageProducts = cnnWEB_HomepageProducts.Execute(SQLWEB_HomepageProducts)
		End If
	End If
	
	
	set rsWEB_HomepageProducts = nothing
	cnnWEB_HomepageProducts.close
	set cnnWEB_HomepageProducts = nothing


%>