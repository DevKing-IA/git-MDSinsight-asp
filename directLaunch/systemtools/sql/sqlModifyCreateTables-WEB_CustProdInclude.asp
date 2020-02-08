<%	

	Set cnnWEB_CustProdInclude = Server.CreateObject("ADODB.Connection")
	cnnWEB_CustProdInclude.open (Session("ClientCnnString"))
	Set rsWEB_CustProdInclude = Server.CreateObject("ADODB.Recordset")
	rsWEB_CustProdInclude.CursorLocation = 3 


	Err.Clear
	on error resume next
	Set rsWEB_CustProdInclude = cnnWEB_CustProdInclude.Execute("SELECT TOP 1 * FROM WEB_CustProdInclude")
	If Err.Description <> "" Then
		If InStr(Ucase(Err.Description),"INVALID OBJECT NAME") <> 0 Then
			On error goto 0		
			'The table is not there, we need to create it
			SQLWEB_CustProdInclude = "CREATE TABLE [WEB_CustProdInclude]( "
			SQLWEB_CustProdInclude = SQLWEB_CustProdInclude & " [InternalRecordIdentifier] [int] IDENTITY(2,1) NOT NULL, "
			SQLWEB_CustProdInclude = SQLWEB_CustProdInclude & " [RecordCreationDateTime] [datetime] NULL CONSTRAINT [DF_WEB_CustProdInclude]  DEFAULT (getdate()), "
			SQLWEB_CustProdInclude = SQLWEB_CustProdInclude & " [CustID] [varchar](255) NULL, "
			SQLWEB_CustProdInclude = SQLWEB_CustProdInclude & " [ProdSKU] [varchar](255) NULL "
			SQLWEB_CustProdInclude = SQLWEB_CustProdInclude & ") ON [PRIMARY]"
			Set rsWEB_CustProdInclude = cnnWEB_CustProdInclude.Execute(SQLWEB_CustProdInclude)
		End If
	End If
	
	
	set rsWEB_CustProdInclude = nothing
	cnnWEB_CustProdInclude.close
	set cnnWEB_CustProdInclude = nothing


%>