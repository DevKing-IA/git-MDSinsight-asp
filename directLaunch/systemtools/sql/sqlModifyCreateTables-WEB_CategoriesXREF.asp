<%	

	Set cnnWEB_CategoriesXREF = Server.CreateObject("ADODB.Connection")
	cnnWEB_CategoriesXREF.open (Session("ClientCnnString"))
	Set rsWEB_CategoriesXREF = Server.CreateObject("ADODB.Recordset")
	rsWEB_CategoriesXREF.CursorLocation = 3 


	Err.Clear
	on error resume next
	Set rsWEB_CategoriesXREF = cnnWEB_CategoriesXREF.Execute("SELECT TOP 1 * FROM WEB_CategoriesXREF")
	If Err.Description <> "" Then
		If InStr(Ucase(Err.Description),"INVALID OBJECT NAME") <> 0 Then
			On error goto 0		
			'The table is not there, we need to create it
			SQLWEB_CategoriesXREF = "CREATE TABLE [WEB_CategoriesXREF]( "
			SQLWEB_CategoriesXREF = SQLWEB_CategoriesXREF & " [InternalRecordIdentifier] [int] IDENTITY(2,1) NOT NULL, "
			SQLWEB_CategoriesXREF = SQLWEB_CategoriesXREF & " [RecordCreationDateTime] [datetime] NULL CONSTRAINT [DF_WEB_CategoriesXREF]  DEFAULT (getdate()), "
			SQLWEB_CategoriesXREF = SQLWEB_CategoriesXREF & " [Category1] [varchar](50) NULL, "
			SQLWEB_CategoriesXREF = SQLWEB_CategoriesXREF & " [Category2] [varchar](50) NULL, "
			SQLWEB_CategoriesXREF = SQLWEB_CategoriesXREF & " [Category3] [varchar](50) NULL, "
			SQLWEB_CategoriesXREF = SQLWEB_CategoriesXREF & " [Category4] [varchar](50) NULL, "
			SQLWEB_CategoriesXREF = SQLWEB_CategoriesXREF & " [ProdSKU] [varchar](50) NULL, "
			SQLWEB_CategoriesXREF = SQLWEB_CategoriesXREF & " [AllParentIDs] [varchar](255) NULL, "
			SQLWEB_CategoriesXREF = SQLWEB_CategoriesXREF & " [DisplayRank] int NOT NULL DEFAULT 0 "
			SQLWEB_CategoriesXREF = SQLWEB_CategoriesXREF & ") ON [PRIMARY]"
			Set rsWEB_CategoriesXREF = cnnWEB_CategoriesXREF.Execute(SQLWEB_CategoriesXREF)
		End If
	End If
	
	
	set rsWEB_CategoriesXREF = nothing
	cnnWEB_CategoriesXREF.close
	set cnnWEB_CategoriesXREF = nothing


%>