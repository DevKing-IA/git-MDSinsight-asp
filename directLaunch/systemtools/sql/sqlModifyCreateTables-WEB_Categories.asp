<%	

	Set cnnWEB_Categories = Server.CreateObject("ADODB.Connection")
	cnnWEB_Categories.open (Session("ClientCnnString"))
	Set rsWEB_Categories = Server.CreateObject("ADODB.Recordset")
	rsWEB_Categories.CursorLocation = 3 


	Err.Clear
	on error resume next
	Set rsWEB_Categories = cnnWEB_Categories.Execute("SELECT TOP 1 * FROM WEB_Categories")
	If Err.Description <> "" Then
		If InStr(Ucase(Err.Description),"INVALID OBJECT NAME") <> 0 Then
			On error goto 0		
			'The table is not there, we need to create it
			SQLWEB_Categories = "CREATE TABLE [WEB_Categories]( "
			SQLWEB_Categories = SQLWEB_Categories & " [InternalRecordIdentifier] [int] IDENTITY(2,1) NOT NULL, "
			SQLWEB_Categories = SQLWEB_Categories & " [RecordCreationDateTime] [datetime] NULL CONSTRAINT [DF_WEB_Categories]  DEFAULT (getdate()), "
			SQLWEB_Categories = SQLWEB_Categories & " [CategoryID] int NULL, "
			SQLWEB_Categories = SQLWEB_Categories & " [CategoryName] [varchar](50) NULL, "
			SQLWEB_Categories = SQLWEB_Categories & " [ParentCategoryID] [varchar](50) NULL, "
			SQLWEB_Categories = SQLWEB_Categories & " [ChildCategoryID] [varchar](50) NULL, "
			SQLWEB_Categories = SQLWEB_Categories & " [AllParentIDs] [varchar](255) NULL, "
			SQLWEB_Categories = SQLWEB_Categories & " [suggestionPrecedence] int NULL, "
			SQLWEB_Categories = SQLWEB_Categories & " [DisplayOnWeb] bit NOT NULL DEFAULT 1, "
			SQLWEB_Categories = SQLWEB_Categories & " [DisplayRank] int NOT NULL DEFAULT 0 "
			SQLWEB_Categories = SQLWEB_Categories & ") ON [PRIMARY]"
			Set rsWEB_Categories = cnnWEB_Categories.Execute(SQLWEB_Categories)
		End If
	End If
	
	
	set rsWEB_Categories = nothing
	cnnWEB_Categories.close
	set cnnWEB_Categories = nothing


%>