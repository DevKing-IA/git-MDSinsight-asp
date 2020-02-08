<%	

	Set cnnFS_WEB_Users = Server.CreateObject("ADODB.Connection")
	cnnFS_WEB_Users.open (Session("ClientCnnString"))
	Set rsFS_WEB_Users = Server.CreateObject("ADODB.Recordset")
	rsFS_WEB_Users.CursorLocation = 3 


	Err.Clear
	on error resume next
	Set rsFS_WEB_Users = cnnFS_WEB_Users.Execute("SELECT TOP 1 * FROM FS_WEB_Users")
	If Err.Description <> "" Then
		If InStr(Ucase(Err.Description),"INVALID OBJECT NAME") <> 0 Then
			On error goto 0		
			'The table is not there, we need to create it	
			
			SQLFS_WEB_Users = "CREATE TABLE [FS_WEB_Users]( "
			SQLFS_WEB_Users = SQLFS_WEB_Users & " [InternalRecordIdentifier] [int] IDENTITY(1,1) NOT NULL,"
			SQLFS_WEB_Users = SQLFS_WEB_Users & " [RecordCreationDateTime] [datetime] NULL CONSTRAINT [DF_FS_WEB_User_RecordCreationDateTime]  DEFAULT (getdate()),"
            SQLFS_WEB_Users = SQLFS_WEB_Users & " [CustID] [varchar](255) NULL, "
            SQLFS_WEB_Users = SQLFS_WEB_Users & " [ShipToID] [varchar](255) NULL, "
            SQLFS_WEB_Users = SQLFS_WEB_Users & " [PosID] [varchar](255) NULL, "
            SQLFS_WEB_Users = SQLFS_WEB_Users & " [fsUserEmail] [varchar](255) NULL, "
            SQLFS_WEB_Users = SQLFS_WEB_Users & " [fsUserPassword] [varchar](255) NULL, "
            SQLFS_WEB_Users = SQLFS_WEB_Users & " [fsUserFirstName] [varchar](255) NULL, "
            SQLFS_WEB_Users = SQLFS_WEB_Users & " [fsUserLastName] [varchar](255) NULL "
			SQLFS_WEB_Users = SQLFS_WEB_Users & " ) ON [PRIMARY] "

			Set rsFS_WEB_Users = cnnFS_WEB_Users.Execute(SQLFS_WEB_Users)
		End If
	End If
	
	
	set rsFS_WEB_Users = nothing
	cnnFS_WEB_Users.close
	set cnnFS_WEB_Users = nothing


%>