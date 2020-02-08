<%	
	Set cnnAPI_RA_PostResults = Server.CreateObject("ADODB.Connection")
	cnnAPI_RA_PostResults.open (Session("ClientCnnString"))
	Set rsAPI_RA_PostResults = Server.CreateObject("ADODB.Recordset")
	rsAPI_RA_PostResults.CursorLocation = 3 


	Err.Clear
	on error resume next
	Set rsAPI_RA_PostResults = cnnAPI_RA_PostResults.Execute("SELECT TOP 1 * FROM API_RA_PostResults ORDER BY InternalRecordIdentifier DESC")
	If Err.Description <> "" Then
		If InStr(Ucase(Err.Description),"INVALID OBJECT NAME") <> 0 Then
			On error goto 0		
			'The table is not there, we need to create it
			SQLBuildResults = "CREATE TABLE [API_RA_PostResults]("
			SQLBuildResults = SQLBuildResults & "	[InternalRecordIdentifier] [int] IDENTITY(1,1) NOT NULL, "
			SQLBuildResults = SQLBuildResults & "	[RecordCreationDateTime] [datetime] NULL CONSTRAINT [DF_API_RA_PostResults_RecordCreationDateTime]  DEFAULT (getdate()),"
			SQLBuildResults = SQLBuildResults & "	[APIKey] [varchar](50) NULL,"
			SQLBuildResults = SQLBuildResults & "	[RAID] [varchar](50) NULL,"
			SQLBuildResults = SQLBuildResults & "	[LogEntryThread] [int] NULL,"
			SQLBuildResults = SQLBuildResults & "	[PostResults] [varchar](50) NULL,"
			SQLBuildResults = SQLBuildResults & "	[PostStatusMessage] [varchar](8000) NULL"
			SQLBuildResults = SQLBuildResults & ") ON [PRIMARY]"
			
			Set rsAPI_RA_PostResults = cnnAPI_RA_PostResults.Execute(SQLBuildResults)
		End If
	End If
				
	set rsAPI_RA_PostResults = nothing
	cnnAPI_RA_PostResults.close
	set cnnAPI_RA_PostResults = nothing
				
%>