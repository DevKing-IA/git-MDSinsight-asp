<%	
	Set cnnAPI_AuditLog = Server.CreateObject("ADODB.Connection")
	cnnAPI_AuditLog.open (Session("ClientCnnString"))
	Set rsAPI_AuditLog = Server.CreateObject("ADODB.Recordset")
	rsAPI_AuditLog.CursorLocation = 3 


	Err.Clear
	on error resume next
	Set rsAPI_AuditLog = cnnAPI_AuditLog.Execute("SELECT TOP 1 * FROM API_AuditLog ORDER BY EntryID DESC")
	If Err.Description <> "" Then
		If InStr(Ucase(Err.Description),"INVALID OBJECT NAME") <> 0 Then
			On error goto 0		
			'The table is not there, we need to create it
			SQLAPI_AuditLog = "CREATE TABLE [API_AuditLog]( "
			SQLAPI_AuditLog = SQLAPI_AuditLog & " [EntryID] [int] IDENTITY(2,1) NOT NULL, "
			SQLAPI_AuditLog = SQLAPI_AuditLog & " [EntryThread] [int] NULL, "
			SQLAPI_AuditLog = SQLAPI_AuditLog & " [Identity] [varchar](50) NULL, "
			SQLAPI_AuditLog = SQLAPI_AuditLog & " [RecordCreated] [datetime] NULL CONSTRAINT [DF_api_tblAuditLog_RecordCreated]  DEFAULT (getdate()), "
			SQLAPI_AuditLog = SQLAPI_AuditLog & " [LogEntry] [varchar](8000) NULL, "
			SQLAPI_AuditLog = SQLAPI_AuditLog & " [LogEntryPart2] [varchar](8000) NULL, "
			SQLAPI_AuditLog = SQLAPI_AuditLog & " [LogEntryPart3] [varchar](8000) NULL, "
			SQLAPI_AuditLog = SQLAPI_AuditLog & " [Mode] [varchar](50) NULL, "
			SQLAPI_AuditLog = SQLAPI_AuditLog & " [ClientID] [varchar](50) NULL, "
			SQLAPI_AuditLog = SQLAPI_AuditLog & " [SerNo] [varchar](50) NULL, "
			SQLAPI_AuditLog = SQLAPI_AuditLog & " [apiModule] [varchar](255) NULL, "
			SQLAPI_AuditLog = SQLAPI_AuditLog & " [IPAddress] [varchar](255) NULL "
			SQLAPI_AuditLog = SQLAPI_AuditLog & ") ON [PRIMARY]"
		
			Set rsAPI_AuditLog = cnnAPI_AuditLog.Execute(SQLAPI_AuditLog)
		End If
	End If
				
	set rsAPI_AuditLog = nothing
	cnnAPI_AuditLog.close
	set cnnAPI_AuditLog = nothing
				
%>