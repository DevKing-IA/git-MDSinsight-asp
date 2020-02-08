<%	

	Set cnnAuditLog = Server.CreateObject("ADODB.Connection")
	cnnAuditLog.open (Session("ClientCnnString"))
	Set rsAuditLog = Server.CreateObject("ADODB.Recordset")
	rsAuditLog.CursorLocation = 3 
	
	Err.Clear
	on error resume next
	Set rsAuditLog = cnnAuditLog.Execute("Select TOP 1 * from SC_AuditLogDLaunch order by EntryThread desc")
	If Err.Description <> "" Then
		If InStr(Ucase(Err.Description),"INVALID OBJECT NAME") <> 0 Then
			On error goto 0		
			'The table is not there, we need to create it
			SQLBuild = "CREATE TABLE SC_AuditLogDLaunch ("
			SQLBuild = SQLBuild & "InternalRecordIdentifier [int] IDENTITY(1,1) NOT NULL, "
			SQLBuild = SQLBuild & "RecordCreationDateTime [datetime] NULL CONSTRAINT [DF_SC_AuditLogDLaunch_RecordCreationDateTime] DEFAULT (getdate()), "
			SQLBuild = SQLBuild & "EntryThread [int] NULL, "
			SQLBuild = SQLBuild & "DirectLaunchName [varchar](255) NULL, "
			SQLBuild = SQLBuild & "DirectLaunchFile [varchar](255) NULL, "
			SQLBuild = SQLBuild & "LogEntry [varchar](8000) NULL "
			SQLBuild = SQLBuild & ") ON [PRIMARY]"
			Set rsAuditLog = cnnAuditLog.Execute(SQLBuild)
			set rsAuditLog = nothing
			cnnAuditLog.close
			set cnnAuditLog = nothing
		End If
	End If
	On Error Goto 0
				
%>