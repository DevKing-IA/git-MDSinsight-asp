<%
	Set cnnPRProspectEmailMessagesBody = Server.CreateObject("ADODB.Connection")
	cnnPRProspectEmailMessagesBody.open (Session("ClientCnnString"))
	Set rsPRProspectEmailMessagesBody = Server.CreateObject("ADODB.Recordset")
	rsPRProspectEmailMessagesBody.CursorLocation = 3 

	Err.Clear
	on error resume next
	Set rsPRProspectEmailMessagesBody = cnnPRProspectEmailMessagesBody.Execute("SELECT * FROM PR_ProspectEmailMessagesBody")
	If Err.Description <> "" Then
		If InStr(Ucase(Err.Description),"INVALID OBJECT NAME") <> 0 Then
			On error goto 0
			'The table is not there, we need to create it
			SQLBuild = "CREATE TABLE PR_ProspectEmailMessagesBody ("
			SQLBuild = SQLBuild & "[InternalRecordIdentifier] [int] IDENTITY(1,1) NOT NULL, "
			SQLBuild = SQLBuild & "	[RecordCreationDateTime] [datetime] NULL CONSTRAINT [DF_PR_ProspectEmailMessagesBody_RecordCreationDateTime]  DEFAULT (getdate()), "
			SQLBuild = SQLBuild & "	[MessageID] [varchar](8000) NULL, "
			SQLBuild = SQLBuild & "	[MessageBody] [varchar](8000) NULL "
			SQLBuild = SQLBuild & ") ON [PRIMARY]"

			Set rsPRProspectEmailMessagesBody = cnnPRProspectEmailMessagesBody.Execute(SQLBuild)
			
		End If
	End If
	On Error Goto 0

	set rsPRProspectEmailMessagesBody = nothing
	cnnPRProspectEmailMessagesBody.close
	set cnnPRProspectEmailMessagesBody = nothing
%>