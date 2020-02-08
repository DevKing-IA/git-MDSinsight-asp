<%
	Set cnnPRProspectEmailMessages = Server.CreateObject("ADODB.Connection")
	cnnPRProspectEmailMessages.open (Session("ClientCnnString"))
	Set rsPRProspectEmailMessages = Server.CreateObject("ADODB.Recordset")
	rsPRProspectEmailMessages.CursorLocation = 3 

	Err.Clear
	on error resume next
	Set rsPRProspectEmailMessages = cnnPRProspectEmailMessages.Execute("SELECT * FROM PR_ProspectEmailMessages")
	If Err.Description <> "" Then
		If InStr(Ucase(Err.Description),"INVALID OBJECT NAME") <> 0 Then
			On error goto 0
			'The table is not there, we need to create it
			SQLBuild = "CREATE TABLE PR_ProspectEmailMessages ("
			SQLBuild = SQLBuild & "[InternalRecordIdentifier] [int] IDENTITY(1,1) NOT NULL, "
			SQLBuild = SQLBuild & "	[RecordCreationDateTime] [datetime] NULL CONSTRAINT [DF_PR_ProspectEmailMessages_RecordCreationDateTime]  DEFAULT (getdate()), "
			SQLBuild = SQLBuild & "	[ProspectIntRecID] [int] NULL, "
			SQLBuild = SQLBuild & "	[MessageID] [varchar](8000) NULL, "
			SQLBuild = SQLBuild & "	[Subject] [varchar](8000) NULL, "
			SQLBuild = SQLBuild & "	[SenderEmail] [varchar](8000) NULL, "
			SQLBuild = SQLBuild & "	[SenderName] [varchar](8000) NULL, "
			SQLBuild = SQLBuild & "	[DateTimeReceived] [varchar](8000) NULL, "
			SQLBuild = SQLBuild & "	[RecipientEmails] [varchar](8000) NULL, "
			SQLBuild = SQLBuild & "	[CCEmails] [varchar](8000) NULL "
			SQLBuild = SQLBuild & ") ON [PRIMARY]"

			Set rsPRProspectEmailMessages = cnnPRProspectEmailMessages.Execute(SQLBuild)
			
		End If
	End If
	On Error Goto 0

	set rsPRProspectEmailMessages = nothing
	cnnPRProspectEmailMessages.close
	set cnnPRProspectEmailMessages = nothing
%>