<%
	Set cnnPRProspectEmailMessagesAttachments = Server.CreateObject("ADODB.Connection")
	cnnPRProspectEmailMessagesAttachments.open (Session("ClientCnnString"))
	Set rsPRProspectEmailMessagesAttachments = Server.CreateObject("ADODB.Recordset")
	rsPRProspectEmailMessagesAttachments.CursorLocation = 3 

	Err.Clear
	on error resume next
	Set rsPRProspectEmailMessagesAttachments = cnnPRProspectEmailMessagesAttachments.Execute("SELECT * FROM PR_ProspectEmailMessagesAttachments")
	If Err.Description <> "" Then
		If InStr(Ucase(Err.Description),"INVALID OBJECT NAME") <> 0 Then
			On error goto 0
			'The table is not there, we need to create it
			SQLBuild = "CREATE TABLE PR_ProspectEmailMessagesAttachments ("
			SQLBuild = SQLBuild & "[InternalRecordIdentifier] [int] IDENTITY(1,1) NOT NULL, "
			SQLBuild = SQLBuild & "	[RecordCreationDateTime] [datetime] NULL CONSTRAINT [DF_PR_ProspectEmailMessagesAttachments_RecordCreationDateTime]  DEFAULT (getdate()), "
			SQLBuild = SQLBuild & "	[MessageID] [varchar](8000) NULL, "
			SQLBuild = SQLBuild & "	[AttachmentFile] [varchar](8000) NULL, "
			SQLBuild = SQLBuild & "	[originalfilename] [varchar](8000) NULL "			
			SQLBuild = SQLBuild & ") ON [PRIMARY]"

			Set rsPRProspectEmailMessagesAttachments = cnnPRProspectEmailMessagesAttachments.Execute(SQLBuild)
			
		End If
	End If
	On Error Goto 0

	set rsPRProspectEmailMessagesAttachments = nothing
	cnnPRProspectEmailMessagesAttachments.close
	set cnnPRProspectEmailMessagesAttachments = nothing
%>