<%	
	Response.Write("sqlModifyCreateTables-SC_ExternalEmailMessages.asp" & "<br>")
	On Error Goto 0

	Set cnnCheckSCExternalEmailMessages = Server.CreateObject("ADODB.Connection")
	cnnCheckSCExternalEmailMessages.open (Session("ClientCnnString"))
	Set rsCheckSCExternalEmailMessages = Server.CreateObject("ADODB.Recordset")

	Err.Clear
	on error resume next
	Set rsCheckSCExternalEmailMessages = cnnCheckSCExternalEmailMessages.Execute("SELECT TOP 1 * FROM SC_ExternalEmailMessages")
	If Err.Description <> "" Then
		If InStr(Ucase(Err.Description),"INVALID OBJECT NAME") <> 0 Then
			On error goto 0		
			'The table is not there, we need to create it
			
		    SQLCheckSCExternalEmailMessages = "CREATE TABLE [SC_ExternalEmailMessages]("
			SQLCheckSCExternalEmailMessages = SQLCheckSCExternalEmailMessages & " [InternalRecordIdentifier] [int] IDENTITY(1,1) NOT NULL,"
			SQLCheckSCExternalEmailMessages = SQLCheckSCExternalEmailMessages & " [RecordCreationDateTime] [datetime] NULL CONSTRAINT [DF_SC_ExternalEmailMessages_RecordCreationDateTime]  DEFAULT (getdate()),"
			SQLCheckSCExternalEmailMessages = SQLCheckSCExternalEmailMessages & " [AccountUsername] [varchar](8000) NULL, "
			SQLCheckSCExternalEmailMessages = SQLCheckSCExternalEmailMessages & " [MessageID] [varchar](8000) NULL, "
			SQLCheckSCExternalEmailMessages = SQLCheckSCExternalEmailMessages & " [Subject] [varchar](8000) NULL, "
			SQLCheckSCExternalEmailMessages = SQLCheckSCExternalEmailMessages & " [SenderEmail] [varchar](8000) NULL, "
			SQLCheckSCExternalEmailMessages = SQLCheckSCExternalEmailMessages & " [SenderName] [varchar](8000) NULL, "
			SQLCheckSCExternalEmailMessages = SQLCheckSCExternalEmailMessages & " [MessageBody] [varchar](8000) NULL, "
			SQLCheckSCExternalEmailMessages = SQLCheckSCExternalEmailMessages & " [DatetimeReceived] [varchar](8000) NULL, "
			SQLCheckSCExternalEmailMessages = SQLCheckSCExternalEmailMessages & " [RecipientEmails] [varchar](8000) NULL, "
			SQLCheckSCExternalEmailMessages = SQLCheckSCExternalEmailMessages & " [CCEmails] [varchar](8000) NULL "
			SQLCheckSCExternalEmailMessages = SQLCheckSCExternalEmailMessages & " ) ON [PRIMARY]"      
		   Set rsCheckSCExternalEmailMessages = cnnCheckSCExternalEmailMessages.Execute(SQLCheckSCExternalEmailMessages)
		   
		End If
	End If


	set rsCheckSCExternalEmailMessages = nothing
	cnnCheckSCExternalEmailMessages.close
	set cnnCheckSCExternalEmailMessages = nothing
				
%>