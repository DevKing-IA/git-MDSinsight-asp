<%	
	Response.Write("sqlModifyCreateTables-SC_ExternalEmailMessagesBody.asp" & "<br>")
	On Error Goto 0

	Set cnnCheckSCExternalEmailMessagesBody = Server.CreateObject("ADODB.Connection")
	cnnCheckSCExternalEmailMessagesBody.open (Session("ClientCnnString"))
	Set rsCheckSCExternalEmailMessagesBody = Server.CreateObject("ADODB.Recordset")

	Err.Clear
	on error resume next
	Set rsCheckSCExternalEmailMessagesBody = cnnCheckSCExternalEmailMessagesBody.Execute("SELECT TOP 1 * FROM SC_ExternalEmailMessagesBody")
	If Err.Description <> "" Then
		If InStr(Ucase(Err.Description),"INVALID OBJECT NAME") <> 0 Then
			On error goto 0		
			'The table is not there, we need to create it
			
		    SQLCheckSCExternalEmailMessagesBody = "CREATE TABLE [SC_ExternalEmailMessagesBody]("
			SQLCheckSCExternalEmailMessagesBody = SQLCheckSCExternalEmailMessagesBody & " [InternalRecordIdentifier] [int] IDENTITY(1,1) NOT NULL,"
			SQLCheckSCExternalEmailMessagesBody = SQLCheckSCExternalEmailMessagesBody & " [RecordCreationDateTime] [datetime] NULL CONSTRAINT [DF_SC_ExternalEmailMessagesBody_RecordCreationDateTime]  DEFAULT (getdate()),"
			SQLCheckSCExternalEmailMessagesBody = SQLCheckSCExternalEmailMessagesBody & " [MessageID] [varchar](8000) NULL, "
			SQLCheckSCExternalEmailMessagesBody = SQLCheckSCExternalEmailMessagesBody & " [MessageBody] [varchar](8000) NULL "
			SQLCheckSCExternalEmailMessagesBody = SQLCheckSCExternalEmailMessagesBody & " ) ON [PRIMARY]"     
			 
		    Set rsCheckSCExternalEmailMessagesBody = cnnCheckSCExternalEmailMessagesBody.Execute(SQLCheckSCExternalEmailMessagesBody)
		   
		End If
	End If

	
	set rsCheckSCExternalEmailMessagesBody = nothing
	cnnCheckSCExternalEmailMessagesBody.close
	set cnnCheckSCExternalEmailMessagesBody = nothing
				
%>