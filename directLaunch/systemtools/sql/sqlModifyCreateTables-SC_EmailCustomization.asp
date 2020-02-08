<%	

	Set cnnCheckEmailCustomization = Server.CreateObject("ADODB.Connection")
	cnnCheckEmailCustomization.open (Session("ClientCnnString"))
	Set rsCheckEmailCustomization = Server.CreateObject("ADODB.Recordset")
	rsCheckEmailCustomization.CursorLocation = 3 

	Err.Clear
	on error resume next
	Set rsCheckEmailCustomization = cnnCheckEmailCustomization.Execute("SELECT * FROM SC_EmailCustomization")
	If Err.Description <> "" Then
		If InStr(Ucase(Err.Description),"INVALID OBJECT NAME") <> 0 Then
			On error goto 0
			'The table is not there, we need to create it
			SQLBuild = "CREATE TABLE SC_EmailCustomization ("
			SQLBuild = SQLBuild & "[InternalRecordIdentifier] [int] IDENTITY(1,1) NOT NULL,"
			SQLBuild = SQLBuild & "	[RecordCreationDateTime] [datetime] NOT NULL CONSTRAINT [DF_SC_EmailCustomization_RecordCreationDateTime]  DEFAULT (getdate()),"
			SQLBuild = SQLBuild & "	[emailModule] [varchar](200) NULL,"
			SQLBuild = SQLBuild & "	[emailName] [varchar](500) NULL,"
			SQLBuild = SQLBuild & "	[emailDescription] [varchar](8000) NULL,"
			SQLBuild = SQLBuild & "	[emailSubjectLine] [varchar](200) NULL,"
			SQLBuild = SQLBuild & "	[emailSubheaderText] [varchar](200) NULL,"
			SQLBuild = SQLBuild & "	[emailBodyCodePart1] [varchar](8000) NULL,"
			SQLBuild = SQLBuild & "	[emailBodyCodePart2] [varchar](8000) NULL,"
			SQLBuild = SQLBuild & "	[emailBodyCodePart3] [varchar](8000) NULL,"
			SQLBuild = SQLBuild & "	[emailAssociatedLink] [varchar](500) NULL,"
			SQLBuild = SQLBuild & "	[emailAssociatedLinkButtonText] [varchar](100) NULL,"
			SQLBuild = SQLBuild & "	[emailType] [varchar](100) NULL,"
			SQLBuild = SQLBuild & "	[emailFileName] [varchar](200) NULL,"
			SQLBuild = SQLBuild & "	[customOrDefault] [varchar](100) NULL"
			SQLBuild = SQLBuild & ") ON [PRIMARY]"

			Set rsCheckEmailCustomization = cnnCheckEmailCustomization.Execute(SQLBuild)
		End If
	End If
	On Error Goto 0
			
	'See if these fields are in the table& add them if not there
	
	SQL_CheckEmailCustomization = "SELECT COL_LENGTH('SC_EmailCustomization', 'customOrDefault') AS IsItThere"
	Set rsCheckEmailCustomization = cnnCheckEmailCustomization.Execute(SQL_CheckEmailCustomization)
	If IsNull(rsCheckEmailCustomization("IsItThere")) Then
		SQL_CheckEmailCustomization = "ALTER TABLE SC_EmailCustomization ADD customOrDefault varchar(100) NULL"
		Set rsCheckEmailCustomization = cnnCheckEmailCustomization.Execute(SQL_CheckEmailCustomization)
		SQL_CheckEmailCustomization = "UPDATE SC_EmailCustomization SET customOrDefault = 'default'"
		Set rsCheckEmailCustomization = cnnCheckEmailCustomization.Execute(SQL_CheckEmailCustomization)
	End If

	Set rsCheckEmailCustomization = Nothing
	cnnCheckEmailCustomization.Close
	Set cnnCheckEmailCustomization = Nothing
				
%>