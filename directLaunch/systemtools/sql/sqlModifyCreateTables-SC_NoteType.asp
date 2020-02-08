<%
	Set cnnSC_NoteType = Server.CreateObject("ADODB.Connection")
	cnnSC_NoteType.open (Session("ClientCnnString"))
	Set rsSC_NoteType = Server.CreateObject("ADODB.Recordset")
	rsSC_NoteType.CursorLocation = 3 

	Err.Clear
	on error resume next
	Set rsSC_NoteType = cnnSC_NoteType.Execute("SELECT * FROM SC_NoteType")
	If Err.Description <> "" Then
		If InStr(Ucase(Err.Description),"INVALID OBJECT NAME") <> 0 Then
			On error goto 0
			'The table is not there, we need to create it
			SQLBuild = "CREATE TABLE SC_NoteType ("
			SQLBuild = SQLBuild & "[InternalRecordIdentifier] [int] IDENTITY(1,1) NOT NULL, "
			SQLBuild = SQLBuild & "[RecordCreationDateTime] [datetime] NULL CONSTRAINT [DF_SC_NoteType_RecordCreationDateTime]  DEFAULT (getdate()), "
			SQLBuild = SQLBuild & "[NoteType] [varchar] (255) NULL, "
			SQLBuild = SQLBuild & "[NoteTypeTabDisplaySortOrder] [int] NULL "
			SQLBuild = SQLBuild & ") ON [PRIMARY]"
			Set rsSC_NoteType = cnnSC_NoteType.Execute(SQLBuild)
			
			'Now Insert The Default Note Types
			SQL_SC_NoteType = "INSERT INTO SC_NoteType (NoteType) VALUES ('General')"
			Set rsSC_NoteType = cnnSC_NoteType.Execute(SQL_SC_NoteType)
			
			SQL_SC_NoteType = "INSERT INTO SC_NoteType (NoteType) VALUES ('Backend')"
			Set rsSC_NoteType = cnnSC_NoteType.Execute(SQL_SC_NoteType)
	
			SQL_SC_NoteType = "INSERT INTO SC_NoteType (NoteType) VALUES ('System')"
			Set rsSC_NoteType = cnnSC_NoteType.Execute(SQL_SC_NoteType)
	
			SQL_SC_NoteType = "INSERT INTO SC_NoteType (NoteType) VALUES ('MCS')"
			Set rsSC_NoteType = cnnSC_NoteType.Execute(SQL_SC_NoteType)
	
			SQL_SC_NoteType = "INSERT INTO SC_NoteType (NoteType) VALUES ('Service')"
			Set rsSC_NoteType = cnnSC_NoteType.Execute(SQL_SC_NoteType)
	
			SQL_SC_NoteType = "INSERT INTO SC_NoteType (NoteType) VALUES ('A/R')"
			Set rsSC_NoteType = cnnSC_NoteType.Execute(SQL_SC_NoteType)
	
			SQL_SC_NoteType = "INSERT INTO SC_NoteType (NoteType) VALUES ('CRM')"
			Set rsSC_NoteType = cnnSC_NoteType.Execute(SQL_SC_NoteType)	
					
		End If
	End If
	On Error Goto 0
	

	SQL_SC_NoteType = "SELECT COL_LENGTH('SC_NoteType', 'NoteTypeTabDisplaySortOrder') AS IsItThere"
	Set rsSC_NoteType = cnnSC_NoteType.Execute(SQL_SC_NoteType)
	If IsNull(rsSC_NoteType("IsItThere")) Then
		SQL_SC_NoteType = "ALTER TABLE SC_NoteType ADD NoteTypeTabDisplaySortOrder INT NULL"
		Set rsSC_NoteType= cnnSC_NoteType.Execute(SQL_SC_NoteType)
	End If

	SQL_SC_NoteType = "SELECT COL_LENGTH('SC_NoteType', 'NoteTypeCanBeCreatedByUser') AS IsItThere"
	Set rsSC_NoteType = cnnSC_NoteType.Execute(SQL_SC_NoteType)
	If IsNull(rsSC_NoteType("IsItThere")) Then
		SQL_SC_NoteType = "ALTER TABLE SC_NoteType ADD NoteTypeCanBeCreatedByUser INT NULL"
		Set rsSC_NoteType= cnnSC_NoteType.Execute(SQL_SC_NoteType)
	End If

	SQL_SC_NoteType = "UPDATE SC_NoteType SET NoteTypeCanBeCreatedByUser = 1"
	Set rsSC_NoteType= cnnSC_NoteType.Execute(SQL_SC_NoteType)

	SQL_SC_NoteType = "UPDATE SC_NoteType SET NoteTypeCanBeCreatedByUser = 0 WHERE NoteType='System' OR NoteType='Backend'"
	Set rsSC_NoteType= cnnSC_NoteType.Execute(SQL_SC_NoteType)

	Set rsSC_NoteType = Nothing
	cnnSC_NoteType.Close
	Set cnnSC_NoteType = Nothing
	
%>