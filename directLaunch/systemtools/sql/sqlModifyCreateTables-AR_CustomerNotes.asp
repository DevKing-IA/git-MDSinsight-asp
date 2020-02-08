<%
	Set cnnAR_CustomerNotes = Server.CreateObject("ADODB.Connection")
	cnnAR_CustomerNotes.open (Session("ClientCnnString"))
	Set rsAR_CustomerNotes = Server.CreateObject("ADODB.Recordset")
	rsAR_CustomerNotes.CursorLocation = 3 

	Err.Clear
	on error resume next
	Set rsAR_CustomerNotes = cnnAR_CustomerNotes.Execute("SELECT * FROM AR_CustomerNotes")
	If Err.Description <> "" Then
		If InStr(Ucase(Err.Description),"INVALID OBJECT NAME") <> 0 Then
			On error goto 0
			'The table is not there, we need to create it
			SQLBuild = "CREATE TABLE AR_CustomerNotes ("
			SQLBuild = SQLBuild & "[InternalRecordIdentifier] [int] IDENTITY(1,1) NOT NULL, "
			SQLBuild = SQLBuild & "[RecordCreationDateTime] [datetime] NULL CONSTRAINT [DF_AR_CustomerAR_CustomerNotes_RecordCreationDateTime]  DEFAULT (getdate()), "
			SQLBuild = SQLBuild & "[CustID] [varchar] (255) NULL, "
			SQLBuild = SQLBuild & "[Category] [int] NULL, "
			SQLBuild = SQLBuild & "[EnteredByUserNo] [int] NULL, "
			SQLBuild = SQLBuild & "[Note] [varchar](8000) NULL "
			SQLBuild = SQLBuild & ") ON [PRIMARY]"

			Set rsAR_CustomerNotes = cnnAR_CustomerNotes.Execute(SQLBuild)
			
		End If
	End If
	On Error Goto 0
	
	
	SQLBuild = "SELECT COL_LENGTH('AR_CustomerNotes', 'MCSReasonIntRecID') AS IsItThere"
	Set rsAR_CustomerNotes = cnnAR_CustomerNotes.Execute(SQLBuild)
	If IsNull(rsAR_CustomerNotes("IsItThere")) Then
		SQLBuild  = "ALTER TABLE AR_CustomerNotes ADD MCSReasonIntRecID int NULL"
		Set rsAR_CustomerNotes = cnnAR_CustomerNotes.Execute(SQLBuild)
	End If
	
	SQLBuild = "SELECT COL_LENGTH('AR_CustomerNotes', 'NoteTypeIntRecID') AS IsItThere"
	Set rsAR_CustomerNotes = cnnAR_CustomerNotes.Execute(SQLBuild)
	If IsNull(rsAR_CustomerNotes("IsItThere")) Then
		SQLBuild  = "ALTER TABLE AR_CustomerNotes ADD NoteTypeIntRecID int NULL"
		Set rsAR_CustomerNotes = cnnAR_CustomerNotes.Execute(SQLBuild)
	End If

	set rsAR_CustomerNotes = nothing
	cnnAR_CustomerNotes.close
	set cnnAR_CustomerNotes = nothing
%>