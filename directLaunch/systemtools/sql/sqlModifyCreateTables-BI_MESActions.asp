<%	
	Set cnnBI_MESActions = Server.CreateObject("ADODB.Connection")
	cnnBI_MESActions.open (Session("ClientCnnString"))
	Set rsBI_MESActions = Server.CreateObject("ADODB.Recordset")
	rsBI_MESActions.CursorLocation = 3 


	Err.Clear
	on error resume next
	Set rsBI_MESActions = cnnBI_MESActions.Execute("SELECT TOP 1 * FROM BI_MESActions ORDER BY InternalRecordIdentifier DESC")
	If Err.Description <> "" Then
		If InStr(Ucase(Err.Description),"INVALID OBJECT NAME") <> 0 Then
			On error goto 0		
			'The table is not there, we need to create it
			SQLBI_MESActions = "CREATE TABLE [BI_MESActions]( "
			SQLBI_MESActions = SQLBI_MESActions & " [InternalRecordIdentifier] [int] IDENTITY(2,1) NOT NULL, "
			SQLBI_MESActions = SQLBI_MESActions & " [RecordCreationDateTime] [datetime] NULL CONSTRAINT [DF_biMESActions]  DEFAULT (getdate()), "
			SQLBI_MESActions = SQLBI_MESActions & " [CustID] [varchar](255) NULL, "
			SQLBI_MESActions = SQLBI_MESActions & " [MESMonth] [varchar](255) NULL, "
			SQLBI_MESActions = SQLBI_MESActions & " [Action] [varchar](255) NULL, "						
			SQLBI_MESActions = SQLBI_MESActions & " [ActionNotes] [varchar](8000) NULL "
			SQLBI_MESActions = SQLBI_MESActions & ") ON [PRIMARY]"
		
			Set rsBI_MESActions = cnnBI_MESActions.Execute(SQLBI_MESActions)
		End If
	End If

	SQL_BI_MESActions = "SELECT COL_LENGTH('BI_MESActions', 'MCSReasonIntRecID') AS IsItThere"
	Set rsBI_MESActions = cnnBI_MESActions.Execute(SQL_BI_MESActions)
	If IsNull(rsBI_MESActions("IsItThere")) Then
		SQL_BI_MESActions = "ALTER TABLE BI_MESActions ADD MCSReasonIntRecID INT NULL"
		Set rsBI_MESActions = cnnBI_MESActions.Execute(SQL_BI_MESActions)
	End If
	
	SQL_BI_MESActions = "SELECT COL_LENGTH('BI_MESActions', 'ActionNotes') AS IsItThere"
	Set rsBI_MESActions = cnnBI_MESActions.Execute(SQL_BI_MESActions)
	If IsNull(rsBI_MESActions("IsItThere")) Then
		SQL_BI_MESActions = "ALTER TABLE BI_MESActions ALTER COLUMN ActionNotes varchar(8000)"
		Set rsBI_MESActions = cnnBI_MESActions.Execute(SQL_BI_MESActions)
	End If

	
	set rsBI_MESActions = nothing
	cnnBI_MESActions.close
	set cnnBI_MESActions = nothing
				
%>