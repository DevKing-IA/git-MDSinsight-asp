<%	
	Set cnnBI_MCSActions = Server.CreateObject("ADODB.Connection")
	cnnBI_MCSActions.open (Session("ClientCnnString"))
	Set rsBI_MCSActions = Server.CreateObject("ADODB.Recordset")
	rsBI_MCSActions.CursorLocation = 3 


	Err.Clear
	on error resume next
	Set rsBI_MCSActions = cnnBI_MCSActions.Execute("SELECT TOP 1 * FROM BI_MCSActions ORDER BY InternalRecordIdentifier DESC")
	If Err.Description <> "" Then
		If InStr(Ucase(Err.Description),"INVALID OBJECT NAME") <> 0 Then
			On error goto 0		
			'The table is not there, we need to create it
			SQLBI_MCSActions = "CREATE TABLE [BI_MCSActions]( "
			SQLBI_MCSActions = SQLBI_MCSActions & " [InternalRecordIdentifier] [int] IDENTITY(2,1) NOT NULL, "
			SQLBI_MCSActions = SQLBI_MCSActions & " [RecordCreationDateTime] [datetime] NULL CONSTRAINT [DF_bimcsactions]  DEFAULT (getdate()), "
			SQLBI_MCSActions = SQLBI_MCSActions & " [CustID] [varchar](255) NULL, "
			SQLBI_MCSActions = SQLBI_MCSActions & " [MCSMonth] [varchar](255) NULL, "
			SQLBI_MCSActions = SQLBI_MCSActions & " [Action] [varchar](255) NULL, "						
			SQLBI_MCSActions = SQLBI_MCSActions & " [ActionNotes] [varchar](8000) NULL "
			SQLBI_MCSActions = SQLBI_MCSActions & ") ON [PRIMARY]"
		
			Set rsBI_MCSActions = cnnBI_MCSActions.Execute(SQLBI_MCSActions)
		End If
	End If

	SQL_BI_MCSActions = "SELECT COL_LENGTH('BI_MCSActions', 'MCSReasonIntRecID') AS IsItThere"
	Set rsBI_MCSActions = cnnBI_MCSActions.Execute(SQL_BI_MCSActions)
	If IsNull(rsBI_MCSActions("IsItThere")) Then
		SQL_BI_MCSActions = "ALTER TABLE BI_MCSActions ADD MCSReasonIntRecID INT NULL"
		Set rsBI_MCSActions = cnnBI_MCSActions.Execute(SQL_BI_MCSActions)
	End If
	
	SQL_BI_MCSActions = "SELECT COL_LENGTH('BI_MCSActions', 'ActionNotes') AS IsItThere"
	Set rsBI_MCSActions = cnnBI_MCSActions.Execute(SQL_BI_MCSActions)
	If IsNull(rsBI_MCSActions("IsItThere")) Then
		SQL_BI_MCSActions = "ALTER TABLE BI_MCSActions ALTER COLUMN ActionNotes varchar(8000)"
		Set rsBI_MCSActions = cnnBI_MCSActions.Execute(SQL_BI_MCSActions)
	End If

	
	set rsBI_MCSActions = nothing
	cnnBI_MCSActions.close
	set cnnBI_MCSActions = nothing
				
%>