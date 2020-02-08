<%

	Set cnnNeedToKnow = Server.CreateObject("ADODB.Connection")
	cnnNeedToKnow.open (Session("ClientCnnString"))
	Set rsNeedToKnow = Server.CreateObject("ADODB.Recordset")
	rsNeedToKnow.CursorLocation = 3 

	Err.Clear
	on error resume next
	Set rsNeedToKnow = cnnNeedToKnow.Execute("SELECT * FROM SC_NeedToKnow")
	If Err.Description <> "" Then
		If InStr(Ucase(Err.Description),"INVALID OBJECT NAME") <> 0 Then
			On error goto 0
			'The table is not there, we need to create it
			SQLBuild = "CREATE TABLE SC_NeedToKnow("
			SQLBuild = SQLBuild & "InternalRecordIdentifier [int] IDENTITY(1,1) NOT NULL, "
			SQLBuild = SQLBuild & "RecordCreationDateTime [datetime] NULL CONSTRAINT [DF_SC_NeedToKnow_RecordCreationDateTime] DEFAULT (getdate()), "
			SQLBuild = SQLBuild & "[Module] [varchar](255) NULL, "
			SQLBuild = SQLBuild & "[SummaryDescription] [varchar](255) NULL, "
			SQLBuild = SQLBuild & "[SummaryURL] [varchar](8000) NULL, "
			SQLBuild = SQLBuild & "[DetailedDescription1] [varchar](8000) NULL, "
			SQLBuild = SQLBuild & "[DetailedDescription2] [varchar](8000) NULL, "
			SQLBuild = SQLBuild & "[DetailedDescription3] [varchar](8000) NULL, "
			SQLBuild = SQLBuild & "[DetailURL] [varchar](8000) NULL, "
			SQLBuild = SQLBuild & "[InsightStaffOnly] [int] NULL "
			SQLBuild = SQLBuild & ") ON [PRIMARY]"
			Set rsNeedToKnow = cnnNeedToKnow.Execute(SQLBuild)
		End If
	End If
	On Error Goto 0


	SQLBuild = "SELECT COL_LENGTH('SC_NeedToKnow', 'CustIDIfApplicable') AS IsItThere"
	Set rsNeedToKnow = cnnNeedToKnow.Execute(SQLBuild)
	If IsNull(rsNeedToKnow("IsItThere")) Then
		SQLBuild  = "ALTER TABLE SC_NeedToKnow ADD CustIDIfApplicable varchar(255) NULL"
		Set rsNeedToKnow = cnnNeedToKnow.Execute(SQLBuild)
	End If

	SQLBuild = "SELECT COL_LENGTH('SC_NeedToKnow', 'EquipIDIfApplicable') AS IsItThere"
	Set rsNeedToKnow = cnnNeedToKnow.Execute(SQLBuild)
	If IsNull(rsNeedToKnow("IsItThere")) Then
		SQLBuild  = "ALTER TABLE SC_NeedToKnow ADD EquipIDIfApplicable varchar(255) NULL"
		Set rsNeedToKnow = cnnNeedToKnow.Execute(SQLBuild)
	End If
	
	SQLBuild = "SELECT COL_LENGTH('SC_NeedToKnow', 'SubModule') AS IsItThere"
	Set rsNeedToKnow = cnnNeedToKnow.Execute(SQLBuild)
	If IsNull(rsNeedToKnow("IsItThere")) Then
		SQLBuild  = "ALTER TABLE SC_NeedToKnow ADD SubModule varchar(255) NULL"
		Set rsNeedToKnow = cnnNeedToKnow.Execute(SQLBuild)
	End If
	
	SQLBuild = "SELECT COL_LENGTH('SC_NeedToKnow', 'prodSKUIfApplicable') AS IsItThere"
	Set rsNeedToKnow = cnnNeedToKnow.Execute(SQLBuild)
	If IsNull(rsNeedToKnow("IsItThere")) Then
		SQLBuild  = "ALTER TABLE SC_NeedToKnow ADD prodSKUIfApplicable varchar(255) NULL"
		Set rsNeedToKnow = cnnNeedToKnow.Execute(SQLBuild)
	End If

	SQLBuild = "SELECT COL_LENGTH('SC_NeedToKnow', 'prodUPCIfApplicable') AS IsItThere"
	Set rsNeedToKnow = cnnNeedToKnow.Execute(SQLBuild)
	If IsNull(rsNeedToKnow("IsItThere")) Then
		SQLBuild  = "ALTER TABLE SC_NeedToKnow ADD prodUPCIfApplicable varchar(255) NULL"
		Set rsNeedToKnow = cnnNeedToKnow.Execute(SQLBuild)
	End If

	SQLBuild = "SELECT COL_LENGTH('SC_NeedToKnow', 'prodBinIfApplicable') AS IsItThere"
	Set rsNeedToKnow = cnnNeedToKnow.Execute(SQLBuild)
	If IsNull(rsNeedToKnow("IsItThere")) Then
		SQLBuild  = "ALTER TABLE SC_NeedToKnow ADD prodBinIfApplicable varchar(255) NULL"
		Set rsNeedToKnow = cnnNeedToKnow.Execute(SQLBuild)
	End If

	set rsNeedToKnow = nothing
	cnnNeedToKnow.close
	set cnnNeedToKnow = nothing
			
%>