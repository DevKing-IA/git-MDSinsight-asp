<%
	Set AccountingPeriods = Server.CreateObject("ADODB.Connection")
	AccountingPeriods.open (Session("ClientCnnString"))
	Set rsBizIntel = Server.CreateObject("ADODB.Recordset")
	rsBizIntel.CursorLocation = 3 

	Err.Clear
	on error resume next
	Set rsBizIntel = AccountingPeriods.Execute("SELECT * FROM Settings_AccountingPeriods")
	If Err.Description <> "" Then
		If InStr(Ucase(Err.Description),"INVALID OBJECT NAME") <> 0 Then
			On error goto 0
			'The table is not there, we need to create it
			SQLBuild = "CREATE TABLE Settings_AccountingPeriods ("
			SQLBuild = SQLBuild & "[InternalRecordIdentifier] [int] IDENTITY(1,1) NOT NULL, "
			SQLBuild = SQLBuild & "[RecordCreationDateTime] [datetime] NULL CONSTRAINT [DF_Settings_AccountingPeriods_RecordCreationDateTime]  DEFAULT (getdate()), "
			SQLBuild = SQLBuild & "[PeriodYear] [int] NULL, "
			SQLBuild = SQLBuild & "[Period] [int] NULL, "
			SQLBuild = SQLBuild & "[BeginDate] [date] NULL, "
			SQLBuild = SQLBuild & "[EndDate] [date] NULL "
			SQLBuild = SQLBuild & ") ON [PRIMARY]"

			Set rsAccountingPeriods = AccountingPeriods.Execute(SQLBuild)
			
		End If
	End If
	On Error Goto 0
	

	Set rsAccountingPeriods = Nothing
	AccountingPeriods.Close
	Set AccountingPeriods = Nothing
		
%>