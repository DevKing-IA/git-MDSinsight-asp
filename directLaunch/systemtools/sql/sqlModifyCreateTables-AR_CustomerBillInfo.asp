<%
	Set cnnBillInfo = Server.CreateObject("ADODB.Connection")
	cnnBillInfo.open (Session("ClientCnnString"))
	Set rsBillInfo = Server.CreateObject("ADODB.Recordset")
	rsBillInfo.CursorLocation = 3 

	Err.Clear
	on error resume next
	Set rsBillInfo = cnnBillInfo.Execute("SELECT * FROM AR_CustomerBillInfo")
	If Err.Description <> "" Then
		If InStr(Ucase(Err.Description),"INVALID OBJECT NAME") <> 0 Then
			On error goto 0
			'The table is not there, we need to create it
			SQLBuild = "CREATE TABLE AR_CustomerBillInfo ("
			SQLBuild = SQLBuild & "[InternalRecordIdentifier] [int] IDENTITY(1,1) NOT NULL, "
			SQLBuild = SQLBuild & "[RecordCreationDateTime] [datetime] NULL CONSTRAINT [DF_AR_CustomerBillInfo_RecordCreationDateTime]  DEFAULT (getdate()), "
			SQLBuild = SQLBuild & "[CustID] [varchar] (255) NULL, "
			SQLBuild = SQLBuild & "[IncludeOnInvoices] [int] NULL CONSTRAINT [DF_AR_CustomerBillInfo_IncludeOnInvoices]  DEFAULT ((0)), "
			SQLBuild = SQLBuild & "[BillInfoFieldTitle] [varchar](8000) NULL, "
			SQLBuild = SQLBuild & "[BillInfoFieldData] [varchar](8000) NULL "
			SQLBuild = SQLBuild & ") ON [PRIMARY]"

			Set rsBillInfo = cnnBillInfo.Execute(SQLBuild)
			
		End If
	End If
	On Error Goto 0

	SQL_BillInfo = "SELECT COL_LENGTH('AR_CustomerBillInfo', 'CustID') AS IsItThere"
	Set rsBillInfo = cnnBillInfo.Execute(SQL_BillInfo)
	If NOT IsNull(rsBillInfo("IsItThere")) Then
		SQL_BillInfo = "SELECT DATA_TYPE FROM INFORMATION_SCHEMA.COLUMNS WHERE (TABLE_NAME = 'AR_CustomerBillInfo') AND (COLUMN_NAME = 'CustID')"
		Set rsBillInfo = cnnBillInfo.Execute(SQL_BillInfo)
		If Not rsBillInfo.Eof Then
			If rsBillInfo("DATA_TYPE")="varbinary" Then
				SQL_BillInfo = "ALTER TABLE AR_CustomerBillInfo DROP COLUMN CustID"
				Set rsBillInfo = cnnBillInfo.Execute(SQL_BillInfo)
			End If
		End If
	End If

	SQL_BillInfo = "SELECT COL_LENGTH('AR_CustomerBillInfo', 'CustID') AS IsItThere"
	Set rsBillInfo = cnnBillInfo.Execute(SQL_BillInfo)
	If IsNull(rsBillInfo("IsItThere")) Then
		SQL_BillInfo = "ALTER TABLE AR_CustomerBillInfo ADD CustID [varchar] (255) NULL"
		Set rsBillInfo = cnnBillInfo.Execute(SQL_BillInfo)
	End If

	set rsBillInfo = nothing
	cnnBillInfo.close
	set cnnBillInfo = nothing
%>