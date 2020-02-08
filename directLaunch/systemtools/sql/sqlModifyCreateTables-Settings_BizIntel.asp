<%

	Set cnnBizIntel = Server.CreateObject("ADODB.Connection")
	cnnBizIntel.open (Session("ClientCnnString"))
	Set rsBizIntel = Server.CreateObject("ADODB.Recordset")
	rsBizIntel.CursorLocation = 3 

	Err.Clear
	on error resume next
	Set rsBizIntel = cnnBizIntel.Execute("SELECT * FROM Settings_BizIntel")
	If Err.Description <> "" Then
		If InStr(Ucase(Err.Description),"INVALID OBJECT NAME") <> 0 Then
			On error goto 0
			'The table is not there, we need to create it
			SQLBuild = "CREATE TABLE Settings_BizIntel ("
			SQLBuild = SQLBuild & "[CustAnalSum1OnOff] [int] NULL CONSTRAINT [DF_Settings_BizIntel_CustAnalSum1OnOff]  DEFAULT ((0)), "
			SQLBuild = SQLBuild & "[CustAnalSum1EmailToUserNos] [varchar](1000) NULL, "
			SQLBuild = SQLBuild & "[CustAnalSum1UserNosToCC] [varchar](1000) NULL, "
			SQLBuild = SQLBuild & "[CustAnalSum1EmailAddressesToCC] [varchar](8000) NULL "
			SQLBuild = SQLBuild & ") ON [PRIMARY]"

			Set rsBizIntel = cnnBizIntel.Execute(SQLBuild)
			
			
		End If
	End If
	On Error Goto 0
	
	SQL_BizIntel = "SELECT COL_LENGTH('Settings_BizIntel', 'CustAnalSum1OnOff') AS IsItThere"
	Set rsBizIntel = cnnBizIntel.Execute(SQL_BizIntel)
	If IsNull(rsBizIntel("IsItThere")) Then
		SQL_BizIntel = "ALTER TABLE Settings_BizIntel ADD CustAnalSum1OnOff [int] NULL "
		Set rsBizIntel = cnnBizIntel.Execute(SQL_BizIntel)
		
		SQL_BizIntel = "ALTER TABLE [Settings_BizIntel] ADD CONSTRAINT [DF_Settings_BizIntel_CustAnalSum1OnOff]  DEFAULT ((0)) FOR [CustAnalSum1OnOff]"
		Set rsBizIntel = cnnBizIntel.Execute(SQL_BizIntel)

	End If

	SQL_BizIntel = "SELECT COL_LENGTH('Settings_BizIntel', 'CustAnalSum1EmailToUserNos') AS IsItThere"
	Set rsBizIntel = cnnBizIntel.Execute(SQL_BizIntel)
	If IsNull(rsBizIntel("IsItThere")) Then
		SQL_BizIntel = "ALTER TABLE Settings_BizIntel ADD CustAnalSum1EmailToUserNos [varchar](1000) NULL "
		Set rsBizIntel = cnnBizIntel.Execute(SQL_BizIntel)
	End If

	SQL_BizIntel = "SELECT COL_LENGTH('Settings_BizIntel', 'CustAnalSum1UserNosToCC') AS IsItThere"
	Set rsBizIntel = cnnBizIntel.Execute(SQL_BizIntel)
	If IsNull(rsBizIntel("IsItThere")) Then
		SQL_BizIntel = "ALTER TABLE Settings_BizIntel ADD CustAnalSum1UserNosToCC [varchar](1000) NULL "
		Set rsBizIntel = cnnBizIntel.Execute(SQL_BizIntel)
	End If

	SQL_BizIntel = "SELECT COL_LENGTH('Settings_BizIntel', 'CustAnalSum1EmailAddressesToCC') AS IsItThere"
	Set rsBizIntel = cnnBizIntel.Execute(SQL_BizIntel)
	If IsNull(rsBizIntel("IsItThere")) Then
		SQL_BizIntel = "ALTER TABLE Settings_BizIntel ADD CustAnalSum1EmailAddressesToCC [varchar](8000) NULL "
		Set rsBizIntel = cnnBizIntel.Execute(SQL_BizIntel)
	End If

	SQL_BizIntel = "SELECT COL_LENGTH('Settings_BizIntel', 'MCSUserNosToCC') AS IsItThere"
	Set rsBizIntel = cnnBizIntel.Execute(SQL_BizIntel)
	If IsNull(rsBizIntel("IsItThere")) Then
		SQL_BizIntel = "ALTER TABLE Settings_BizIntel ADD MCSUserNosToCC [varchar](1000) NULL "
		Set rsBizIntel = cnnBizIntel.Execute(SQL_BizIntel)
	End If

	'Drop fields that had the old names
	SQL_BizIntel = "SELECT COL_LENGTH('Settings_BizIntel', 'SalesVarianceReportOnOff') AS IsItThere"
	Set rsBizIntel = cnnBizIntel.Execute(SQL_BizIntel)
	If NOT IsNull(rsBizIntel("IsItThere")) Then

		SQL_BizIntel = "ALTER TABLE [Settings_BizIntel] DROP CONSTRAINT [DF_Settings_BizIntel_SalesVarianceReportOnOff]"
		Set rsBizIntel = cnnBizIntel.Execute(SQL_BizIntel)

		SQL_BizIntel = "ALTER TABLE Settings_BizIntel DROP COLUMN SalesVarianceReportOnOff"
		Set rsBizIntel = cnnBizIntel.Execute(SQL_BizIntel)
	End If

	SQL_BizIntel = "SELECT COL_LENGTH('Settings_BizIntel', 'EmailToUserNos') AS IsItThere"
	Set rsBizIntel = cnnBizIntel.Execute(SQL_BizIntel)
	If NOT IsNull(rsBizIntel("IsItThere")) Then
		SQL_BizIntel = "ALTER TABLE Settings_BizIntel DROP COLUMN EmailToUserNos"
		Set rsBizIntel = cnnBizIntel.Execute(SQL_BizIntel)
	End If
	
	SQL_BizIntel = "SELECT COL_LENGTH('Settings_BizIntel', 'UserNosToCC') AS IsItThere"
	Set rsBizIntel = cnnBizIntel.Execute(SQL_BizIntel)
	If NOT IsNull(rsBizIntel("IsItThere")) Then
		SQL_BizIntel = "ALTER TABLE Settings_BizIntel DROP COLUMN UserNosToCC"
		Set rsBizIntel = cnnBizIntel.Execute(SQL_BizIntel)
	End If
	
	SQL_BizIntel = "SELECT COL_LENGTH('Settings_BizIntel', 'EmailAddressesToCC') AS IsItThere"
	Set rsBizIntel = cnnBizIntel.Execute(SQL_BizIntel)
	If NOT IsNull(rsBizIntel("IsItThere")) Then
		SQL_BizIntel = "ALTER TABLE Settings_BizIntel DROP COLUMN EmailAddressesToCC"
		Set rsBizIntel = cnnBizIntel.Execute(SQL_BizIntel)
	End If

	SQL_BizIntel = "SELECT COL_LENGTH('Settings_BizIntel', 'MCSActivitySummaryOnOff') AS IsItThere"
	Set rsBizIntel = cnnBizIntel.Execute(SQL_BizIntel)
	If IsNull(rsBizIntel("IsItThere")) Then
		SQL_BizIntel = "ALTER TABLE Settings_BizIntel ADD MCSActivitySummaryOnOff [int] NULL "
		Set rsBizIntel = cnnBizIntel.Execute(SQL_BizIntel)
		
		SQL_BizIntel = "ALTER TABLE [Settings_BizIntel] ADD CONSTRAINT [DF_Settings_BizIntel_MCSActivitySummaryOnOff]  DEFAULT ((0)) FOR [MCSActivitySummaryOnOff]"
		Set rsBizIntel = cnnBizIntel.Execute(SQL_BizIntel)

	End If

	SQL_BizIntel = "SELECT COL_LENGTH('Settings_BizIntel', 'MCSActivitySummaryEmailToUserNos') AS IsItThere"
	Set rsBizIntel = cnnBizIntel.Execute(SQL_BizIntel)
	If IsNull(rsBizIntel("IsItThere")) Then
		SQL_BizIntel = "ALTER TABLE Settings_BizIntel ADD MCSActivitySummaryEmailToUserNos [varchar](1000) NULL "
		Set rsBizIntel = cnnBizIntel.Execute(SQL_BizIntel)
	End If

	SQL_BizIntel = "SELECT COL_LENGTH('Settings_BizIntel', 'MCSActivitySummaryUserNosToCC') AS IsItThere"
	Set rsBizIntel = cnnBizIntel.Execute(SQL_BizIntel)
	If IsNull(rsBizIntel("IsItThere")) Then
		SQL_BizIntel = "ALTER TABLE Settings_BizIntel ADD MCSActivitySummaryUserNosToCC [varchar](1000) NULL "
		Set rsBizIntel = cnnBizIntel.Execute(SQL_BizIntel)
	End If

	SQL_BizIntel = "SELECT COL_LENGTH('Settings_BizIntel', 'MCSActivitySummaryEmailAddressesToCC') AS IsItThere"
	Set rsBizIntel = cnnBizIntel.Execute(SQL_BizIntel)
	If IsNull(rsBizIntel("IsItThere")) Then
		SQL_BizIntel = "ALTER TABLE Settings_BizIntel ADD MCSActivitySummaryEmailAddressesToCC [varchar](8000) NULL "
		Set rsBizIntel = cnnBizIntel.Execute(SQL_BizIntel)
	End If

	SQL_BizIntel = "SELECT COL_LENGTH('Settings_BizIntel', 'MCSUseAlternateHeader') AS IsItThere"
	Set rsBizIntel = cnnBizIntel.Execute(SQL_BizIntel)
	If IsNull(rsBizIntel("IsItThere")) Then
		SQL_BizIntel = "ALTER TABLE Settings_BizIntel ADD MCSUseAlternateHeader[int] NULL "
		Set rsBizIntel = cnnBizIntel.Execute(SQL_BizIntel)
		
		SQL_BizIntel = "ALTER TABLE [Settings_BizIntel] ADD CONSTRAINT [DF_Settings_BizIntel_MCSUseAlternateHeader]  DEFAULT ((0)) FOR [MCSUseAlternateHeader]"
		Set rsBizIntel = cnnBizIntel.Execute(SQL_BizIntel)

		SQL_BizIntel = "UPDATE [Settings_BizIntel] SET MCSUseAlternateHeader = 0 "
		Set rsBizIntel = cnnBizIntel.Execute(SQL_BizIntel)

	End If

	SQL_BizIntel = "SELECT COL_LENGTH('Settings_BizIntel', 'Schedule_AutomaticCustomerAnalysisSummary1ReportGeneration') AS IsItThere"
	Set rsBizIntel = cnnBizIntel.Execute(SQL_BizIntel)
	If IsNull(rsBizIntel("IsItThere")) Then
		SQL_BizIntel = "ALTER TABLE Settings_BizIntel ADD Schedule_AutomaticCustomerAnalysisSummary1ReportGeneration [varchar](1000) NULL "
		Set rsBizIntel = cnnBizIntel.Execute(SQL_BizIntel)
		SQL_BizIntel = "UPDATE Settings_BizIntel SET Schedule_AutomaticCustomerAnalysisSummary1ReportGeneration = '0,0,0,0,0,0,0,10:00 AM,10:00 AM,10:00 AM,10:00 AM,10:00 AM,10:00 AM,10:00 AM,0,0'"
		Set rsBizIntel = cnnBizIntel.Execute(SQL_BizIntel)
	End If

	SQL_BizIntel = "SELECT COL_LENGTH('Settings_BizIntel', 'Schedule_MCSActivityReportGeneration') AS IsItThere"
	Set rsBizIntel = cnnBizIntel.Execute(SQL_BizIntel)
	If IsNull(rsBizIntel("IsItThere")) Then
		SQL_BizIntel = "ALTER TABLE Settings_BizIntel ADD Schedule_MCSActivityReportGeneration [varchar](1000) NULL "
		Set rsBizIntel = cnnBizIntel.Execute(SQL_BizIntel)
		SQL_BizIntel = "UPDATE Settings_BizIntel SET Schedule_MCSActivityReportGeneration = '0,0,0,0,0,0,0,10:00 AM,10:00 AM,10:00 AM,10:00 AM,10:00 AM,10:00 AM,10:00 AM,0,0'"
		Set rsBizIntel = cnnBizIntel.Execute(SQL_BizIntel)
	End If

	set rsBizIntel = nothing
	cnnBizIntel.close
	set cnnBizIntel = nothing

On error goto 0
	'**********************************************************************************
	'This code should be at the bottom after all the other structure change code is run
	'It will see if there is a record in the table & if not it will insert one with
	'a dummy value
	'Then it runs an update where all the proper default values are set
	'It was only done this way to make the code more readable
	'**********************************************************************************	
	
	Set cnnBizIntel = Server.CreateObject("ADODB.Connection")
	cnnBizIntel.open (Session("ClientCnnString"))
	Set rsBizIntel = Server.CreateObject("ADODB.Recordset")

	SQL_BizIntel = "SELECT COUNT (*) AS BizIntelCount FROM Settings_BizIntel"
	Set rsBizIntel = cnnBizIntel.Execute(SQL_BizIntel)

	If rsBizIntel("BizIntelCount") <> 1 Then

		SQL_BizIntel = "INSERT INTO Settings_BizIntel (CustAnalSum1OnOff) VALUES (0)"
		Set rsBizIntel = cnnBizIntel.Execute(SQL_BizIntel)

		SQL_BizIntel = "UPDATE Settings_BizIntel SET "
		SQL_BizIntel = SQL_BizIntel & "  CustAnalSum1OnOff = 0"
		SQL_BizIntel = SQL_BizIntel & ",  CustAnalSum1EmailToUserNos = ''"
		SQL_BizIntel = SQL_BizIntel & ",  CustAnalSum1UserNosToCC = ''"
		SQL_BizIntel = SQL_BizIntel & ",  CustAnalSum1EmailAddressesToCC = ''"
		SQL_BizIntel = SQL_BizIntel & ",  MCSUserNosToCC = ''"
		SQL_BizIntel = SQL_BizIntel & ",  MCSActivitySummaryOnOff = 0"
		SQL_BizIntel = SQL_BizIntel & ",  MCSActivitySummaryEmailToUserNos = ''"
		SQL_BizIntel = SQL_BizIntel & ",  MCSActivitySummaryUserNosToCC = ''"
		SQL_BizIntel = SQL_BizIntel & ",  MCSActivitySummaryEmailAddressesToCC = ''"

		Set rsBizIntel = cnnBizIntel.Execute(SQL_BizIntel)
	
	End If
	

	' This code makes sure scheduled process information is not NULL
	SQL_BizIntel = "SELECT * FROM Settings_BizIntel"
	Set rsBizIntel = cnnBizIntel.Execute(SQL_BizIntel)
	If NOT rsBizIntel.EOF Then
		If IsNull(rsBizIntel("Schedule_AutomaticCustomerAnalysisSummary1ReportGeneration")) Then
			SQL_BizIntel = "UPDATE Settings_BizIntel SET Schedule_AutomaticCustomerAnalysisSummary1ReportGeneration = '0,0,0,0,0,0,0,10:00 AM,10:00 AM,10:00 AM,10:00 AM,10:00 AM,10:00 AM,10:00 AM,0,0'"
			Set rsBizIntel = cnnBizIntel.Execute(SQL_BizIntel)
		End If
	End If

	Set rsBizIntel = Nothing
	cnnBizIntel.Close
	Set cnnBizIntel = Nothing
			
%>