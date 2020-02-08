<%	
	Set cnnSettings_Prospecting = Server.CreateObject("ADODB.Connection")
	cnnSettings_Prospecting.open (Session("ClientCnnString"))
	Set rsSettings_Prospecting = Server.CreateObject("ADODB.Recordset")
	rsSettings_Prospecting.CursorLocation = 3 


	Err.Clear
	on error resume next
	Set rsSettings_Prospecting = cnnSettings_Prospecting.Execute("SELECT TOP 1 * FROM Settings_Prospecting")
	If Err.Description <> "" Then
		If InStr(Ucase(Err.Description),"INVALID OBJECT NAME") <> 0 Then
			On error goto 0		
			'The table is not there, we need to create it
			SQLSettings_Prospecting = "CREATE TABLE [Settings_Prospecting]( "
			SQLSettings_Prospecting = SQLSettings_Prospecting & " [ShowLivePoolProspectSearchBox] [int] NULL, "
			SQLSettings_Prospecting = SQLSettings_Prospecting & " [ProspectActivityDefaultDaysToShow] [int] NULL, "
			SQLSettings_Prospecting = SQLSettings_Prospecting & " [TabSocialMediaColor] varchar(50) NULL "
			SQLSettings_Prospecting = SQLSettings_Prospecting & " [ProspectingWeeklyAgendaReportOnOff] [int] NULL, "
			SQLSettings_Prospecting = SQLSettings_Prospecting & " [ProspectingWeeklyAgendaReportEmailSubject] [varchar](255) NULL, "
			SQLSettings_Prospecting = SQLSettings_Prospecting & " [ProspectingWeeklyAgendaReportUserNos] [varchar](1000) NULL,"
			SQLSettings_Prospecting = SQLSettings_Prospecting & " [ProspectingWeeklyAgendaReportAdditionalEmails] [varchar](1000) NULL, "
			SQLSettings_Prospecting = SQLSettings_Prospecting & " [Schedule_ProspectingWeeklyAgendaReportGeneration] [varchar](1000) NULL "
			SQLSettings_Prospecting = SQLSettings_Prospecting & ") ON [PRIMARY]"
						
			Set rsSettings_Prospecting = cnnSettings_Prospecting.Execute(SQLSettings_Prospecting)
			
			SQLSettings_Prospecting = "INSERT INTO Settings_Prospecting (ShowLivePoolProspectSearchBox,ProspectActivityDefaultDaysToShow, ProspectingWeeklyAgendaReportOnOff"
			SQLSettings_Prospecting = SQLSettings_Prospecting & " VALUES "
			SQLSettings_Prospecting = SQLSettings_Prospecting & " (1,10,0) "
			Set rsSettings_Prospecting = cnnSettings_Prospecting.Execute(SQLSettings_Prospecting)
			
		End If
	End If

	On Error Goto 0
	SQLSettings_Prospecting = "SELECT COL_LENGTH('Settings_Prospecting', 'ProspectActivityDefaultDaysToShow') AS IsItThere"
	Set rsSettings_Prospecting = cnnSettings_Prospecting.Execute(SQLSettings_Prospecting)
	If IsNull(rsSettings_Prospecting("IsItThere")) Then
		SQLSettings_Prospecting = "ALTER TABLE Settings_Prospecting ADD ProspectActivityDefaultDaysToShow INT NULL"
		Set rsSettings_Prospecting = cnnSettings_Prospecting.Execute(SQLSettings_Prospecting)
		SQLSettings_Prospecting = "UPDATE Settings_Prospecting SET ProspectActivityDefaultDaysToShow = 10"
		Set rsSettings_Prospecting = cnnSettings_Prospecting.Execute(SQLSettings_Prospecting)
	End If
	
	
	SQLSettings_Prospecting = "SELECT COL_LENGTH('Settings_Prospecting', 'TabSocialMediaColor') AS IsItThere"
	Set rsSettings_Prospecting = cnnSettings_Prospecting.Execute(SQLSettings_Prospecting)
	If IsNull(rsSettings_Prospecting("IsItThere")) Then
		SQLSettings_Prospecting = "ALTER TABLE Settings_Prospecting ADD TabSocialMediaColor varchar(50) NULL"
		Set rsSettings_Prospecting = cnnSettings_Prospecting.Execute(SQLSettings_Prospecting)
		SQLSettings_Prospecting = "UPDATE Settings_Prospecting SET TabSocialMediaColor = '#000000'"
		Set rsSettings_Prospecting = cnnSettings_Prospecting.Execute(SQLSettings_Prospecting)
	End If


	
	' See if there are any records in the table & if not, insert one with default values
	Set rsSettings_Prospecting = cnnSettings_Prospecting.Execute("SELECT COUNT(*) AS Settings_Prospecting_Count FROM Settings_Prospecting")
	If rsSettings_Prospecting("Settings_Prospecting_Count") <> 1 Then
		SQLSettings_Prospecting = "INSERT INTO Settings_Prospecting (ShowLivePoolProspectSearchBox,ProspectActivityDefaultDaysToShow) "
		SQLSettings_Prospecting = SQLSettings_Prospecting & " VALUES "
		SQLSettings_Prospecting = SQLSettings_Prospecting & " (1,10)"
		Set rsSettings_Prospecting = cnnSettings_Prospecting.Execute(SQLSettings_Prospecting)
	End If


	SQLSettings_Prospecting = "SELECT COL_LENGTH('Settings_Prospecting', 'Schedule_ProspectingSnapshotReportGeneration') AS IsItThere"
	Set rsSettings_Prospecting = cnnSettings_Prospecting.Execute(SQLSettings_Prospecting)
	If IsNull(rsSettings_Prospecting("IsItThere")) Then
		SQLSettings_Prospecting = "ALTER TABLE Settings_Prospecting ADD Schedule_ProspectingSnapshotReportGeneration varchar(1000) NULL"
		Set rsSettings_Prospecting = cnnSettings_Prospecting.Execute(SQLSettings_Prospecting)
		SQLSettings_Prospecting = "UPDATE Settings_Prospecting SET Schedule_ProspectingSnapshotReportGeneration = '0,0,0,0,0,0,0,6:00 AM,6:00 AM,6:00 AM,6:00 AM,6:00 AM,6:00 AM,6:00 AM,0,0'"
		Set rsSettings_Prospecting = cnnSettings_Prospecting.Execute(SQLSettings_Prospecting)
	End If
	
	SQLSettings_Prospecting = "SELECT COL_LENGTH('Settings_Prospecting', 'ProspectingWeeklyAgendaReportOnOff') AS IsItThere"
	Set rsSettings_Prospecting = cnnSettings_Prospecting.Execute(SQLSettings_Prospecting)
	If IsNull(rsSettings_Prospecting("IsItThere")) Then
		SQLSettings_Prospecting = "ALTER TABLE Settings_Prospecting ADD ProspectingWeeklyAgendaReportOnOff int NULL"
		Set rsSettings_Prospecting = cnnSettings_Prospecting.Execute(SQLSettings_Prospecting)
		SQLSettings_Prospecting = "UPDATE Settings_Prospecting SET ProspectingWeeklyAgendaReportOnOff = 0"
		Set rsSettings_Prospecting = cnnSettings_Prospecting.Execute(SQLSettings_Prospecting)
	End If

	SQLSettings_Prospecting = "SELECT COL_LENGTH('Settings_Prospecting', 'ProspectingWeeklyAgendaReportEmailSubject') AS IsItThere"
	Set rsSettings_Prospecting = cnnSettings_Prospecting.Execute(SQLSettings_Prospecting)
	If IsNull(rsSettings_Prospecting("IsItThere")) Then
		SQLSettings_Prospecting = "ALTER TABLE Settings_Prospecting ADD ProspectingWeeklyAgendaReportEmailSubject varchar(1000)"
		Set rsSettings_Prospecting = cnnSettings_Prospecting.Execute(SQLSettings_Prospecting)
	End If

	SQLSettings_Prospecting = "SELECT COL_LENGTH('Settings_Prospecting', 'ProspectingWeeklyAgendaReportUserNos') AS IsItThere"
	Set rsSettings_Prospecting = cnnSettings_Prospecting.Execute(SQLSettings_Prospecting)
	If IsNull(rsSettings_Prospecting("IsItThere")) Then
		SQLSettings_Prospecting = "ALTER TABLE Settings_Prospecting ADD ProspectingWeeklyAgendaReportUserNos varchar(1000)"
		Set rsSettings_Prospecting = cnnSettings_Prospecting.Execute(SQLSettings_Prospecting)
	End If
	
	SQLSettings_Prospecting = "SELECT COL_LENGTH('Settings_Prospecting', 'ProspectingWeeklyAgendaReportAdditionalEmails') AS IsItThere"
	Set rsSettings_Prospecting = cnnSettings_Prospecting.Execute(SQLSettings_Prospecting)
	If IsNull(rsSettings_Prospecting("IsItThere")) Then
		SQLSettings_Prospecting = "ALTER TABLE Settings_Prospecting ADD ProspectingWeeklyAgendaReportAdditionalEmails varchar(1000)"
		Set rsSettings_Prospecting = cnnSettings_Prospecting.Execute(SQLSettings_Prospecting)
	End If

	SQLSettings_Prospecting = "SELECT COL_LENGTH('Settings_Prospecting', 'Schedule_ProspectingWeeklyAgendaReportGeneration') AS IsItThere"
	Set rsSettings_Prospecting = cnnSettings_Prospecting.Execute(SQLSettings_Prospecting)
	If IsNull(rsSettings_Prospecting("IsItThere")) Then
		SQLSettings_Prospecting = "ALTER TABLE Settings_Prospecting ADD Schedule_ProspectingWeeklyAgendaReportGeneration varchar(1000) NULL"
		Set rsSettings_Prospecting = cnnSettings_Prospecting.Execute(SQLSettings_Prospecting)
		SQLSettings_Prospecting = "UPDATE Settings_Prospecting SET Schedule_ProspectingWeeklyAgendaReportGeneration = '0,1,0,0,0,0,0,6:00 AM,6:00 AM,6:00 AM,6:00 AM,6:00 AM,6:00 AM,6:00 AM,0,0'"
		Set rsSettings_Prospecting = cnnSettings_Prospecting.Execute(SQLSettings_Prospecting)
	End If



	' This code makes sure scheduled process information is not NULL
	SQLSettings_Prospecting = "SELECT * FROM Settings_Prospecting"
	Set rsSettings_Prospecting = cnnSettings_Prospecting.Execute(SQLSettings_Prospecting)
	If NOT rsSettings_Prospecting.EOF Then
		If IsNull(rsSettings_Prospecting ("Schedule_ProspectingWeeklyAgendaReportGeneration")) Then
			SQLSettings_Prospecting = "UPDATE Settings_Prospecting SET Schedule_ProspectingWeeklyAgendaReportGeneration = '0,1,0,0,0,0,0,6:00 AM,6:00 AM,6:00 AM,6:00 AM,6:00 AM,6:00 AM,6:00 AM,0,0'"
			Set rsSettings_Prospecting = cnnSettings_Prospecting.Execute(SQLSettings_Prospecting)
		End If
	End If

	' This code makes sure scheduled process information is not NULL
	SQLSettings_Prospecting = "SELECT * FROM Settings_Prospecting"
	Set rsSettings_Prospecting = cnnSettings_Prospecting.Execute(SQLSettings_Prospecting)
	If NOT rsSettings_Prospecting.EOF Then
		If IsNull(rsSettings_Prospecting ("Schedule_ProspectingSnapshotReportGeneration")) Then
			SQLSettings_Prospecting = "UPDATE Settings_Prospecting SET Schedule_ProspectingSnapshotReportGeneration = '0,0,0,0,0,0,0,6:00 AM,6:00 AM,6:00 AM,6:00 AM,6:00 AM,6:00 AM,6:00 AM,0,0'"
			Set rsSettings_Prospecting = cnnSettings_Prospecting.Execute(SQLSettings_Prospecting)
		End If
	End If

	
	set rsSettings_Prospecting = nothing
	cnnSettings_Prospecting.close
	set cnnSettings_Prospecting = nothing
				
%>