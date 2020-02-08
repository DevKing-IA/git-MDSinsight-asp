<%	
	Set cnnSettings_InventoryControl = Server.CreateObject("ADODB.Connection")
	cnnSettings_InventoryControl.open (Session("ClientCnnString"))
	Set rsSettings_InventoryControl = Server.CreateObject("ADODB.Recordset")
	rsSettings_InventoryControl.CursorLocation = 3 


	Err.Clear
	on error resume next
	Set rsSettings_InventoryControl = cnnSettings_InventoryControl.Execute("SELECT * FROM Settings_InventoryControl")
	If Err.Description <> "" Then
		If InStr(Ucase(Err.Description),"INVALID OBJECT NAME") <> 0 Then
			On error goto 0		
			'The table is not there, we need to create it
			SQLSettings_InventoryControl = "CREATE TABLE [Settings_InventoryControl]( "
			SQLSettings_InventoryControl = SQLSettings_InventoryControl & " [InternalRecordIdentifier] [int] IDENTITY(2,1) NOT NULL, "
			SQLSettings_InventoryControl = SQLSettings_InventoryControl & " [RecordCreationDateTime] [datetime] NULL CONSTRAINT [DF_Settings_InventoryControl]  DEFAULT (getdate()), "
			SQLSettings_InventoryControl = SQLSettings_InventoryControl & " [InventoryAPIRepostONOFF] [int] NOT NULL DEFAULT 0, "
			SQLSettings_InventoryControl = SQLSettings_InventoryControl & " [InventoryAPIRepostMode] [varchar](10) NULL, "
			SQLSettings_InventoryControl = SQLSettings_InventoryControl & " [InventoryAPIRepostURL] [varchar](255) NULL, "
			SQLSettings_InventoryControl = SQLSettings_InventoryControl & " [InventoryAPIRepostOnHandONOFF] [int] NOT NULL DEFAULT 0, "
			SQLSettings_InventoryControl = SQLSettings_InventoryControl & " [InventoryAPIRepostOnHandMode] [varchar](10) NULL, "
			SQLSettings_InventoryControl = SQLSettings_InventoryControl & " [InventoryAPIRepostOnHandURL] [varchar](255) NULL, "
			SQLSettings_InventoryControl = SQLSettings_InventoryControl & " [InventoryPostOnHandMode] [varchar](10) NULL, "
			SQLSettings_InventoryControl = SQLSettings_InventoryControl & " [InventoryPostOnHandURL] [varchar](255) NULL, "
			SQLSettings_InventoryControl = SQLSettings_InventoryControl & " [InventoryWebAppPostOnHandMode] [varchar](10) NULL, "
			SQLSettings_InventoryControl = SQLSettings_InventoryControl & " [InventoryWebAppPostOnHandURL] [varchar](255) NULL, "
			SQLSettings_InventoryControl = SQLSettings_InventoryControl & " [InventoryAPIDailyActivityReportOnOff] [int] NOT NULL DEFAULT 0, "
			SQLSettings_InventoryControl = SQLSettings_InventoryControl & " [InventoryAPIDailyActivityReportUserNos] [varchar](255) NULL, "
			SQLSettings_InventoryControl = SQLSettings_InventoryControl & " [InventoryAPIDailyActivityReportAdditionalEmails] [varchar](1000) NULL,"
			SQLSettings_InventoryControl = SQLSettings_InventoryControl & " [InventoryAPIDailyActivityReportEmailSubject] [varchar](1000) NULL, "
			SQLSettings_InventoryControl = SQLSettings_InventoryControl & " [InventoryProductChangesReportOnOff] [int] NOT NULL DEFAULT 0, "
			SQLSettings_InventoryControl = SQLSettings_InventoryControl & " [InventoryProductChangesReportEmailSubject] [varchar](255) NULL, "
			SQLSettings_InventoryControl = SQLSettings_InventoryControl & " [InventoryProductChangesReportUserNos] [varchar](1000) NULL,"
			SQLSettings_InventoryControl = SQLSettings_InventoryControl & " [InventoryProductChangesReportAdditionalEmails] [varchar](1000) NULL "
			SQLSettings_InventoryControl = SQLSettings_InventoryControl & ") ON [PRIMARY]"
						
			Set rsSettings_InventoryControl = cnnSettings_InventoryControl.Execute(SQLSettings_InventoryControl)
			
		End If
	End If
	
	SQLSettings_InventoryControl = "SELECT COL_LENGTH('Settings_InventoryControl', 'Schedule_DailyInventoryAPIActivityByPartnerReportGeneration') AS IsItThere"
	Set rsSettings_InventoryControl = cnnSettings_InventoryControl.Execute(SQLSettings_InventoryControl)
	If IsNull(rsSettings_InventoryControl("IsItThere")) Then
		SQLSettings_InventoryControl = "ALTER TABLE Settings_InventoryControl ADD Schedule_DailyInventoryAPIActivityByPartnerReportGeneration varchar(1000) NULL"
		Set rsSettings_InventoryControl = cnnSettings_InventoryControl.Execute(SQLSettings_InventoryControl)
		SQLSettings_InventoryControl = "UPDATE Settings_InventoryControl SET Schedule_DailyInventoryAPIActivityByPartnerReportGeneration = '0,0,0,0,0,0,0,6:00 AM,6:00 AM,6:00 AM,6:00 AM,6:00 AM,6:00 AM,6:00 AM,0,0'"
		Set rsSettings_InventoryControl = cnnSettings_InventoryControl.Execute(SQLSettings_InventoryControl)
	End If


	SQLSettings_InventoryControl = "SELECT COL_LENGTH('Settings_InventoryControl', 'Schedule_InventoryProductChangesReportGeneration') AS IsItThere"
	Set rsSettings_InventoryControl = cnnSettings_InventoryControl.Execute(SQLSettings_InventoryControl)
	If IsNull(rsSettings_InventoryControl("IsItThere")) Then
		SQLSettings_InventoryControl = "ALTER TABLE Settings_InventoryControl ADD Schedule_InventoryProductChangesReportGeneration varchar(1000) NULL"
		Set rsSettings_InventoryControl = cnnSettings_InventoryControl.Execute(SQLSettings_InventoryControl)
		SQLSettings_InventoryControl = "UPDATE Settings_InventoryControl SET Schedule_InventoryProductChangesReportGeneration = '0,0,0,0,0,0,0,6:00 AM,6:00 AM,6:00 AM,6:00 AM,6:00 AM,6:00 AM,6:00 AM,0,0'"
		Set rsSettings_InventoryControl = cnnSettings_InventoryControl.Execute(SQLSettings_InventoryControl)
	End If


	' This code makes sure scheduled process information is not NULL
	SQLSettings_InventoryControl = "SELECT * FROM Settings_InventoryControl"
	Set rsSettings_InventoryControl = cnnSettings_InventoryControl.Execute(SQLSettings_InventoryControl)
	If NOT rsSettings_InventoryControl.EOF Then
		If IsNull(rsSettings_InventoryControl("Schedule_DailyInventoryAPIActivityByPartnerReportGeneration")) Then
			SQLSettings_InventoryControl = "UPDATE Settings_InventoryControl SET Schedule_DailyInventoryAPIActivityByPartnerReportGeneration = '0,0,0,0,0,0,0,6:00 AM,6:00 AM,6:00 AM,6:00 AM,6:00 AM,6:00 AM,6:00 AM,0,0'"
			Set rsSettings_InventoryControl = cnnSettings_InventoryControl.Execute(SQLSettings_InventoryControl)
		End If
	End If
	SQLSettings_InventoryControl = "SELECT * FROM Settings_InventoryControl"
	Set rsSettings_InventoryControl = cnnSettings_InventoryControl.Execute(SQLSettings_InventoryControl)
	If NOT rsSettings_InventoryControl.EOF Then
		If IsNull(rsSettings_InventoryControl("Schedule_InventoryProductChangesReportGeneration")) Then
			SQLSettings_InventoryControl = "UPDATE Settings_InventoryControl SET Schedule_InventoryProductChangesReportGeneration = '0,0,0,0,0,0,0,6:00 AM,6:00 AM,6:00 AM,6:00 AM,6:00 AM,6:00 AM,6:00 AM,0,0'"
			Set rsSettings_InventoryControl = cnnSettings_InventoryControl.Execute(SQLSettings_InventoryControl)
		End If
	End If

	
	set rsSettings_InventoryControl = nothing
	cnnSettings_InventoryControl.close
	set cnnSettings_InventoryControl = nothing
%>