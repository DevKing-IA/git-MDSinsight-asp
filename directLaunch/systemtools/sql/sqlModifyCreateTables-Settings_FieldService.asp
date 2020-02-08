<%	
	Set cnnSettingsFieldService = Server.CreateObject("ADODB.Connection")
	cnnSettingsFieldService.open (Session("ClientCnnString"))
	Set rsSettingsFieldService = Server.CreateObject("ADODB.Recordset")
	rsSettingsFieldService.CursorLocation = 3 


	Err.Clear
	on error resume next
	Set rsSettingsFieldService = cnnSettingsFieldService.Execute("SELECT * FROM Settings_FieldService")
	If Err.Description <> "" Then
		If InStr(Ucase(Err.Description),"INVALID OBJECT NAME") <> 0 Then
			On error goto 0		
			'The table is not there, we need to create it
			SQLSettings_FieldService = "CREATE TABLE [Settings_FieldService]( "
			SQLSettings_FieldService = SQLSettings_FieldService & " [ServiceTicketCarryoverReportOnOff] [int] NULL, "
			SQLSettings_FieldService = SQLSettings_FieldService & " [ServiceTicketCarryoverReportToPrimarySalesman] [int] NULL, "
			SQLSettings_FieldService = SQLSettings_FieldService & " [ServiceTicketCarryoverReportToSecondarySalesman] [int] NULL, "
			SQLSettings_FieldService = SQLSettings_FieldService & " [ServiceTicketCarryoverReportEmailSubject] [varchar](255) NULL, "
			SQLSettings_FieldService = SQLSettings_FieldService & " [ServiceTicketCarryoverReportUserNos] [varchar](1000) NULL,"
			SQLSettings_FieldService = SQLSettings_FieldService & " [ServiceTicketCarryoverReportAdditionalEmails] [varchar](1000) NULL, "
			SQLSettings_FieldService = SQLSettings_FieldService & " [FieldServiceNotesReportOnOff] [int] NULL, "
			SQLSettings_FieldService = SQLSettings_FieldService & " [FieldServiceNotesReportEmailSubject] [varchar](255) NULL, "
			SQLSettings_FieldService = SQLSettings_FieldService & " [FieldServiceNotesReportUserNos] [varchar](1000) NULL,"
			SQLSettings_FieldService = SQLSettings_FieldService & " [FieldServiceNotesReportAdditionalEmails] [varchar](1000) NULL, "
			SQLSettings_FieldService = SQLSettings_FieldService & " [ServiceTicketCarryoverReportTextSummaryOnOff] [int] NULL, "
			SQLSettings_FieldService = SQLSettings_FieldService & " [ServiceTicketCarryoverReportTextSummaryUserNos] [varchar](1000) NULL, "
			SQLSettings_FieldService = SQLSettings_FieldService & " [ServiceTicketCarryoverReportTeamIntRecIDs] [varchar](1000) NULL, "
			SQLSettings_FieldService = SQLSettings_FieldService & " [ServiceTicketCarryoverReportIncludeRegions] [int] NULL, "


			SQLSettings_FieldService = SQLSettings_FieldService & ") ON [PRIMARY]"
						
			Set rsSettingsFieldService = cnnSettingsFieldService.Execute(SQLSettings_FieldService)
			
			SQLSettings_FieldService = "INSERT INTO Settings_FieldService (ServiceTicketCarryoverReportOnOff,ServiceTicketCarryoverReportToPrimarySalesman,ServiceTicketCarryoverReportToSecondarySalesman,FieldServiceNotesReportOnOff,ServiceTicketCarryoverReportIncludeRegions) "
			SQLSettings_FieldService = SQLSettings_FieldService & " VALUES "
			SQLSettings_FieldService = SQLSettings_FieldService & " (0,1,0,0,0)"
			Set rsSettingsFieldService = cnnSettingsFieldService.Execute(SQLSettings_FieldService)
			
		End If
	End If

	On error goto 0
	
	' See if there are any records in the table & if not, insert one with default values
	Set rsSettingsFieldService = cnnSettingsFieldService.Execute("SELECT COUNT(*) AS Settings_FieldService_Count FROM Settings_FieldService")
	If rsSettingsFieldService("Settings_FieldService_Count") <> 1 Then
		SQLSettings_FieldService = "INSERT INTO Settings_FieldService (ServiceTicketCarryoverReportOnOff,ServiceTicketCarryoverReportToPrimarySalesman,ServiceTicketCarryoverReportToSecondarySalesman,FieldServiceNotesReportOnOff) "
		SQLSettings_FieldService = SQLSettings_FieldService & " VALUES "
		SQLSettings_FieldService = SQLSettings_FieldService & " (0,1,0,0)"
		Set rsSettingsFieldService = cnnSettingsFieldService.Execute(SQLSettings_FieldService)
	End If

	
	SQLSettings_FieldService = "SELECT COL_LENGTH('Settings_FieldService', 'ServiceDayStartTime') AS IsItThere"
	Set rsSettingsFieldService = cnnSettingsFieldService.Execute(SQLSettings_FieldService)
	If IsNull(rsSettingsFieldService("IsItThere")) Then
		SQLSettings_FieldService = "ALTER TABLE Settings_FieldService ADD ServiceDayStartTime varchar(50)"
		Set rsSettingsFieldService = cnnSettingsFieldService.Execute(SQLSettings_FieldService)
	End If
	
	SQLSettings_FieldService = "SELECT COL_LENGTH('Settings_FieldService', 'ServiceDayEndTime') AS IsItThere"
	Set rsSettingsFieldService = cnnSettingsFieldService.Execute(SQLSettings_FieldService)
	If IsNull(rsSettingsFieldService("IsItThere")) Then
		SQLSettings_FieldService = "ALTER TABLE Settings_FieldService ADD ServiceDayEndTime varchar(50)"
		Set rsSettingsFieldService = cnnSettingsFieldService.Execute(SQLSettings_FieldService)
	End If
	
	SQLSettings_FieldService = "SELECT COL_LENGTH('Settings_FieldService', 'ServiceDayElapsedTimeCalculationMethod') AS IsItThere"
	Set rsSettingsFieldService = cnnSettingsFieldService.Execute(SQLSettings_FieldService)
	If IsNull(rsSettingsFieldService("IsItThere")) Then
		SQLSettings_FieldService = "ALTER TABLE Settings_FieldService ADD ServiceDayElapsedTimeCalculationMethod varchar(50)"
		Set rsSettingsFieldService = cnnSettingsFieldService.Execute(SQLSettings_FieldService)
	End If
	
	SQLSettingsFieldService = "SELECT COL_LENGTH('Settings_FieldService', 'FieldServiceNotesReportEmailSubject') AS IsItThere"
	Set rsSettingsFieldService = cnnSettingsFieldService.Execute(SQLSettingsFieldService)
	If IsNull(rsSettingsFieldService("IsItThere")) Then
		SQLSettingsFieldService = "ALTER TABLE Settings_FieldService ADD FieldServiceNotesReportEmailSubject varchar(1000)"
		Set rsSettingsFieldService = cnnSettingsFieldService.Execute(SQLSettingsFieldService)
	End If
	
	SQLSettingsFieldService = "SELECT COL_LENGTH('Settings_FieldService', 'FieldServiceNotesReportUserNos') AS IsItThere"
	Set rsSettingsFieldService = cnnSettingsFieldService.Execute(SQLSettingsFieldService)
	If IsNull(rsSettingsFieldService("IsItThere")) Then
		SQLSettingsFieldService = "ALTER TABLE Settings_FieldService ADD FieldServiceNotesReportUserNos varchar(1000)"
		Set rsSettingsFieldService = cnnSettingsFieldService.Execute(SQLSettingsFieldService)
	End If
	
	SQLSettingsFieldService = "SELECT COL_LENGTH('Settings_FieldService', 'FieldServiceNotesReportAdditionalEmails') AS IsItThere"
	Set rsSettingsFieldService = cnnSettingsFieldService.Execute(SQLSettingsFieldService)
	If IsNull(rsSettingsFieldService("IsItThere")) Then
		SQLSettingsFieldService = "ALTER TABLE Settings_FieldService ADD FieldServiceNotesReportAdditionalEmails varchar(1000)"
		Set rsSettingsFieldService = cnnSettingsFieldService.Execute(SQLSettingsFieldService)
	End If
	
	SQLSettingsFieldService = "SELECT COL_LENGTH('Settings_FieldService', 'FieldServiceNotesReportOnOff') AS IsItThere"
	Set rsSettingsFieldService = cnnSettingsFieldService.Execute(SQLSettingsFieldService)
	If IsNull(rsSettingsFieldService("IsItThere")) Then
		SQLSettingsFieldService = "ALTER TABLE Settings_FieldService ADD FieldServiceNotesReportOnOff Int NULL"
		Set rsSettingsFieldService = cnnSettingsFieldService.Execute(SQLSettingsFieldService)
		SQLSettingsFieldService = "UPDATE Settings_FieldService SET FieldServiceNotesReportOnOff = 0"
		Set rsSettingsFieldService = cnnSettingsFieldService.Execute(SQLSettingsFieldService)
	End If


	SQLSettingsFieldService = "SELECT COL_LENGTH('Settings_FieldService', 'ServiceTicketCarryoverReportTeamIntRecIDs') AS IsItThere"
	Set rsSettingsFieldService = cnnSettingsFieldService.Execute(SQLSettingsFieldService)
	If IsNull(rsSettingsFieldService("IsItThere")) Then
		SQLSettingsFieldService = "ALTER TABLE Settings_FieldService ADD ServiceTicketCarryoverReportTeamIntRecIDs varchar(1000)"
		Set rsSettingsFieldService = cnnSettingsFieldService.Execute(SQLSettingsFieldService)
	End If
	

	SQLSettingsFieldService = "SELECT COL_LENGTH('Settings_FieldService', 'ServiceTicketCarryoverReportIncludeRegions') AS IsItThere"
	Set rsSettingsFieldService = cnnSettingsFieldService.Execute(SQLSettingsFieldService)
	If IsNull(rsSettingsFieldService("IsItThere")) Then
		SQLSettingsFieldService = "ALTER TABLE Settings_FieldService ADD ServiceTicketCarryoverReportIncludeRegions Int NULL"
		Set rsSettingsFieldService = cnnSettingsFieldService.Execute(SQLSettingsFieldService)
		SQLSettingsFieldService = "UPDATE Settings_FieldService SET ServiceTicketCarryoverReportIncludeRegions = 0"
		Set rsSettingsFieldService = cnnSettingsFieldService.Execute(SQLSettingsFieldService)
	End If



	' This one is a DROP
	SQLSettings_FieldService = "SELECT COL_LENGTH('Settings_FieldService', 'PMCAllDays') AS IsItThere"
	Set rsSettingsFieldService = cnnSettingsFieldService.Execute(SQLSettings_FieldService)
	If NOT IsNull(rsSettingsFieldService("IsItThere")) Then
		SQLSettings_FieldService = "ALTER TABLE Settings_FieldService DROP COLUMN PMCAllDays"
		Set rsSettingsFieldService = cnnSettingsFieldService.Execute(SQLSettings_FieldService)
	End If

	' This one is a DROP
	SQLSettings_FieldService = "SELECT COL_LENGTH('Settings_FieldService', 'PMCAllDaysFieldService') AS IsItThere"
	Set rsSettingsFieldService = cnnSettingsFieldService.Execute(SQLSettings_FieldService)
	If NOT IsNull(rsSettingsFieldService("IsItThere")) Then
		SQLSettings_FieldService = "ALTER TABLE Settings_FieldService DROP COLUMN PMCAllDaysFieldService"
		Set rsSettingsFieldService = cnnSettingsFieldService.Execute(SQLSettings_FieldService)
	End If
	

	SQLSettingsFieldService = "SELECT COL_LENGTH('Settings_FieldService', 'AutoDispatchUsersOnOff') AS IsItThere"
	Set rsSettingsFieldService = cnnSettingsFieldService.Execute(SQLSettingsFieldService)
	If IsNull(rsSettingsFieldService("IsItThere")) Then
		SQLSettingsFieldService = "ALTER TABLE Settings_FieldService ADD AutoDispatchUsersOnOff Int NULL"
		Set rsSettingsFieldService = cnnSettingsFieldService.Execute(SQLSettingsFieldService)
		SQLSettingsFieldService = "UPDATE Settings_FieldService SET AutoDispatchUsersOnOff = 0"
		Set rsSettingsFieldService = cnnSettingsFieldService.Execute(SQLSettingsFieldService)
	End If

	SQLSettingsFieldService = "SELECT COL_LENGTH('Settings_FieldService', 'AutoDispatchUserNos') AS IsItThere"
	Set rsSettingsFieldService = cnnSettingsFieldService.Execute(SQLSettingsFieldService)
	If IsNull(rsSettingsFieldService("IsItThere")) Then
		SQLSettingsFieldService = "ALTER TABLE Settings_FieldService ADD AutoDispatchUserNos varchar(255) NULL"
		Set rsSettingsFieldService = cnnSettingsFieldService.Execute(SQLSettingsFieldService)
	End If
	
	SQLSettingsFieldService = "SELECT COL_LENGTH('Settings_FieldService', 'FilterChangeDays') AS IsItThere"
	Set rsSettingsFieldService = cnnSettingsFieldService.Execute(SQLSettingsFieldService)
	If IsNull(rsSettingsFieldService("IsItThere")) Then
		SQLSettingsFieldService = "ALTER TABLE Settings_FieldService ADD FilterChangeDays int NULL"
		Set rsSettingsFieldService = cnnSettingsFieldService.Execute(SQLSettingsFieldService)
		SQLSettingsFieldService = "UPDATE Settings_FieldService SET FilterChangeDays = 30"
		Set rsSettingsFieldService = cnnSettingsFieldService.Execute(SQLSettingsFieldService)
	End If

	SQLSettingsFieldService = "SELECT COL_LENGTH('Settings_FieldService', 'FilterChangeDaysFieldService') AS IsItThere"
	Set rsSettingsFieldService = cnnSettingsFieldService.Execute(SQLSettingsFieldService)
	If IsNull(rsSettingsFieldService("IsItThere")) Then
		SQLSettingsFieldService = "ALTER TABLE Settings_FieldService ADD FilterChangeDaysFieldService int NULL"
		Set rsSettingsFieldService = cnnSettingsFieldService.Execute(SQLSettingsFieldService)
		SQLSettingsFieldService = "UPDATE Settings_FieldService SET FilterChangeDaysFieldService = 30"
		Set rsSettingsFieldService = cnnSettingsFieldService.Execute(SQLSettingsFieldService)
	End If


	SQLSettingsFieldService = "SELECT COL_LENGTH('Settings_FieldService', 'CarryoverReportInclCustType') AS IsItThere"
	Set rsSettingsFieldService = cnnSettingsFieldService.Execute(SQLSettingsFieldService)
	If IsNull(rsSettingsFieldService("IsItThere")) Then
		SQLSettingsFieldService = "ALTER TABLE Settings_FieldService ADD CarryoverReportInclCustType int NULL"
		Set rsSettingsFieldService = cnnSettingsFieldService.Execute(SQLSettingsFieldService)
		SQLSettingsFieldService = "UPDATE Settings_FieldService SET CarryoverReportInclCustType = 1"
		Set rsSettingsFieldService = cnnSettingsFieldService.Execute(SQLSettingsFieldService)
	End If


	SQLSettingsFieldService = "SELECT COL_LENGTH('Settings_FieldService', 'CarryoverReportInclTicketNum') AS IsItThere"
	Set rsSettingsFieldService = cnnSettingsFieldService.Execute(SQLSettingsFieldService)
	If IsNull(rsSettingsFieldService("IsItThere")) Then
		SQLSettingsFieldService = "ALTER TABLE Settings_FieldService ADD CarryoverReportInclTicketNum int NULL"
		Set rsSettingsFieldService = cnnSettingsFieldService.Execute(SQLSettingsFieldService)
		SQLSettingsFieldService = "UPDATE Settings_FieldService SET CarryoverReportInclTicketNum = 1"
		Set rsSettingsFieldService = cnnSettingsFieldService.Execute(SQLSettingsFieldService)
	End If


	SQLSettingsFieldService = "SELECT COL_LENGTH('Settings_FieldService', 'CarryoverReportShowRedoBreakdown') AS IsItThere"
	Set rsSettingsFieldService = cnnSettingsFieldService.Execute(SQLSettingsFieldService)
	If IsNull(rsSettingsFieldService("IsItThere")) Then
		SQLSettingsFieldService = "ALTER TABLE Settings_FieldService ADD CarryoverReportShowRedoBreakdown int NULL"
		Set rsSettingsFieldService = cnnSettingsFieldService.Execute(SQLSettingsFieldService)
		SQLSettingsFieldService = "UPDATE Settings_FieldService SET CarryoverReportShowRedoBreakdown = 1"
		Set rsSettingsFieldService = cnnSettingsFieldService.Execute(SQLSettingsFieldService)
	End If


	SQLSettingsFieldService = "SELECT COL_LENGTH('Settings_FieldService', 'FilterChangeIndicatorAndButtonColor') AS IsItThere"
	Set rsSettingsFieldService = cnnSettingsFieldService.Execute(SQLSettingsFieldService)
	If IsNull(rsSettingsFieldService("IsItThere")) Then
		SQLSettingsFieldService = "ALTER TABLE Settings_FieldService ADD FilterChangeIndicatorAndButtonColor varchar(255) NULL"
		Set rsSettingsFieldService = cnnSettingsFieldService.Execute(SQLSettingsFieldService)
		SQLSettingsFieldService = "UPDATE Settings_FieldService SET FilterChangeIndicatorAndButtonColor = '#dddd53'"
		Set rsSettingsFieldService = cnnSettingsFieldService.Execute(SQLSettingsFieldService)
	End If
	
	
	SQLSettingsFieldService = "SELECT COL_LENGTH('Settings_FieldService', 'ShowSeparateFilterChangesTabOnServiceScreen') AS IsItThere"
	Set rsSettingsFieldService = cnnSettingsFieldService.Execute(SQLSettingsFieldService)
	If IsNull(rsSettingsFieldService("IsItThere")) Then
		SQLSettingsFieldService = "ALTER TABLE Settings_FieldService ADD ShowSeparateFilterChangesTabOnServiceScreen Int NULL"
		Set rsSettingsFieldService = cnnSettingsFieldService.Execute(SQLSettingsFieldService)
		SQLSettingsFieldService = "UPDATE Settings_FieldService SET ShowSeparateFilterChangesTabOnServiceScreen = 0"
		Set rsSettingsFieldService = cnnSettingsFieldService.Execute(SQLSettingsFieldService)
	End If
	
	SQLSettingsFieldService = "SELECT COL_LENGTH('Settings_FieldService', 'AutoFilterChangeGenerationONOFF') AS IsItThere"
	Set rsSettingsFieldService = cnnSettingsFieldService.Execute(SQLSettingsFieldService)
	If IsNull(rsSettingsFieldService("IsItThere")) Then
		SQLSettingsFieldService = "ALTER TABLE Settings_FieldService ADD AutoFilterChangeGenerationONOFF Int NULL"
		Set rsSettingsFieldService = cnnSettingsFieldService.Execute(SQLSettingsFieldService)
		SQLSettingsFieldService = "UPDATE Settings_FieldService SET AutoFilterChangeGenerationONOFF = 0"
		Set rsSettingsFieldService = cnnSettingsFieldService.Execute(SQLSettingsFieldService)
	End If

	SQLSettingsFieldService = "SELECT COL_LENGTH('Settings_FieldService', 'ShowPartsButton') AS IsItThere"
	Set rsSettingsFieldService = cnnSettingsFieldService.Execute(SQLSettingsFieldService)
	If IsNull(rsSettingsFieldService("IsItThere")) Then
		SQLSettingsFieldService = "ALTER TABLE Settings_FieldService ADD ShowPartsButton Int NULL"
		Set rsSettingsFieldService = cnnSettingsFieldService.Execute(SQLSettingsFieldService)
		SQLSettingsFieldService = "UPDATE Settings_FieldService SET ShowPartsButton = 0"
		Set rsSettingsFieldService = cnnSettingsFieldService.Execute(SQLSettingsFieldService)
	End If
	
	
	SQLSettingsFieldService = "SELECT COL_LENGTH('Settings_FieldService', 'AutoFilterChangeUseRegions') AS IsItThere"
	Set rsSettingsFieldService = cnnSettingsFieldService.Execute(SQLSettingsFieldService)
	If IsNull(rsSettingsFieldService("IsItThere")) Then
		SQLSettingsFieldService = "ALTER TABLE Settings_FieldService ADD AutoFilterChangeUseRegions Int NULL"
		Set rsSettingsFieldService = cnnSettingsFieldService.Execute(SQLSettingsFieldService)
		SQLSettingsFieldService = "UPDATE Settings_FieldService SET AutoFilterChangeUseRegions = 0"
		Set rsSettingsFieldService = cnnSettingsFieldService.Execute(SQLSettingsFieldService)
	End If
	
	
	SQLSettingsFieldService = "SELECT COL_LENGTH('Settings_FieldService', 'AutoFilterChangeMaxNumTicketsPerDay') AS IsItThere"
	Set rsSettingsFieldService = cnnSettingsFieldService.Execute(SQLSettingsFieldService)
	If IsNull(rsSettingsFieldService("IsItThere")) Then
		SQLSettingsFieldService = "ALTER TABLE Settings_FieldService ADD AutoFilterChangeMaxNumTicketsPerDay Int NULL"
		Set rsSettingsFieldService = cnnSettingsFieldService.Execute(SQLSettingsFieldService)
		SQLSettingsFieldService = "UPDATE Settings_FieldService SET AutoFilterChangeMaxNumTicketsPerDay = 25"
		Set rsSettingsFieldService = cnnSettingsFieldService.Execute(SQLSettingsFieldService)
	End If

	SQLSettingsFieldService = "SELECT COL_LENGTH('Settings_FieldService', 'Schedule_FilterGeneration') AS IsItThere"
	Set rsSettingsFieldService = cnnSettingsFieldService.Execute(SQLSettingsFieldService)
	If IsNull(rsSettingsFieldService("IsItThere")) Then
		SQLSettingsFieldService = "ALTER TABLE Settings_FieldService ADD Schedule_FilterGeneration varchar(1000) NULL"
		Set rsSettingsFieldService = cnnSettingsFieldService.Execute(SQLSettingsFieldService)
		SQLSettingsFieldService = "UPDATE Settings_FieldService SET Schedule_FilterGeneration = '0,0,0,0,0,0,0,12:00 AM,12:00 AM,12:00 AM,12:00 AM,12:00 AM,12:00 AM,12:00 AM,0,0'"
		Set rsSettingsFieldService = cnnSettingsFieldService.Execute(SQLSettingsFieldService)
	End If
	
	SQLSettingsFieldService = "SELECT COL_LENGTH('Settings_FieldService', 'Schedule_FieldServiceNotesReportGeneration') AS IsItThere"
	Set rsSettingsFieldService = cnnSettingsFieldService.Execute(SQLSettingsFieldService)
	If IsNull(rsSettingsFieldService("IsItThere")) Then
		SQLSettingsFieldService = "ALTER TABLE Settings_FieldService ADD Schedule_FieldServiceNotesReportGeneration varchar(1000) NULL"
		Set rsSettingsFieldService = cnnSettingsFieldService.Execute(SQLSettingsFieldService)
		SQLSettingsFieldService = "UPDATE Settings_FieldService SET Schedule_FieldServiceNotesReportGeneration = '0,0,0,0,0,0,0,8:00 AM,8:00 AM,8:00 AM,8:00 AM,8:00 AM,8:00 AM,8:00 AM,0,0'"
		Set rsSettingsFieldService = cnnSettingsFieldService.Execute(SQLSettingsFieldService)
	End If

	SQLSettingsFieldService = "SELECT COL_LENGTH('Settings_FieldService', 'Schedule_ServiceTicketCarryoverReportGeneration') AS IsItThere"
	Set rsSettingsFieldService = cnnSettingsFieldService.Execute(SQLSettingsFieldService)
	If IsNull(rsSettingsFieldService("IsItThere")) Then
		SQLSettingsFieldService = "ALTER TABLE Settings_FieldService ADD Schedule_ServiceTicketCarryoverReportGeneration varchar(1000) NULL"
		Set rsSettingsFieldService = cnnSettingsFieldService.Execute(SQLSettingsFieldService)
		SQLSettingsFieldService = "UPDATE Settings_FieldService SET Schedule_ServiceTicketCarryoverReportGeneration = '0,1,1,1,1,1,0,6:00 PM,6:00 PM,6:00 PM,6:00 PM,6:00 PM,6:00 PM,6:00 PM,0,0'"
		Set rsSettingsFieldService = cnnSettingsFieldService.Execute(SQLSettingsFieldService)
	End If

	SQLSettingsFieldService = "SELECT COL_LENGTH('Settings_FieldService', 'ServiceTicketthresholdReportONOFF') AS IsItThere"
	Set rsSettingsFieldService = cnnSettingsFieldService.Execute(SQLSettingsFieldService)
	If IsNull(rsSettingsFieldService("IsItThere")) Then
		SQLSettingsFieldService = "ALTER TABLE Settings_FieldService ADD ServiceTicketthresholdReportONOFF Int NULL"
		Set rsSettingsFieldService = cnnSettingsFieldService.Execute(SQLSettingsFieldService)
		SQLSettingsFieldService = "UPDATE Settings_FieldService SET ServiceTicketthresholdReportONOFF = 0"
		Set rsSettingsFieldService = cnnSettingsFieldService.Execute(SQLSettingsFieldService)
	End If

	SQLSettingsFieldService = "SELECT COL_LENGTH('Settings_FieldService', 'ServiceTicketthresholdReportOnlyUndispatched') AS IsItThere"
	Set rsSettingsFieldService = cnnSettingsFieldService.Execute(SQLSettingsFieldService)
	If IsNull(rsSettingsFieldService("IsItThere")) Then
		SQLSettingsFieldService = "ALTER TABLE Settings_FieldService ADD ServiceTicketthresholdReportOnlyUndispatched Int NULL"
		Set rsSettingsFieldService = cnnSettingsFieldService.Execute(SQLSettingsFieldService)
		SQLSettingsFieldService = "UPDATE Settings_FieldService SET ServiceTicketthresholdReportOnlyUndispatched = 1"
		Set rsSettingsFieldService = cnnSettingsFieldService.Execute(SQLSettingsFieldService)
	End If

	SQLSettingsFieldService = "SELECT COL_LENGTH('Settings_FieldService', 'ServiceTicketthresholdReportOnlySkipFilterChanges') AS IsItThere"
	Set rsSettingsFieldService = cnnSettingsFieldService.Execute(SQLSettingsFieldService)
	If IsNull(rsSettingsFieldService("IsItThere")) Then
		SQLSettingsFieldService = "ALTER TABLE Settings_FieldService ADD ServiceTicketthresholdReportOnlySkipFilterChanges Int NULL"
		Set rsSettingsFieldService = cnnSettingsFieldService.Execute(SQLSettingsFieldService)
		SQLSettingsFieldService = "UPDATE Settings_FieldService SET ServiceTicketthresholdReportOnlySkipFilterChanges = 0"
		Set rsSettingsFieldService = cnnSettingsFieldService.Execute(SQLSettingsFieldService)
	End If

	SQLSettingsFieldService = "SELECT COL_LENGTH('Settings_FieldService', 'ServiceTicketthresholdReportthresholdHours') AS IsItThere"
	Set rsSettingsFieldService = cnnSettingsFieldService.Execute(SQLSettingsFieldService)
	If IsNull(rsSettingsFieldService("IsItThere")) Then
		SQLSettingsFieldService = "ALTER TABLE Settings_FieldService ADD ServiceTicketthresholdReportthresholdHours Int NULL"
		Set rsSettingsFieldService = cnnSettingsFieldService.Execute(SQLSettingsFieldService)
		SQLSettingsFieldService = "UPDATE Settings_FieldService SET ServiceTicketthresholdReportthresholdHours = 16"
		Set rsSettingsFieldService = cnnSettingsFieldService.Execute(SQLSettingsFieldService)
	End If

	SQLSettingsFieldService = "SELECT COL_LENGTH('Settings_FieldService', 'ServiceTicketthresholdReportUserNos') AS IsItThere"
	Set rsSettingsFieldService = cnnSettingsFieldService.Execute(SQLSettingsFieldService)
	If IsNull(rsSettingsFieldService("IsItThere")) Then
		SQLSettingsFieldService = "ALTER TABLE Settings_FieldService ADD ServiceTicketthresholdReportUserNos varchar(1000)"
		Set rsSettingsFieldService = cnnSettingsFieldService.Execute(SQLSettingsFieldService)
	End If
	
	SQLSettingsFieldService = "SELECT COL_LENGTH('Settings_FieldService', 'ServiceTicketthresholdReportAdditionalEmails') AS IsItThere"
	Set rsSettingsFieldService = cnnSettingsFieldService.Execute(SQLSettingsFieldService)
	If IsNull(rsSettingsFieldService("IsItThere")) Then
		SQLSettingsFieldService = "ALTER TABLE Settings_FieldService ADD ServiceTicketthresholdReportAdditionalEmails varchar(1000)"
		Set rsSettingsFieldService = cnnSettingsFieldService.Execute(SQLSettingsFieldService)
	End If

	SQLSettingsFieldService = "SELECT COL_LENGTH('Settings_FieldService', 'Schedule_ServiceTicketThreshholdReportGeneration') AS IsItThere"
	Set rsSettingsFieldService = cnnSettingsFieldService.Execute(SQLSettingsFieldService)
	If IsNull(rsSettingsFieldService("IsItThere")) Then
		SQLSettingsFieldService = "ALTER TABLE Settings_FieldService ADD Schedule_ServiceTicketThreshholdReportGeneration varchar(1000) NULL"
		Set rsSettingsFieldService = cnnSettingsFieldService.Execute(SQLSettingsFieldService)
		SQLSettingsFieldService = "UPDATE Settings_FieldService SET Schedule_ServiceTicketThreshholdReportGeneration = '0,1,1,1,1,1,0,8:30 AM,8:30 AM,8:30 AM,8:30 AM,8:30 AM,8:30 AM,8:30 AM,0,0'"
		Set rsSettingsFieldService = cnnSettingsFieldService.Execute(SQLSettingsFieldService)
	End If
	
	
	SQLSettingsFieldService = "SELECT COL_LENGTH('Settings_FieldService', 'ServiceTicketCarryoverReportTextSummaryOnOff') AS IsItThere"
	Set rsSettingsFieldService = cnnSettingsFieldService.Execute(SQLSettingsFieldService)
	If IsNull(rsSettingsFieldService("IsItThere")) Then
		SQLSettingsFieldService = "ALTER TABLE Settings_FieldService ADD ServiceTicketCarryoverReportTextSummaryOnOff Int NULL"
		Set rsSettingsFieldService = cnnSettingsFieldService.Execute(SQLSettingsFieldService)
		SQLSettingsFieldService = "UPDATE Settings_FieldService SET ServiceTicketCarryoverReportTextSummaryOnOff = 0"
		Set rsSettingsFieldService = cnnSettingsFieldService.Execute(SQLSettingsFieldService)
	End If
	
	SQLSettingsFieldService = "SELECT COL_LENGTH('Settings_FieldService', 'ServiceTicketCarryoverReportTextSummaryUserNos') AS IsItThere"
	Set rsSettingsFieldService = cnnSettingsFieldService.Execute(SQLSettingsFieldService)
	If IsNull(rsSettingsFieldService("IsItThere")) Then
		SQLSettingsFieldService = "ALTER TABLE Settings_FieldService ADD ServiceTicketCarryoverReportTextSummaryUserNos varchar(1000)"
		Set rsSettingsFieldService = cnnSettingsFieldService.Execute(SQLSettingsFieldService)
	End If



	' INCORRECT SPELLING
		
	' This one is a DROP
	SQLSettings_FieldService = "SELECT COL_LENGTH('Settings_FieldService', 'ServiceTicketthreshholdReportONOFF') AS IsItThere"
	Set rsSettingsFieldService = cnnSettingsFieldService.Execute(SQLSettings_FieldService)
	If NOT IsNull(rsSettingsFieldService("IsItThere")) Then
		SQLSettings_FieldService = "ALTER TABLE Settings_FieldService DROP COLUMN ServiceTicketthreshholdReportONOFF"
		Set rsSettingsFieldService = cnnSettingsFieldService.Execute(SQLSettings_FieldService)
	End If
		
	' This one is a DROP
	SQLSettings_FieldService = "SELECT COL_LENGTH('Settings_FieldService', 'ServiceTicketthreshholdReportOnlyUndispatched') AS IsItThere"
	Set rsSettingsFieldService = cnnSettingsFieldService.Execute(SQLSettings_FieldService)
	If NOT IsNull(rsSettingsFieldService("IsItThere")) Then
		SQLSettings_FieldService = "ALTER TABLE Settings_FieldService DROP COLUMN ServiceTicketthreshholdReportOnlyUndispatched"
		Set rsSettingsFieldService = cnnSettingsFieldService.Execute(SQLSettings_FieldService)
	End If
		
	' This one is a DROP
	SQLSettings_FieldService = "SELECT COL_LENGTH('Settings_FieldService', 'ServiceTicketthreshholdReportOnlySkipFilterChanges') AS IsItThere"
	Set rsSettingsFieldService = cnnSettingsFieldService.Execute(SQLSettings_FieldService)
	If NOT IsNull(rsSettingsFieldService("IsItThere")) Then
		SQLSettings_FieldService = "ALTER TABLE Settings_FieldService DROP COLUMN ServiceTicketthreshholdReportOnlySkipFilterChanges"
		Set rsSettingsFieldService = cnnSettingsFieldService.Execute(SQLSettings_FieldService)
	End If
		
	' This one is a DROP
	SQLSettings_FieldService = "SELECT COL_LENGTH('Settings_FieldService', 'ServiceTicketthreshholdReportthresholdHours') AS IsItThere"
	Set rsSettingsFieldService = cnnSettingsFieldService.Execute(SQLSettings_FieldService)
	If NOT IsNull(rsSettingsFieldService("IsItThere")) Then
		SQLSettings_FieldService = "ALTER TABLE Settings_FieldService DROP COLUMN ServiceTicketthreshholdReportthresholdHours"
		Set rsSettingsFieldService = cnnSettingsFieldService.Execute(SQLSettings_FieldService)
	End If
		
	' This one is a DROP
	SQLSettings_FieldService = "SELECT COL_LENGTH('Settings_FieldService', 'ServiceTicketthreshholdReportUserNos') AS IsItThere"
	Set rsSettingsFieldService = cnnSettingsFieldService.Execute(SQLSettings_FieldService)
	If NOT IsNull(rsSettingsFieldService("IsItThere")) Then
		SQLSettings_FieldService = "ALTER TABLE Settings_FieldService DROP COLUMN ServiceTicketthreshholdReportUserNos"
		Set rsSettingsFieldService = cnnSettingsFieldService.Execute(SQLSettings_FieldService)
	End If
		
	' This one is a DROP
	SQLSettings_FieldService = "SELECT COL_LENGTH('Settings_FieldService', 'ServiceTicketthreshholdReportAdditionalEmails') AS IsItThere"
	Set rsSettingsFieldService = cnnSettingsFieldService.Execute(SQLSettings_FieldService)
	If NOT IsNull(rsSettingsFieldService("IsItThere")) Then
		SQLSettings_FieldService = "ALTER TABLE Settings_FieldService DROP COLUMN ServiceTicketthreshholdReportAdditionalEmails"
		Set rsSettingsFieldService = cnnSettingsFieldService.Execute(SQLSettings_FieldService)
	End If
		
	' This one is a DROP
	SQLSettings_FieldService = "SELECT COL_LENGTH('Settings_FieldService', 'ServiceTicketthreshholdReportGeneration') AS IsItThere"
	Set rsSettingsFieldService = cnnSettingsFieldService.Execute(SQLSettings_FieldService)
	If NOT IsNull(rsSettingsFieldService("IsItThere")) Then
		SQLSettings_FieldService = "ALTER TABLE Settings_FieldService DROP COLUMN ServiceTicketthreshholdReportGeneration"
		Set rsSettingsFieldService = cnnSettingsFieldService.Execute(SQLSettings_FieldService)
	End If
	

	SQLSettingsFieldService = "SELECT COL_LENGTH('Settings_FieldService', 'ServiceTicketScreenShowHoldTab') AS IsItThere"
	Set rsSettingsFieldService = cnnSettingsFieldService.Execute(SQLSettingsFieldService)
	If IsNull(rsSettingsFieldService("IsItThere")) Then
		SQLSettingsFieldService = "ALTER TABLE Settings_FieldService ADD ServiceTicketScreenShowHoldTab Int NULL"
		Set rsSettingsFieldService = cnnSettingsFieldService.Execute(SQLSettingsFieldService)
		SQLSettingsFieldService = "UPDATE Settings_FieldService SET ServiceTicketScreenShowHoldTab = 0"
		Set rsSettingsFieldService = cnnSettingsFieldService.Execute(SQLSettingsFieldService)
	End If


	' This code makes sure scheduled process information is not NULL
	SQLSettingsFieldService = "SELECT * FROM Settings_FieldService"
	Set rsSettingsFieldService = cnnSettingsFieldService.Execute(SQLSettingsFieldService)
	If NOT rsSettingsFieldService.EOF Then
		If IsNull(rsSettingsFieldService("Schedule_FilterGeneration")) Then
			SQLSettingsFieldService = "UPDATE Settings_FieldService SET Schedule_FilterGeneration = '0,0,0,0,0,0,0,12:00 AM,12:00 AM,12:00 AM,12:00 AM,12:00 AM,12:00 AM,12:00 AM,0,0'"
			Set rsSettingsFieldService = cnnSettingsFieldService.Execute(SQLSettingsFieldService)
		End If
	End If
	
	SQLSettingsFieldService = "SELECT * FROM Settings_FieldService"
	Set rsSettingsFieldService = cnnSettingsFieldService.Execute(SQLSettingsFieldService)
	If NOT rsSettingsFieldService.EOF Then
		If IsNull(rsSettingsFieldService("Schedule_ServiceTicketCarryoverReportGeneration")) Then
			SQLSettingsFieldService = "UPDATE Settings_FieldService SET Schedule_ServiceTicketCarryoverReportGeneration = '0,1,1,1,1,1,0,6:00 PM,6:00 PM,6:00 PM,6:00 PM,6:00 PM,6:00 PM,6:00 PM,0,0'"
			Set rsSettingsFieldService = cnnSettingsFieldService.Execute(SQLSettingsFieldService)
		End If
	End If
	
	SQLSettingsFieldService = "SELECT * FROM Settings_FieldService"
	Set rsSettingsFieldService = cnnSettingsFieldService.Execute(SQLSettingsFieldService)
	If NOT rsSettingsFieldService.EOF Then
		If IsNull(rsSettingsFieldService("Schedule_FieldServiceNotesReportGeneration")) Then
			SQLSettingsFieldService = "UPDATE Settings_FieldService SET Schedule_FieldServiceNotesReportGeneration = '0,0,0,0,0,0,0,8:00 AM,8:00 AM,8:00 AM,8:00 AM,8:00 AM,8:00 AM,8:00 AM,0,0'"
			Set rsSettingsFieldService = cnnSettingsFieldService.Execute(SQLSettingsFieldService)
		End If
	End If
	
	SQLSettingsFieldService = "SELECT * FROM Settings_FieldService"
	Set rsSettingsFieldService = cnnSettingsFieldService.Execute(SQLSettingsFieldService)
	If NOT rsSettingsFieldService.EOF Then
		If IsNull(rsSettingsFieldService("Schedule_ServiceTicketThreshholdReportGeneration")) Then
			SQLSettingsFieldService = "UPDATE Settings_FieldService SET Schedule_ServiceTicketThreshholdReportGeneration = '0,1,1,1,1,1,0,8:30 AM,8:30 AM,8:30 AM,8:30 AM,8:30 AM,8:30 AM,8:30 AM,0,0'"
			Set rsSettingsFieldService = cnnSettingsFieldService.Execute(SQLSettingsFieldService)
		End If
	End If


	SQLSettingsFieldService = "SELECT COL_LENGTH('Settings_FieldService', 'ServiceTicketScreenShowHoldTab') AS IsItThere"
	Set rsSettingsFieldService = cnnSettingsFieldService.Execute(SQLSettingsFieldService)
	If IsNull(rsSettingsFieldService("IsItThere")) Then
		SQLSettingsFieldService = "ALTER TABLE Settings_FieldService ADD ServiceTicketScreenShowHoldTab Int NULL"
		Set rsSettingsFieldService = cnnSettingsFieldService.Execute(SQLSettingsFieldService)
		SQLSettingsFieldService = "UPDATE Settings_FieldService SET ServiceTicketScreenShowHoldTab = 0"
		Set rsSettingsFieldService = cnnSettingsFieldService.Execute(SQLSettingsFieldService)
	End If
	
	
	SQLSettingsFieldService = "SELECT COL_LENGTH('Settings_FieldService', 'ServiceTicketthresholdReportUserNos') AS IsItThere"
	Set rsSettingsFieldService = cnnSettingsFieldService.Execute(SQLSettingsFieldService)
	If IsNull(rsSettingsFieldService("IsItThere")) Then
		SQLSettingsFieldService = "ALTER TABLE Settings_FieldService ADD ServiceTicketthresholdReportUserNos varchar(1000)"
		Set rsSettingsFieldService = cnnSettingsFieldService.Execute(SQLSettingsFieldService)
	End If

	'*****************************************************
	'NEW FIELD SERVICE GLOBAL COLOR SETTINGS
	'*****************************************************

	SQL_SettingsFieldService = "SELECT COL_LENGTH('Settings_FieldService', 'FSBoardKioskGlobalUseRegions') AS IsItThere"
	Set rsSettingsFieldService = cnnSettingsFieldService.Execute(SQL_SettingsFieldService)
	If IsNull(rsSettingsFieldService("IsItThere")) Then
		SQL_SettingsFieldService = "ALTER TABLE Settings_FieldService ADD FSBoardKioskGlobalUseRegions INT NULL"
		Set rsSettingsFieldService = cnnSettingsFieldService.Execute(SQL_SettingsFieldService)
		SQL_SettingsFieldService = "UPDATE Settings_FieldService SET FSBoardKioskGlobalUseRegions = 1"
		Set rsSettingsFieldService = cnnSettingsFieldService.Execute(SQL_SettingsFieldService)		
	End If

	SQL_SettingsFieldService = "SELECT COL_LENGTH('Settings_FieldService', 'FSBoardKioskGlobalTitleText') AS IsItThere"
	Set rsSettingsFieldService = cnnSettingsFieldService.Execute(SQL_SettingsFieldService)
	If IsNull(rsSettingsFieldService("IsItThere")) Then
		SQL_SettingsFieldService = "ALTER TABLE Settings_FieldService ADD FSBoardKioskGlobalTitleText varchar(100)"
		Set rsSettingsFieldService = cnnSettingsFieldService.Execute(SQL_SettingsFieldService)
		SQL_SettingsFieldService = "UPDATE Settings_FieldService SET FSBoardKioskGlobalTitleText = 'Field Service Status'"
		Set rsSettingsFieldService = cnnSettingsFieldService.Execute(SQL_SettingsFieldService)		
	End If
	
	SQL_SettingsFieldService = "SELECT COL_LENGTH('Settings_FieldService', 'FSBoardKioskGlobalTitleTextFontColor') AS IsItThere"
	Set rsSettingsFieldService = cnnSettingsFieldService.Execute(SQL_SettingsFieldService)
	If IsNull(rsSettingsFieldService("IsItThere")) Then
		SQL_SettingsFieldService = "ALTER TABLE Settings_FieldService ADD FSBoardKioskGlobalTitleTextFontColor varchar(50)"
		Set rsSettingsFieldService = cnnSettingsFieldService.Execute(SQL_SettingsFieldService)
		SQL_SettingsFieldService = "UPDATE Settings_FieldService SET FSBoardKioskGlobalTitleTextFontColor = '#000000'"
		Set rsSettingsFieldService = cnnSettingsFieldService.Execute(SQL_SettingsFieldService)		
	End If

	SQL_SettingsFieldService = "SELECT COL_LENGTH('Settings_FieldService', 'FSBoardKioskGlobalTitleGradientColor') AS IsItThere"
	Set rsSettingsFieldService = cnnSettingsFieldService.Execute(SQL_SettingsFieldService)
	If IsNull(rsSettingsFieldService("IsItThere")) Then
		SQL_SettingsFieldService = "ALTER TABLE Settings_FieldService ADD FSBoardKioskGlobalTitleGradientColor varchar(50)"
		Set rsSettingsFieldService = cnnSettingsFieldService.Execute(SQL_SettingsFieldService)
		SQL_SettingsFieldService = "UPDATE Settings_FieldService SET FSBoardKioskGlobalTitleGradientColor = '#80B8FF'"
		Set rsSettingsFieldService = cnnSettingsFieldService.Execute(SQL_SettingsFieldService)		
	End If
	
	SQL_SettingsFieldService = "SELECT COL_LENGTH('Settings_FieldService', 'FSBoardKioskGlobalColorPieTimer') AS IsItThere"
	Set rsSettingsFieldService = cnnSettingsFieldService.Execute(SQL_SettingsFieldService)
	If IsNull(rsSettingsFieldService("IsItThere")) Then
		SQL_SettingsFieldService = "ALTER TABLE Settings_FieldService ADD FSBoardKioskGlobalColorPieTimer varchar(50) NULL"
		Set rsSettingsFieldService = cnnSettingsFieldService.Execute(SQL_SettingsFieldService)
		SQL_SettingsFieldService = "UPDATE Settings_FieldService SET FSBoardKioskGlobalColorPieTimer = '#000000'"
		Set rsSettingsFieldService = cnnSettingsFieldService.Execute(SQL_SettingsFieldService)
	End If

	SQL_SettingsFieldService = "SELECT COL_LENGTH('Settings_FieldService', 'FSBoardKioskGlobalColorAwaitingDispatch') AS IsItThere"
	Set rsSettingsFieldService = cnnSettingsFieldService.Execute(SQL_SettingsFieldService)
	If IsNull(rsSettingsFieldService("IsItThere")) Then
		SQL_SettingsFieldService = "ALTER TABLE Settings_FieldService ADD FSBoardKioskGlobalColorAwaitingDispatch varchar(50) NULL"
		Set rsSettingsFieldService = cnnSettingsFieldService.Execute(SQL_SettingsFieldService)
		SQL_SettingsFieldService = "UPDATE Settings_FieldService SET FSBoardKioskGlobalColorAwaitingDispatch = '#FC9901'"
		Set rsSettingsFieldService = cnnSettingsFieldService.Execute(SQL_SettingsFieldService)
	End If

	SQL_SettingsFieldService = "SELECT COL_LENGTH('Settings_FieldService', 'FSBoardKioskGlobalColorAwaitingAcknowledgement') AS IsItThere"
	Set rsSettingsFieldService = cnnSettingsFieldService.Execute(SQL_SettingsFieldService)
	If IsNull(rsSettingsFieldService("IsItThere")) Then
		SQL_SettingsFieldService = "ALTER TABLE Settings_FieldService ADD FSBoardKioskGlobalColorAwaitingAcknowledgement varchar(50) NULL"
		Set rsSettingsFieldService = cnnSettingsFieldService.Execute(SQL_SettingsFieldService)
		SQL_SettingsFieldService = "UPDATE Settings_FieldService SET FSBoardKioskGlobalColorAwaitingAcknowledgement = '#FC0802'"
		Set rsSettingsFieldService = cnnSettingsFieldService.Execute(SQL_SettingsFieldService)
	End If

	SQL_SettingsFieldService = "SELECT COL_LENGTH('Settings_FieldService', 'FSBoardKioskGlobalColorDispatchAcknowledged') AS IsItThere"
	Set rsSettingsFieldService = cnnSettingsFieldService.Execute(SQL_SettingsFieldService)
	If IsNull(rsSettingsFieldService("IsItThere")) Then
		SQL_SettingsFieldService = "ALTER TABLE Settings_FieldService ADD FSBoardKioskGlobalColorDispatchAcknowledged varchar(50) NULL"
		Set rsSettingsFieldService = cnnSettingsFieldService.Execute(SQL_SettingsFieldService)
		SQL_SettingsFieldService = "UPDATE Settings_FieldService SET FSBoardKioskGlobalColorDispatchAcknowledged = '#4474f7'"
		Set rsSettingsFieldService = cnnSettingsFieldService.Execute(SQL_SettingsFieldService)
	End If

	SQL_SettingsFieldService = "SELECT COL_LENGTH('Settings_FieldService', 'FSBoardKioskGlobalColorDispatchDeclined') AS IsItThere"
	Set rsSettingsFieldService = cnnSettingsFieldService.Execute(SQL_SettingsFieldService)
	If IsNull(rsSettingsFieldService("IsItThere")) Then
		SQL_SettingsFieldService = "ALTER TABLE Settings_FieldService ADD FSBoardKioskGlobalColorDispatchDeclined varchar(50) NULL"
		Set rsSettingsFieldService = cnnSettingsFieldService.Execute(SQL_SettingsFieldService)
		SQL_SettingsFieldService = "UPDATE Settings_FieldService SET FSBoardKioskGlobalColorDispatchDeclined = '#9AB0D5'"
		Set rsSettingsFieldService = cnnSettingsFieldService.Execute(SQL_SettingsFieldService)
	End If

	SQL_SettingsFieldService = "SELECT COL_LENGTH('Settings_FieldService', 'FSBoardKioskGlobalColorEnRoute') AS IsItThere"
	Set rsSettingsFieldService = cnnSettingsFieldService.Execute(SQL_SettingsFieldService)
	If IsNull(rsSettingsFieldService("IsItThere")) Then
		SQL_SettingsFieldService = "ALTER TABLE Settings_FieldService ADD FSBoardKioskGlobalColorEnRoute varchar(50) NULL"
		Set rsSettingsFieldService = cnnSettingsFieldService.Execute(SQL_SettingsFieldService)
		SQL_SettingsFieldService = "UPDATE Settings_FieldService SET FSBoardKioskGlobalColorEnRoute = '#bbbb40'"
		Set rsSettingsFieldService = cnnSettingsFieldService.Execute(SQL_SettingsFieldService)
	End If

	SQL_SettingsFieldService = "SELECT COL_LENGTH('Settings_FieldService', 'FSBoardKioskGlobalColorOnSite') AS IsItThere"
	Set rsSettingsFieldService = cnnSettingsFieldService.Execute(SQL_SettingsFieldService)
	If IsNull(rsSettingsFieldService("IsItThere")) Then
		SQL_SettingsFieldService = "ALTER TABLE Settings_FieldService ADD FSBoardKioskGlobalColorOnSite varchar(50) NULL"
		Set rsSettingsFieldService = cnnSettingsFieldService.Execute(SQL_SettingsFieldService)
		SQL_SettingsFieldService = "UPDATE Settings_FieldService SET FSBoardKioskGlobalColorOnSite = '#4cae4c'"
		Set rsSettingsFieldService = cnnSettingsFieldService.Execute(SQL_SettingsFieldService)
	End If

	SQL_SettingsFieldService = "SELECT COL_LENGTH('Settings_FieldService', 'FSBoardKioskGlobalColorRedoSwap') AS IsItThere"
	Set rsSettingsFieldService = cnnSettingsFieldService.Execute(SQL_SettingsFieldService)
	If IsNull(rsSettingsFieldService("IsItThere")) Then
		SQL_SettingsFieldService = "ALTER TABLE Settings_FieldService ADD FSBoardKioskGlobalColorRedoSwap varchar(50) NULL"
		Set rsSettingsFieldService = cnnSettingsFieldService.Execute(SQL_SettingsFieldService)
		SQL_SettingsFieldService = "UPDATE Settings_FieldService SET FSBoardKioskGlobalColorRedoSwap = '#666666'"
		Set rsSettingsFieldService = cnnSettingsFieldService.Execute(SQL_SettingsFieldService)
	End If

	SQL_SettingsFieldService = "SELECT COL_LENGTH('Settings_FieldService', 'FSBoardKioskGlobalColorRedoWaitForParts') AS IsItThere"
	Set rsSettingsFieldService = cnnSettingsFieldService.Execute(SQL_SettingsFieldService)
	If IsNull(rsSettingsFieldService("IsItThere")) Then
		SQL_SettingsFieldService = "ALTER TABLE Settings_FieldService ADD FSBoardKioskGlobalColorRedoWaitForParts varchar(50) NULL"
		Set rsSettingsFieldService = cnnSettingsFieldService.Execute(SQL_SettingsFieldService)
		SQL_SettingsFieldService = "UPDATE Settings_FieldService SET FSBoardKioskGlobalColorRedoWaitForParts = '#ca24ca'"
		Set rsSettingsFieldService = cnnSettingsFieldService.Execute(SQL_SettingsFieldService)
	End If

	SQL_SettingsFieldService = "SELECT COL_LENGTH('Settings_FieldService', 'FSBoardKioskGlobalColorRedoFollowUp') AS IsItThere"
	Set rsSettingsFieldService = cnnSettingsFieldService.Execute(SQL_SettingsFieldService)
	If IsNull(rsSettingsFieldService("IsItThere")) Then
		SQL_SettingsFieldService = "ALTER TABLE Settings_FieldService ADD FSBoardKioskGlobalColorRedoFollowUp varchar(50) NULL"
		Set rsSettingsFieldService = cnnSettingsFieldService.Execute(SQL_SettingsFieldService)
		SQL_SettingsFieldService = "UPDATE Settings_FieldService SET FSBoardKioskGlobalColorRedoFollowUp = '#7b0bc7'"
		Set rsSettingsFieldService = cnnSettingsFieldService.Execute(SQL_SettingsFieldService)
	End If

	SQL_SettingsFieldService = "SELECT COL_LENGTH('Settings_FieldService', 'FSBoardKioskGlobalColorRedoUnableToWork') AS IsItThere"
	Set rsSettingsFieldService = cnnSettingsFieldService.Execute(SQL_SettingsFieldService)
	If IsNull(rsSettingsFieldService("IsItThere")) Then
		SQL_SettingsFieldService = "ALTER TABLE Settings_FieldService ADD FSBoardKioskGlobalColorRedoUnableToWork varchar(50) NULL"
		Set rsSettingsFieldService = cnnSettingsFieldService.Execute(SQL_SettingsFieldService)
		SQL_SettingsFieldService = "UPDATE Settings_FieldService SET FSBoardKioskGlobalColorRedoUnableToWork = '#cc0000'"
		Set rsSettingsFieldService = cnnSettingsFieldService.Execute(SQL_SettingsFieldService)
	End If

	SQL_SettingsFieldService = "SELECT COL_LENGTH('Settings_FieldService', 'FSBoardKioskGlobalColorClosed') AS IsItThere"
	Set rsSettingsFieldService = cnnSettingsFieldService.Execute(SQL_SettingsFieldService)
	If IsNull(rsSettingsFieldService("IsItThere")) Then
		SQL_SettingsFieldService = "ALTER TABLE Settings_FieldService ADD FSBoardKioskGlobalColorClosed varchar(50) NULL"
		Set rsSettingsFieldService = cnnSettingsFieldService.Execute(SQL_SettingsFieldService)
		SQL_SettingsFieldService = "UPDATE Settings_FieldService SET FSBoardKioskGlobalColorClosed = '#dbefd2'"
		Set rsSettingsFieldService = cnnSettingsFieldService.Execute(SQL_SettingsFieldService)
	End If

	SQL_SettingsFieldService = "SELECT COL_LENGTH('Settings_FieldService', 'FSBoardKioskGlobalColorUrgent') AS IsItThere"
	Set rsSettingsFieldService = cnnSettingsFieldService.Execute(SQL_SettingsFieldService)
	If IsNull(rsSettingsFieldService("IsItThere")) Then
		SQL_SettingsFieldService = "ALTER TABLE Settings_FieldService ADD FSBoardKioskGlobalColorUrgent varchar(50) NULL"
		Set rsSettingsFieldService = cnnSettingsFieldService.Execute(SQL_SettingsFieldService)
		SQL_SettingsFieldService = "UPDATE Settings_FieldService SET FSBoardKioskGlobalColorUrgent = '#00ff00'"
		Set rsSettingsFieldService = cnnSettingsFieldService.Execute(SQL_SettingsFieldService)
	End If
	
	'*****************************************************
	'END NEW FIELD SERVICE GLOBAL COLOR SETTINGS
	'*****************************************************

	set rsSettingsFieldService = nothing
	cnnSettingsFieldService.close
	set cnnSettingsFieldService = nothing
%>