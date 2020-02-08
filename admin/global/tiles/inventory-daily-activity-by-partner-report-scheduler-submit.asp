<!--#include file="../../../inc/header.asp"-->

<%

	'***********************************************************
	'Get Values Of All Form Fields Posted
	'***********************************************************
	
	GenerateDailyInventoryAPIActivityByPartnerReportSunday = Request.Form("chkNoDailyInventoryAPIActivityByPartnerReportSunday")
	GenerateDailyInventoryAPIActivityByPartnerReportMonday = Request.Form("chkNoDailyInventoryAPIActivityByPartnerReportMonday")
	GenerateDailyInventoryAPIActivityByPartnerReportTuesday = Request.Form("chkNoDailyInventoryAPIActivityByPartnerReportTuesday")
	GenerateDailyInventoryAPIActivityByPartnerReportWednesday = Request.Form("chkNoDailyInventoryAPIActivityByPartnerReportWednesday")
	GenerateDailyInventoryAPIActivityByPartnerReportThursday = Request.Form("chkNoDailyInventoryAPIActivityByPartnerReportThursday")
	GenerateDailyInventoryAPIActivityByPartnerReportFriday = Request.Form("chkNoDailyInventoryAPIActivityByPartnerReportFriday")
	GenerateDailyInventoryAPIActivityByPartnerReportSaturday = Request.Form("chkNoDailyInventoryAPIActivityByPartnerReportSaturday")
	
	GenerateDailyInventoryAPIActivityByPartnerReportSundayTime = Request.Form("txtDailyInventoryAPIActivityByPartnerReportSchedulerSundayTime")
	GenerateDailyInventoryAPIActivityByPartnerReportMondayTime = Request.Form("txtDailyInventoryAPIActivityByPartnerReportSchedulerMondayTime")
	GenerateDailyInventoryAPIActivityByPartnerReportTuesdayTime = Request.Form("txtDailyInventoryAPIActivityByPartnerReportSchedulerTuesdayTime")
	GenerateDailyInventoryAPIActivityByPartnerReportWednesdayTime = Request.Form("txtDailyInventoryAPIActivityByPartnerReportSchedulerWednesdayTime")
	GenerateDailyInventoryAPIActivityByPartnerReportThursdayTime = Request.Form("txtDailyInventoryAPIActivityByPartnerReportSchedulerThursdayTime")
	GenerateDailyInventoryAPIActivityByPartnerReportFridayTime = Request.Form("txtDailyInventoryAPIActivityByPartnerReportSchedulerFridayTime")
	GenerateDailyInventoryAPIActivityByPartnerReportSaturdayTime = Request.Form("txtDailyInventoryAPIActivityByPartnerReportSchedulerSaturdayTime")
	
	RunDailyInventoryAPIActivityByPartnerReportIfClosed = Request.Form("chkNoDailyInventoryAPIActivityByPartnerReportIfClosed")
	RunDailyInventoryAPIActivityByPartnerReportIfClosingEarly = Request.Form("chkNoDailyInventoryAPIActivityByPartnerReportIfClosingEarly")


	If Request.Form("chkNoDailyInventoryAPIActivityByPartnerReportSunday") = "on" Then
		GenerateDailyInventoryAPIActivityByPartnerReportSunday = 0
		GenerateDailyInventoryAPIActivityByPartnerReportSundayTime = ""
	Else 
		GenerateDailyInventoryAPIActivityByPartnerReportSunday = 1
	End If

	If Request.Form("chkNoDailyInventoryAPIActivityByPartnerReportMonday") = "on" Then
		GenerateDailyInventoryAPIActivityByPartnerReportMonday = 0
		GenerateDailyInventoryAPIActivityByPartnerReportMondayTime = ""
	Else 
		GenerateDailyInventoryAPIActivityByPartnerReportMonday = 1
	End If

	If Request.Form("chkNoDailyInventoryAPIActivityByPartnerReportTuesday") = "on" Then
		GenerateDailyInventoryAPIActivityByPartnerReportTuesday = 0
		GenerateDailyInventoryAPIActivityByPartnerReportTuesdayTime = ""
	Else 
		GenerateDailyInventoryAPIActivityByPartnerReportTuesday = 1
	End If

	If Request.Form("chkNoDailyInventoryAPIActivityByPartnerReportWednesday") = "on" Then
		GenerateDailyInventoryAPIActivityByPartnerReportWednesday = 0
		GenerateDailyInventoryAPIActivityByPartnerReportWednesdayTime = ""
	Else 
		GenerateDailyInventoryAPIActivityByPartnerReportWednesday = 1
	End If

	If Request.Form("chkNoDailyInventoryAPIActivityByPartnerReportThursday") = "on" Then
		GenerateDailyInventoryAPIActivityByPartnerReportThursday = 0
		GenerateDailyInventoryAPIActivityByPartnerReportThursdayTime = ""
	Else 
		GenerateDailyInventoryAPIActivityByPartnerReportThursday = 1
	End If

	If Request.Form("chkNoDailyInventoryAPIActivityByPartnerReportFriday") = "on" Then
		GenerateDailyInventoryAPIActivityByPartnerReportFriday = 0
		GenerateDailyInventoryAPIActivityByPartnerReportFridayTime = ""
	Else 
		GenerateDailyInventoryAPIActivityByPartnerReportFriday = 1
	End If

	If Request.Form("chkNoDailyInventoryAPIActivityByPartnerReportSaturday") = "on" Then
		GenerateDailyInventoryAPIActivityByPartnerReportSaturday = 0
		GenerateDailyInventoryAPIActivityByPartnerReportSaturdayTime = ""
	Else 
		GenerateDailyInventoryAPIActivityByPartnerReportSaturday = 1
	End If

	If Request.Form("chkNoDailyInventoryAPIActivityByPartnerReportIfClosed") = "on" Then RunDailyInventoryAPIActivityByPartnerReportIfClosed = 0 Else RunDailyInventoryAPIActivityByPartnerReportIfClosed = 1
	If Request.Form("chkNoDailyInventoryAPIActivityByPartnerReportIfClosingEarly") = "on" Then RunDailyInventoryAPIActivityByPartnerReportIfClosingEarly = 0 Else RunDailyInventoryAPIActivityByPartnerReportIfClosingEarly = 1
	
	'***********************************************************
	'Get Original Values For Audit Trail Entries
	'***********************************************************
	
	SQLFieldServiceSettings = "SELECT * FROM Settings_InventoryControl"
	
	Set cnnFieldServiceSettings = Server.CreateObject("ADODB.Connection")
	cnnFieldServiceSettings.open (Session("ClientCnnString"))
	Set rsFieldServiceSettings = Server.CreateObject("ADODB.Recordset")
	rsFieldServiceSettings.CursorLocation = 3 
	Set rsFieldServiceSettings = cnnFieldServiceSettings.Execute(SQLFieldServiceSettings)
		
	If NOT rsFieldServiceSettings.EOF Then
	
		Schedule_DailyInventoryAPIActivityByPartnerReportGeneration = rsFieldServiceSettings("Schedule_DailyInventoryAPIActivityByPartnerReportGeneration")
		
		Schedule_DailyInventoryAPIActivityByPartnerReportGenerationSettings = Split(Schedule_DailyInventoryAPIActivityByPartnerReportGeneration,",")

		GenerateDailyInventoryAPIActivityByPartnerReportSunday_ORIG = cInt(Schedule_DailyInventoryAPIActivityByPartnerReportGenerationSettings(0))
		GenerateDailyInventoryAPIActivityByPartnerReportMonday_ORIG = cInt(Schedule_DailyInventoryAPIActivityByPartnerReportGenerationSettings(1))
		GenerateDailyInventoryAPIActivityByPartnerReportTuesday_ORIG = cInt(Schedule_DailyInventoryAPIActivityByPartnerReportGenerationSettings(2))
		GenerateDailyInventoryAPIActivityByPartnerReportWednesday_ORIG = cInt(Schedule_DailyInventoryAPIActivityByPartnerReportGenerationSettings(3))
		GenerateDailyInventoryAPIActivityByPartnerReportThursday_ORIG = cInt(Schedule_DailyInventoryAPIActivityByPartnerReportGenerationSettings(4))
		GenerateDailyInventoryAPIActivityByPartnerReportFriday_ORIG = cInt(Schedule_DailyInventoryAPIActivityByPartnerReportGenerationSettings(5))
		GenerateDailyInventoryAPIActivityByPartnerReportSaturday_ORIG = cInt(Schedule_DailyInventoryAPIActivityByPartnerReportGenerationSettings(6))
		GenerateDailyInventoryAPIActivityByPartnerReportSundayTime_ORIG = Schedule_DailyInventoryAPIActivityByPartnerReportGenerationSettings(7)
		GenerateDailyInventoryAPIActivityByPartnerReportMondayTime_ORIG = Schedule_DailyInventoryAPIActivityByPartnerReportGenerationSettings(8)
		GenerateDailyInventoryAPIActivityByPartnerReportTuesdayTime_ORIG = Schedule_DailyInventoryAPIActivityByPartnerReportGenerationSettings(9)
		GenerateDailyInventoryAPIActivityByPartnerReportWednesdayTime_ORIG = Schedule_DailyInventoryAPIActivityByPartnerReportGenerationSettings(10)
		GenerateDailyInventoryAPIActivityByPartnerReportThursdayTime_ORIG = Schedule_DailyInventoryAPIActivityByPartnerReportGenerationSettings(11)
		GenerateDailyInventoryAPIActivityByPartnerReportFridayTime_ORIG = Schedule_DailyInventoryAPIActivityByPartnerReportGenerationSettings(12)
		GenerateDailyInventoryAPIActivityByPartnerReportSaturdayTime_ORIG = Schedule_DailyInventoryAPIActivityByPartnerReportGenerationSettings(13)
		RunDailyInventoryAPIActivityByPartnerReportIfClosed_ORIG = cInt(Schedule_DailyInventoryAPIActivityByPartnerReportGenerationSettings(14))
		RunDailyInventoryAPIActivityByPartnerReportIfClosingEarly_ORIG = cInt(Schedule_DailyInventoryAPIActivityByPartnerReportGenerationSettings(15))
	
	End If
	
	set rsFieldServiceSettings = Nothing
	cnnFieldServiceSettings.close
	set cnnFieldServiceSettings = Nothing
	

	'***********************************************************
	'Perform Audit Trail Entry Inserts
	'***********************************************************

	
	If Request.Form("chkNoDailyInventoryAPIActivityByPartnerReportSunday") = "on" then GenerateDailyInventoryAPIActivityByPartnerReportSundayMsg = "On" Else GenerateDailyInventoryAPIActivityByPartnerReportSundayMsg = "Off"
	If GenerateDailyInventoryAPIActivityByPartnerReportSunday_ORIG = 1 then GenerateDailyInventoryAPIActivityByPartnerReportSundayMsgOrig = "On" Else GenerateDailyInventoryAPIActivityByPartnerReportSundayMsgOrig = "Off"
	
	If GenerateDailyInventoryAPIActivityByPartnerReportSunday <> GenerateDailyInventoryAPIActivityByPartnerReportSunday_ORIG Then
		CreateAuditLogEntry "Global Settings Change", "Global Settings Change", "Major", 1, "Daily Inventory API Activity By Partner Report generation schedule for SUNDAY changed from " & GenerateDailyInventoryAPIActivityByPartnerReportSundayMsgOrig & " to " & GenerateDailyInventoryAPIActivityByPartnerReportSundayMsg
	End If
	
	If GenerateDailyInventoryAPIActivityByPartnerReportSundayTime <> GenerateDailyInventoryAPIActivityByPartnerReportSundayTime_ORIG Then
		If GenerateDailyInventoryAPIActivityByPartnerReportSunday_ORIG = 0 AND GenerateDailyInventoryAPIActivityByPartnerReportSunday = 1 Then
			CreateAuditLogEntry "Global Settings Change", "Global Settings Change", "Major", 1, "Daily Inventory API Activity By Partner Report generation schedule for SUNDAY turned on and set to run at " & GenerateDailyInventoryAPIActivityByPartnerReportSundayTime
		Else
			CreateAuditLogEntry "Global Settings Change", "Global Settings Change", "Major", 1, "Daily Inventory API Activity By Partner Report generation scheduled run time for SUNDAY changed from " & GenerateDailyInventoryAPIActivityByPartnerReportSundayTime_ORIG & " to " & GenerateDailyInventoryAPIActivityByPartnerReportSundayTime
		End If
	End If
	
	

	If Request.Form("chkNoDailyInventoryAPIActivityByPartnerReportMonday") = "on" then GenerateDailyInventoryAPIActivityByPartnerReportMondayMsg = "On" Else GenerateDailyInventoryAPIActivityByPartnerReportMondayMsg = "Off"
	If GenerateDailyInventoryAPIActivityByPartnerReportMonday_ORIG = 1 then GenerateDailyInventoryAPIActivityByPartnerReportMondayMsgOrig = "On" Else GenerateDailyInventoryAPIActivityByPartnerReportMondayMsgOrig = "Off"
	
	If GenerateDailyInventoryAPIActivityByPartnerReportMonday <> GenerateDailyInventoryAPIActivityByPartnerReportMonday_ORIG Then
		CreateAuditLogEntry "Global Settings Change", "Global Settings Change", "Major", 1, "Daily Inventory API Activity By Partner Report generation schedule for Monday changed from " & GenerateDailyInventoryAPIActivityByPartnerReportMondayMsgOrig & " to " & GenerateDailyInventoryAPIActivityByPartnerReportMondayMsg
	End If
	
	If GenerateDailyInventoryAPIActivityByPartnerReportMondayTime <> GenerateDailyInventoryAPIActivityByPartnerReportMondayTime_ORIG Then
		If GenerateDailyInventoryAPIActivityByPartnerReportMonday_ORIG = 0 AND GenerateDailyInventoryAPIActivityByPartnerReportMonday = 1 Then
			CreateAuditLogEntry "Global Settings Change", "Global Settings Change", "Major", 1, "Daily Inventory API Activity By Partner Report generation schedule for Monday turned on and set to run at " & GenerateDailyInventoryAPIActivityByPartnerReportMondayTime
		Else
			CreateAuditLogEntry "Global Settings Change", "Global Settings Change", "Major", 1, "Daily Inventory API Activity By Partner Report generation scheduled run time for Monday changed from " & GenerateDailyInventoryAPIActivityByPartnerReportMondayTime_ORIG & " to " & GenerateDailyInventoryAPIActivityByPartnerReportMondayTime
		End If
	End If
	


	If Request.Form("chkNoDailyInventoryAPIActivityByPartnerReportTuesday") = "on" then GenerateDailyInventoryAPIActivityByPartnerReportTuesdayMsg = "On" Else GenerateDailyInventoryAPIActivityByPartnerReportTuesdayMsg = "Off"
	If GenerateDailyInventoryAPIActivityByPartnerReportTuesday_ORIG = 1 then GenerateDailyInventoryAPIActivityByPartnerReportTuesdayMsgOrig = "On" Else GenerateDailyInventoryAPIActivityByPartnerReportTuesdayMsgOrig = "Off"
	
	If GenerateDailyInventoryAPIActivityByPartnerReportTuesday <> GenerateDailyInventoryAPIActivityByPartnerReportTuesday_ORIG Then
		CreateAuditLogEntry "Global Settings Change", "Global Settings Change", "Major", 1, "Daily Inventory API Activity By Partner Report generation schedule for Tuesday changed from " & GenerateDailyInventoryAPIActivityByPartnerReportTuesdayMsgOrig & " to " & GenerateDailyInventoryAPIActivityByPartnerReportTuesdayMsg
	End If
	
	If GenerateDailyInventoryAPIActivityByPartnerReportTuesdayTime <> GenerateDailyInventoryAPIActivityByPartnerReportTuesdayTime_ORIG Then
		If GenerateDailyInventoryAPIActivityByPartnerReportTuesday_ORIG = 0 AND GenerateDailyInventoryAPIActivityByPartnerReportTuesday = 1 Then
			CreateAuditLogEntry "Global Settings Change", "Global Settings Change", "Major", 1, "Daily Inventory API Activity By Partner Report generation schedule for Tuesday turned on and set to run at " & GenerateDailyInventoryAPIActivityByPartnerReportTuesdayTime
		Else
			CreateAuditLogEntry "Global Settings Change", "Global Settings Change", "Major", 1, "Daily Inventory API Activity By Partner Report generation scheduled run time for Tuesday changed from " & GenerateDailyInventoryAPIActivityByPartnerReportTuesdayTime_ORIG & " to " & GenerateDailyInventoryAPIActivityByPartnerReportTuesdayTime
		End If
	End If



	If Request.Form("chkNoDailyInventoryAPIActivityByPartnerReportWednesday") = "on" then GenerateDailyInventoryAPIActivityByPartnerReportWednesdayMsg = "On" Else GenerateDailyInventoryAPIActivityByPartnerReportWednesdayMsg = "Off"
	If GenerateDailyInventoryAPIActivityByPartnerReportWednesday_ORIG = 1 then GenerateDailyInventoryAPIActivityByPartnerReportWednesdayMsgOrig = "On" Else GenerateDailyInventoryAPIActivityByPartnerReportWednesdayMsgOrig = "Off"
	
	If GenerateDailyInventoryAPIActivityByPartnerReportWednesday <> GenerateDailyInventoryAPIActivityByPartnerReportWednesday_ORIG Then
		CreateAuditLogEntry "Global Settings Change", "Global Settings Change", "Major", 1, "Daily Inventory API Activity By Partner Report generation schedule for Wednesday changed from " & GenerateDailyInventoryAPIActivityByPartnerReportWednesdayMsgOrig & " to " & GenerateDailyInventoryAPIActivityByPartnerReportWednesdayMsg
	End If
	
	If GenerateDailyInventoryAPIActivityByPartnerReportWednesdayTime <> GenerateDailyInventoryAPIActivityByPartnerReportWednesdayTime_ORIG Then
		If GenerateDailyInventoryAPIActivityByPartnerReportWednesday_ORIG = 0 AND GenerateDailyInventoryAPIActivityByPartnerReportWednesday = 1 Then
			CreateAuditLogEntry "Global Settings Change", "Global Settings Change", "Major", 1, "Daily Inventory API Activity By Partner Report generation schedule for Wednesday turned on and set to run at " & GenerateDailyInventoryAPIActivityByPartnerReportWednesdayTime
		Else
			CreateAuditLogEntry "Global Settings Change", "Global Settings Change", "Major", 1, "Daily Inventory API Activity By Partner Report generation scheduled run time for Wednesday changed from " & GenerateDailyInventoryAPIActivityByPartnerReportWednesdayTime_ORIG & " to " & GenerateDailyInventoryAPIActivityByPartnerReportWednesdayTime
		End If
	End If



	If Request.Form("chkNoDailyInventoryAPIActivityByPartnerReportThursday") = "on" then GenerateDailyInventoryAPIActivityByPartnerReportThursdayMsg = "On" Else GenerateDailyInventoryAPIActivityByPartnerReportThursdayMsg = "Off"
	If GenerateDailyInventoryAPIActivityByPartnerReportThursday_ORIG = 1 then GenerateDailyInventoryAPIActivityByPartnerReportThursdayMsgOrig = "On" Else GenerateDailyInventoryAPIActivityByPartnerReportThursdayMsgOrig = "Off"
	
	If GenerateDailyInventoryAPIActivityByPartnerReportThursday <> GenerateDailyInventoryAPIActivityByPartnerReportThursday_ORIG Then
		CreateAuditLogEntry "Global Settings Change", "Global Settings Change", "Major", 1, "Daily Inventory API Activity By Partner Report generation schedule for Thursday changed from " & GenerateDailyInventoryAPIActivityByPartnerReportThursdayMsgOrig & " to " & GenerateDailyInventoryAPIActivityByPartnerReportThursdayMsg
	End If
	
	If GenerateDailyInventoryAPIActivityByPartnerReportThursdayTime <> GenerateDailyInventoryAPIActivityByPartnerReportThursdayTime_ORIG Then
		If GenerateDailyInventoryAPIActivityByPartnerReportThursday_ORIG = 0 AND GenerateDailyInventoryAPIActivityByPartnerReportThursday = 1 Then
			CreateAuditLogEntry "Global Settings Change", "Global Settings Change", "Major", 1, "Daily Inventory API Activity By Partner Report generation schedule for Thursday turned on and set to run at " & GenerateDailyInventoryAPIActivityByPartnerReportThursdayTime
		Else
			CreateAuditLogEntry "Global Settings Change", "Global Settings Change", "Major", 1, "Daily Inventory API Activity By Partner Report generation scheduled run time for Thursday changed from " & GenerateDailyInventoryAPIActivityByPartnerReportThursdayTime_ORIG & " to " & GenerateDailyInventoryAPIActivityByPartnerReportThursdayTime
		End If
	End If



	If Request.Form("chkNoDailyInventoryAPIActivityByPartnerReportFriday") = "on" then GenerateDailyInventoryAPIActivityByPartnerReportFridayMsg = "On" Else GenerateDailyInventoryAPIActivityByPartnerReportFridayMsg = "Off"
	If GenerateDailyInventoryAPIActivityByPartnerReportFriday_ORIG = 1 then GenerateDailyInventoryAPIActivityByPartnerReportFridayMsgOrig = "On" Else GenerateDailyInventoryAPIActivityByPartnerReportFridayMsgOrig = "Off"
	
	If GenerateDailyInventoryAPIActivityByPartnerReportFriday <> GenerateDailyInventoryAPIActivityByPartnerReportFriday_ORIG Then
		CreateAuditLogEntry "Global Settings Change", "Global Settings Change", "Major", 1, "Daily Inventory API Activity By Partner Report generation schedule for Friday changed from " & GenerateDailyInventoryAPIActivityByPartnerReportFridayMsgOrig & " to " & GenerateDailyInventoryAPIActivityByPartnerReportFridayMsg
	End If
	
	If GenerateDailyInventoryAPIActivityByPartnerReportFridayTime <> GenerateDailyInventoryAPIActivityByPartnerReportFridayTime_ORIG Then
		If GenerateDailyInventoryAPIActivityByPartnerReportFriday_ORIG = 0 AND GenerateDailyInventoryAPIActivityByPartnerReportFriday = 1 Then
			CreateAuditLogEntry "Global Settings Change", "Global Settings Change", "Major", 1, "Daily Inventory API Activity By Partner Report generation schedule for Friday turned on and set to run at " & GenerateDailyInventoryAPIActivityByPartnerReportFridayTime
		Else
			CreateAuditLogEntry "Global Settings Change", "Global Settings Change", "Major", 1, "Daily Inventory API Activity By Partner Report generation scheduled run time for Friday changed from " & GenerateDailyInventoryAPIActivityByPartnerReportFridayTime_ORIG & " to " & GenerateDailyInventoryAPIActivityByPartnerReportFridayTime
		End If
	End If



	If Request.Form("chkNoDailyInventoryAPIActivityByPartnerReportSaturday") = "on" then GenerateDailyInventoryAPIActivityByPartnerReportSaturdayMsg = "On" Else GenerateDailyInventoryAPIActivityByPartnerReportSaturdayMsg = "Off"
	If GenerateDailyInventoryAPIActivityByPartnerReportSaturday_ORIG = 1 then GenerateDailyInventoryAPIActivityByPartnerReportSaturdayMsgOrig = "On" Else GenerateDailyInventoryAPIActivityByPartnerReportSaturdayMsgOrig = "Off"
	
	If GenerateDailyInventoryAPIActivityByPartnerReportSaturday <> GenerateDailyInventoryAPIActivityByPartnerReportSaturday_ORIG Then
		CreateAuditLogEntry "Global Settings Change", "Global Settings Change", "Major", 1, "Daily Inventory API Activity By Partner Report generation schedule for Saturday changed from " & GenerateDailyInventoryAPIActivityByPartnerReportSaturdayMsgOrig & " to " & GenerateDailyInventoryAPIActivityByPartnerReportSaturdayMsg
	End If
	
	If GenerateDailyInventoryAPIActivityByPartnerReportSaturdayTime <> GenerateDailyInventoryAPIActivityByPartnerReportSaturdayTime_ORIG Then
		If GenerateDailyInventoryAPIActivityByPartnerReportSaturday_ORIG = 0 AND GenerateDailyInventoryAPIActivityByPartnerReportSaturday = 1 Then
			CreateAuditLogEntry "Global Settings Change", "Global Settings Change", "Major", 1, "Daily Inventory API Activity By Partner Report generation schedule for Saturday turned on and set to run at " & GenerateDailyInventoryAPIActivityByPartnerReportSaturdayTime
		Else
			CreateAuditLogEntry "Global Settings Change", "Global Settings Change", "Major", 1, "Daily Inventory API Activity By Partner Report generation scheduled run time for Saturday changed from " & GenerateDailyInventoryAPIActivityByPartnerReportSaturdayTime_ORIG & " to " & GenerateDailyInventoryAPIActivityByPartnerReportSaturdayTime
		End If
	End If


	If Request.Form("chkNoDailyInventoryAPIActivityByPartnerReportIfClosed") = "on" then RunDailyInventoryAPIActivityByPartnerReportIfClosedMsg = "On" Else RunDailyInventoryAPIActivityByPartnerReportIfClosedMsg = "Off"
	If RunDailyInventoryAPIActivityByPartnerReportIfClosed_ORIG = 1 then RunDailyInventoryAPIActivityByPartnerReportIfClosedMsgOrig = "On" Else RunDailyInventoryAPIActivityByPartnerReportIfClosedMsgOrig = "Off"
	
	If RunDailyInventoryAPIActivityByPartnerReportIfClosed <> RunDailyInventoryAPIActivityByPartnerReportIfClosed_ORIG Then
		CreateAuditLogEntry "Global Settings Change", "Global Settings Change", "Major", 1, "Daily Inventory API Activity By Partner Report generation schedule - Don''t Run If Closed (Monday-Friday Only) -  changed from " & RunDailyInventoryAPIActivityByPartnerReportIfClosedMsgOrig & " to " & RunDailyInventoryAPIActivityByPartnerReportIfClosedMsg
	End If


	If Request.Form("chkNoDailyInventoryAPIActivityByPartnerReportIfClosingEarly") = "on" then RunDailyInventoryAPIActivityByPartnerReportIfClosingEarlyMsg = "On" Else RunDailyInventoryAPIActivityByPartnerReportIfClosingEarlyMsg = "Off"
	If RunDailyInventoryAPIActivityByPartnerReportIfClosingEarly_ORIG = 1 then RunDailyInventoryAPIActivityByPartnerReportIfClosingEarlyMsgOrig = "On" Else RunDailyInventoryAPIActivityByPartnerReportIfClosingEarlyMsgOrig = "Off"
	
	If RunDailyInventoryAPIActivityByPartnerReportIfClosingEarly <> RunDailyInventoryAPIActivityByPartnerReportIfClosingEarly_ORIG Then
		CreateAuditLogEntry "Global Settings Change", "Global Settings Change", "Major", 1, "Daily Inventory API Activity By Partner Report generation schedule - Don''t Run If Closed (Monday-Friday Only) -  changed from " & RunDailyInventoryAPIActivityByPartnerReportIfClosingEarlyMsgOrig & " to " & RunDailyInventoryAPIActivityByPartnerReportIfClosingEarlyMsg
	End If


	'*********************************************************************
	'Build Array/String of Schedule Data From Request Form Field Data
	'*********************************************************************

	Schedule_DailyInventoryAPIActivityByPartnerReportGenerationUpdated = ""
	
	Schedule_DailyInventoryAPIActivityByPartnerReportGenerationUpdated = GenerateDailyInventoryAPIActivityByPartnerReportSunday
	Schedule_DailyInventoryAPIActivityByPartnerReportGenerationUpdated = Schedule_DailyInventoryAPIActivityByPartnerReportGenerationUpdated & "," & GenerateDailyInventoryAPIActivityByPartnerReportMonday
	Schedule_DailyInventoryAPIActivityByPartnerReportGenerationUpdated = Schedule_DailyInventoryAPIActivityByPartnerReportGenerationUpdated & "," & GenerateDailyInventoryAPIActivityByPartnerReportTuesday
	Schedule_DailyInventoryAPIActivityByPartnerReportGenerationUpdated = Schedule_DailyInventoryAPIActivityByPartnerReportGenerationUpdated & "," & GenerateDailyInventoryAPIActivityByPartnerReportWednesday
	Schedule_DailyInventoryAPIActivityByPartnerReportGenerationUpdated = Schedule_DailyInventoryAPIActivityByPartnerReportGenerationUpdated & "," & GenerateDailyInventoryAPIActivityByPartnerReportThursday
	Schedule_DailyInventoryAPIActivityByPartnerReportGenerationUpdated = Schedule_DailyInventoryAPIActivityByPartnerReportGenerationUpdated & "," & GenerateDailyInventoryAPIActivityByPartnerReportFriday
	Schedule_DailyInventoryAPIActivityByPartnerReportGenerationUpdated = Schedule_DailyInventoryAPIActivityByPartnerReportGenerationUpdated & "," & GenerateDailyInventoryAPIActivityByPartnerReportSaturday
	
	Schedule_DailyInventoryAPIActivityByPartnerReportGenerationUpdated = Schedule_DailyInventoryAPIActivityByPartnerReportGenerationUpdated & "," & GenerateDailyInventoryAPIActivityByPartnerReportSundayTime
	Schedule_DailyInventoryAPIActivityByPartnerReportGenerationUpdated = Schedule_DailyInventoryAPIActivityByPartnerReportGenerationUpdated & "," & GenerateDailyInventoryAPIActivityByPartnerReportMondayTime
	Schedule_DailyInventoryAPIActivityByPartnerReportGenerationUpdated = Schedule_DailyInventoryAPIActivityByPartnerReportGenerationUpdated & "," & GenerateDailyInventoryAPIActivityByPartnerReportTuesdayTime
	Schedule_DailyInventoryAPIActivityByPartnerReportGenerationUpdated = Schedule_DailyInventoryAPIActivityByPartnerReportGenerationUpdated & "," & GenerateDailyInventoryAPIActivityByPartnerReportWednesdayTime
	Schedule_DailyInventoryAPIActivityByPartnerReportGenerationUpdated = Schedule_DailyInventoryAPIActivityByPartnerReportGenerationUpdated & "," & GenerateDailyInventoryAPIActivityByPartnerReportThursdayTime
	Schedule_DailyInventoryAPIActivityByPartnerReportGenerationUpdated = Schedule_DailyInventoryAPIActivityByPartnerReportGenerationUpdated & "," & GenerateDailyInventoryAPIActivityByPartnerReportFridayTime
	Schedule_DailyInventoryAPIActivityByPartnerReportGenerationUpdated = Schedule_DailyInventoryAPIActivityByPartnerReportGenerationUpdated & "," & GenerateDailyInventoryAPIActivityByPartnerReportSaturdayTime

	
	Schedule_DailyInventoryAPIActivityByPartnerReportGenerationUpdated = Schedule_DailyInventoryAPIActivityByPartnerReportGenerationUpdated & "," & RunDailyInventoryAPIActivityByPartnerReportIfClosed
	Schedule_DailyInventoryAPIActivityByPartnerReportGenerationUpdated = Schedule_DailyInventoryAPIActivityByPartnerReportGenerationUpdated & "," & RunDailyInventoryAPIActivityByPartnerReportIfClosingEarly
	
	Response.Write("<br><br><br>Schedule_DailyInventoryAPIActivityByPartnerReportGenerationUpdated: " & Schedule_DailyInventoryAPIActivityByPartnerReportGenerationUpdated)

	'*********************************************************************
	'Update SQL with Array/String of Schedule Data
	'*********************************************************************
		
	SQL = "UPDATE Settings_InventoryControl SET Schedule_DailyInventoryAPIActivityByPartnerReportGeneration = '" & cStr(Schedule_DailyInventoryAPIActivityByPartnerReportGenerationUpdated) & "' "
	
	Response.Write("<br><br><br>SQL: " & SQL)
	
	Set cnn8 = Server.CreateObject("ADODB.Connection")
	cnn8.open (Session("ClientCnnString"))
	Set rs = Server.CreateObject("ADODB.Recordset")
	rs.CursorLocation = 3 
	Set rs = cnn8.Execute(SQL)
	
	set rs = Nothing
	cnn8.close
	set cnn8 = Nothing


	 Response.Redirect("inventory.asp")
	
%><!--#include file="../../../inc/footer-main.asp"-->