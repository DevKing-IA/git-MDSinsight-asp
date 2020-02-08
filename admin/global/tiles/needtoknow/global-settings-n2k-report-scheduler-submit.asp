<!--#include file="../../../../inc/header.asp"-->

<%

	'***********************************************************
	'Get Values Of All Form Fields Posted 
	'***********************************************************
	
	GenerateGlobalSettingsNeedToKnowReportSunday = Request.Form("chkNoGlobalSettingsNeedToKnowReportSunday")
	GenerateGlobalSettingsNeedToKnowReportMonday = Request.Form("chkNoGlobalSettingsNeedToKnowReportMonday")
	GenerateGlobalSettingsNeedToKnowReportTuesday = Request.Form("chkNoGlobalSettingsNeedToKnowReportTuesday")
	GenerateGlobalSettingsNeedToKnowReportWednesday = Request.Form("chkNoGlobalSettingsNeedToKnowReportWednesday")
	GenerateGlobalSettingsNeedToKnowReportThursday = Request.Form("chkNoGlobalSettingsNeedToKnowReportThursday")
	GenerateGlobalSettingsNeedToKnowReportFriday = Request.Form("chkNoGlobalSettingsNeedToKnowReportFriday")
	GenerateGlobalSettingsNeedToKnowReportSaturday = Request.Form("chkNoGlobalSettingsNeedToKnowReportSaturday")
	
	GenerateGlobalSettingsNeedToKnowReportSundayTime = Request.Form("txtGlobalSettingsNeedToKnowReportSchedulerSundayTime")
	GenerateGlobalSettingsNeedToKnowReportMondayTime = Request.Form("txtGlobalSettingsNeedToKnowReportSchedulerMondayTime")
	GenerateGlobalSettingsNeedToKnowReportTuesdayTime = Request.Form("txtGlobalSettingsNeedToKnowReportSchedulerTuesdayTime")
	GenerateGlobalSettingsNeedToKnowReportWednesdayTime = Request.Form("txtGlobalSettingsNeedToKnowReportSchedulerWednesdayTime")
	GenerateGlobalSettingsNeedToKnowReportThursdayTime = Request.Form("txtGlobalSettingsNeedToKnowReportSchedulerThursdayTime")
	GenerateGlobalSettingsNeedToKnowReportFridayTime = Request.Form("txtGlobalSettingsNeedToKnowReportSchedulerFridayTime")
	GenerateGlobalSettingsNeedToKnowReportSaturdayTime = Request.Form("txtGlobalSettingsNeedToKnowReportSchedulerSaturdayTime")
	
	RunGlobalSettingsNeedToKnowReportIfClosed = Request.Form("chkNoGlobalSettingsNeedToKnowReportIfClosed")
	RunGlobalSettingsNeedToKnowReportIfClosingEarly = Request.Form("chkNoGlobalSettingsNeedToKnowReportIfClosingEarly")


	If Request.Form("chkNoGlobalSettingsNeedToKnowReportSunday") = "on" Then
		GenerateGlobalSettingsNeedToKnowReportSunday = 0
		GenerateGlobalSettingsNeedToKnowReportSundayTime = ""
	Else 
		GenerateGlobalSettingsNeedToKnowReportSunday = 1
	End If

	If Request.Form("chkNoGlobalSettingsNeedToKnowReportMonday") = "on" Then
		GenerateGlobalSettingsNeedToKnowReportMonday = 0
		GenerateGlobalSettingsNeedToKnowReportMondayTime = ""
	Else 
		GenerateGlobalSettingsNeedToKnowReportMonday = 1
	End If

	If Request.Form("chkNoGlobalSettingsNeedToKnowReportTuesday") = "on" Then
		GenerateGlobalSettingsNeedToKnowReportTuesday = 0
		GenerateGlobalSettingsNeedToKnowReportTuesdayTime = ""
	Else 
		GenerateGlobalSettingsNeedToKnowReportTuesday = 1
	End If

	If Request.Form("chkNoGlobalSettingsNeedToKnowReportWednesday") = "on" Then
		GenerateGlobalSettingsNeedToKnowReportWednesday = 0
		GenerateGlobalSettingsNeedToKnowReportWednesdayTime = ""
	Else 
		GenerateGlobalSettingsNeedToKnowReportWednesday = 1
	End If

	If Request.Form("chkNoGlobalSettingsNeedToKnowReportThursday") = "on" Then
		GenerateGlobalSettingsNeedToKnowReportThursday = 0
		GenerateGlobalSettingsNeedToKnowReportThursdayTime = ""
	Else 
		GenerateGlobalSettingsNeedToKnowReportThursday = 1
	End If

	If Request.Form("chkNoGlobalSettingsNeedToKnowReportFriday") = "on" Then
		GenerateGlobalSettingsNeedToKnowReportFriday = 0
		GenerateGlobalSettingsNeedToKnowReportFridayTime = ""
	Else 
		GenerateGlobalSettingsNeedToKnowReportFriday = 1
	End If

	If Request.Form("chkNoGlobalSettingsNeedToKnowReportSaturday") = "on" Then
		GenerateGlobalSettingsNeedToKnowReportSaturday = 0
		GenerateGlobalSettingsNeedToKnowReportSaturdayTime = ""
	Else 
		GenerateGlobalSettingsNeedToKnowReportSaturday = 1
	End If

	If Request.Form("chkNoGlobalSettingsNeedToKnowReportIfClosed") = "on" Then RunGlobalSettingsNeedToKnowReportIfClosed = 0 Else RunGlobalSettingsNeedToKnowReportIfClosed = 1
	If Request.Form("chkNoGlobalSettingsNeedToKnowReportIfClosingEarly") = "on" Then RunGlobalSettingsNeedToKnowReportIfClosingEarly = 0 Else RunGlobalSettingsNeedToKnowReportIfClosingEarly = 1
	
	'***********************************************************
	'Get Original Values For Audit Trail Entries
	'***********************************************************
	
	SQLPropsectingSettings = "SELECT * FROM Settings_NeedToKnow"
	
	Set cnnPropsectingSettings = Server.CreateObject("ADODB.Connection")
	cnnPropsectingSettings.open (Session("ClientCnnString"))
	Set rsPropsectingSettings = Server.CreateObject("ADODB.Recordset")
	rsPropsectingSettings.CursorLocation = 3 
	Set rsPropsectingSettings = cnnPropsectingSettings.Execute(SQLPropsectingSettings)
		
	If NOT rsPropsectingSettings.EOF Then
	
		Schedule_GlobalSettingsNeedToKnowReportGeneration = rsPropsectingSettings("Schedule_GlobalSettingsNeedToKnowReportGeneration")
		
		Schedule_GlobalSettingsNeedToKnowReportGenerationSettings = Split(Schedule_GlobalSettingsNeedToKnowReportGeneration,",")

		GenerateGlobalSettingsNeedToKnowReportSunday_ORIG = cInt(Schedule_GlobalSettingsNeedToKnowReportGenerationSettings(0))
		GenerateGlobalSettingsNeedToKnowReportMonday_ORIG = cInt(Schedule_GlobalSettingsNeedToKnowReportGenerationSettings(1))
		GenerateGlobalSettingsNeedToKnowReportTuesday_ORIG = cInt(Schedule_GlobalSettingsNeedToKnowReportGenerationSettings(2))
		GenerateGlobalSettingsNeedToKnowReportWednesday_ORIG = cInt(Schedule_GlobalSettingsNeedToKnowReportGenerationSettings(3))
		GenerateGlobalSettingsNeedToKnowReportThursday_ORIG = cInt(Schedule_GlobalSettingsNeedToKnowReportGenerationSettings(4))
		GenerateGlobalSettingsNeedToKnowReportFriday_ORIG = cInt(Schedule_GlobalSettingsNeedToKnowReportGenerationSettings(5))
		GenerateGlobalSettingsNeedToKnowReportSaturday_ORIG = cInt(Schedule_GlobalSettingsNeedToKnowReportGenerationSettings(6))
		GenerateGlobalSettingsNeedToKnowReportSundayTime_ORIG = Schedule_GlobalSettingsNeedToKnowReportGenerationSettings(7)
		GenerateGlobalSettingsNeedToKnowReportMondayTime_ORIG = Schedule_GlobalSettingsNeedToKnowReportGenerationSettings(8)
		GenerateGlobalSettingsNeedToKnowReportTuesdayTime_ORIG = Schedule_GlobalSettingsNeedToKnowReportGenerationSettings(9)
		GenerateGlobalSettingsNeedToKnowReportWednesdayTime_ORIG = Schedule_GlobalSettingsNeedToKnowReportGenerationSettings(10)
		GenerateGlobalSettingsNeedToKnowReportThursdayTime_ORIG = Schedule_GlobalSettingsNeedToKnowReportGenerationSettings(11)
		GenerateGlobalSettingsNeedToKnowReportFridayTime_ORIG = Schedule_GlobalSettingsNeedToKnowReportGenerationSettings(12)
		GenerateGlobalSettingsNeedToKnowReportSaturdayTime_ORIG = Schedule_GlobalSettingsNeedToKnowReportGenerationSettings(13)
		RunGlobalSettingsNeedToKnowReportIfClosed_ORIG = cInt(Schedule_GlobalSettingsNeedToKnowReportGenerationSettings(14))
		RunGlobalSettingsNeedToKnowReportIfClosingEarly_ORIG = cInt(Schedule_GlobalSettingsNeedToKnowReportGenerationSettings(15))
	
	End If
	
	set rsPropsectingSettings = Nothing
	cnnPropsectingSettings.close
	set cnnPropsectingSettings = Nothing
	

	'***********************************************************
	'Perform Audit Trail Entry Inserts
	'***********************************************************

	
	If Request.Form("chkNoGlobalSettingsNeedToKnowReportSunday") = "on" then GenerateGlobalSettingsNeedToKnowReportSundayMsg = "On" Else GenerateGlobalSettingsNeedToKnowReportSundayMsg = "Off"
	If GenerateGlobalSettingsNeedToKnowReportSunday_ORIG = 1 then GenerateGlobalSettingsNeedToKnowReportSundayMsgOrig = "On" Else GenerateGlobalSettingsNeedToKnowReportSundayMsgOrig = "Off"
	
	If GenerateGlobalSettingsNeedToKnowReportSunday <> GenerateGlobalSettingsNeedToKnowReportSunday_ORIG Then
		CreateAuditLogEntry "Global Settings Change", "Global Settings Change", "Major", 1, "Global Settings Need To Know Report generation schedule for SUNDAY changed from " & GenerateGlobalSettingsNeedToKnowReportSundayMsgOrig & " to " & GenerateGlobalSettingsNeedToKnowReportSundayMsg
	End If
	
	If GenerateGlobalSettingsNeedToKnowReportSundayTime <> GenerateGlobalSettingsNeedToKnowReportSundayTime_ORIG Then
		If GenerateGlobalSettingsNeedToKnowReportSunday_ORIG = 0 AND GenerateGlobalSettingsNeedToKnowReportSunday = 1 Then
			CreateAuditLogEntry "Global Settings Change", "Global Settings Change", "Major", 1, "Global Settings Need To Know Report generation schedule for SUNDAY turned on and set to run at " & GenerateGlobalSettingsNeedToKnowReportSundayTime
		Else
			CreateAuditLogEntry "Global Settings Change", "Global Settings Change", "Major", 1, "Global Settings Need To Know Report generation scheduled run time for SUNDAY changed from " & GenerateGlobalSettingsNeedToKnowReportSundayTime_ORIG & " to " & GenerateGlobalSettingsNeedToKnowReportSundayTime
		End If
	End If
	
	

	If Request.Form("chkNoGlobalSettingsNeedToKnowReportMonday") = "on" then GenerateGlobalSettingsNeedToKnowReportMondayMsg = "On" Else GenerateGlobalSettingsNeedToKnowReportMondayMsg = "Off"
	If GenerateGlobalSettingsNeedToKnowReportMonday_ORIG = 1 then GenerateGlobalSettingsNeedToKnowReportMondayMsgOrig = "On" Else GenerateGlobalSettingsNeedToKnowReportMondayMsgOrig = "Off"
	
	If GenerateGlobalSettingsNeedToKnowReportMonday <> GenerateGlobalSettingsNeedToKnowReportMonday_ORIG Then
		CreateAuditLogEntry "Global Settings Change", "Global Settings Change", "Major", 1, "Global Settings Need To Know Report generation schedule for Monday changed from " & GenerateGlobalSettingsNeedToKnowReportMondayMsgOrig & " to " & GenerateGlobalSettingsNeedToKnowReportMondayMsg
	End If
	
	If GenerateGlobalSettingsNeedToKnowReportMondayTime <> GenerateGlobalSettingsNeedToKnowReportMondayTime_ORIG Then
		If GenerateGlobalSettingsNeedToKnowReportMonday_ORIG = 0 AND GenerateGlobalSettingsNeedToKnowReportMonday = 1 Then
			CreateAuditLogEntry "Global Settings Change", "Global Settings Change", "Major", 1, "Global Settings Need To Know Report generation schedule for Monday turned on and set to run at " & GenerateGlobalSettingsNeedToKnowReportMondayTime
		Else
			CreateAuditLogEntry "Global Settings Change", "Global Settings Change", "Major", 1, "Global Settings Need To Know Report generation scheduled run time for Monday changed from " & GenerateGlobalSettingsNeedToKnowReportMondayTime_ORIG & " to " & GenerateGlobalSettingsNeedToKnowReportMondayTime
		End If
	End If
	


	If Request.Form("chkNoGlobalSettingsNeedToKnowReportTuesday") = "on" then GenerateGlobalSettingsNeedToKnowReportTuesdayMsg = "On" Else GenerateGlobalSettingsNeedToKnowReportTuesdayMsg = "Off"
	If GenerateGlobalSettingsNeedToKnowReportTuesday_ORIG = 1 then GenerateGlobalSettingsNeedToKnowReportTuesdayMsgOrig = "On" Else GenerateGlobalSettingsNeedToKnowReportTuesdayMsgOrig = "Off"
	
	If GenerateGlobalSettingsNeedToKnowReportTuesday <> GenerateGlobalSettingsNeedToKnowReportTuesday_ORIG Then
		CreateAuditLogEntry "Global Settings Change", "Global Settings Change", "Major", 1, "Global Settings Need To Know Report generation schedule for Tuesday changed from " & GenerateGlobalSettingsNeedToKnowReportTuesdayMsgOrig & " to " & GenerateGlobalSettingsNeedToKnowReportTuesdayMsg
	End If
	
	If GenerateGlobalSettingsNeedToKnowReportTuesdayTime <> GenerateGlobalSettingsNeedToKnowReportTuesdayTime_ORIG Then
		If GenerateGlobalSettingsNeedToKnowReportTuesday_ORIG = 0 AND GenerateGlobalSettingsNeedToKnowReportTuesday = 1 Then
			CreateAuditLogEntry "Global Settings Change", "Global Settings Change", "Major", 1, "Global Settings Need To Know Report generation schedule for Tuesday turned on and set to run at " & GenerateGlobalSettingsNeedToKnowReportTuesdayTime
		Else
			CreateAuditLogEntry "Global Settings Change", "Global Settings Change", "Major", 1, "Global Settings Need To Know Report generation scheduled run time for Tuesday changed from " & GenerateGlobalSettingsNeedToKnowReportTuesdayTime_ORIG & " to " & GenerateGlobalSettingsNeedToKnowReportTuesdayTime
		End If
	End If



	If Request.Form("chkNoGlobalSettingsNeedToKnowReportWednesday") = "on" then GenerateGlobalSettingsNeedToKnowReportWednesdayMsg = "On" Else GenerateGlobalSettingsNeedToKnowReportWednesdayMsg = "Off"
	If GenerateGlobalSettingsNeedToKnowReportWednesday_ORIG = 1 then GenerateGlobalSettingsNeedToKnowReportWednesdayMsgOrig = "On" Else GenerateGlobalSettingsNeedToKnowReportWednesdayMsgOrig = "Off"
	
	If GenerateGlobalSettingsNeedToKnowReportWednesday <> GenerateGlobalSettingsNeedToKnowReportWednesday_ORIG Then
		CreateAuditLogEntry "Global Settings Change", "Global Settings Change", "Major", 1, "Global Settings Need To Know Report generation schedule for Wednesday changed from " & GenerateGlobalSettingsNeedToKnowReportWednesdayMsgOrig & " to " & GenerateGlobalSettingsNeedToKnowReportWednesdayMsg
	End If
	
	If GenerateGlobalSettingsNeedToKnowReportWednesdayTime <> GenerateGlobalSettingsNeedToKnowReportWednesdayTime_ORIG Then
		If GenerateGlobalSettingsNeedToKnowReportWednesday_ORIG = 0 AND GenerateGlobalSettingsNeedToKnowReportWednesday = 1 Then
			CreateAuditLogEntry "Global Settings Change", "Global Settings Change", "Major", 1, "Global Settings Need To Know Report generation schedule for Wednesday turned on and set to run at " & GenerateGlobalSettingsNeedToKnowReportWednesdayTime
		Else
			CreateAuditLogEntry "Global Settings Change", "Global Settings Change", "Major", 1, "Global Settings Need To Know Report generation scheduled run time for Wednesday changed from " & GenerateGlobalSettingsNeedToKnowReportWednesdayTime_ORIG & " to " & GenerateGlobalSettingsNeedToKnowReportWednesdayTime
		End If
	End If



	If Request.Form("chkNoGlobalSettingsNeedToKnowReportThursday") = "on" then GenerateGlobalSettingsNeedToKnowReportThursdayMsg = "On" Else GenerateGlobalSettingsNeedToKnowReportThursdayMsg = "Off"
	If GenerateGlobalSettingsNeedToKnowReportThursday_ORIG = 1 then GenerateGlobalSettingsNeedToKnowReportThursdayMsgOrig = "On" Else GenerateGlobalSettingsNeedToKnowReportThursdayMsgOrig = "Off"
	
	If GenerateGlobalSettingsNeedToKnowReportThursday <> GenerateGlobalSettingsNeedToKnowReportThursday_ORIG Then
		CreateAuditLogEntry "Global Settings Change", "Global Settings Change", "Major", 1, "Global Settings Need To Know Report generation schedule for Thursday changed from " & GenerateGlobalSettingsNeedToKnowReportThursdayMsgOrig & " to " & GenerateGlobalSettingsNeedToKnowReportThursdayMsg
	End If
	
	If GenerateGlobalSettingsNeedToKnowReportThursdayTime <> GenerateGlobalSettingsNeedToKnowReportThursdayTime_ORIG Then
		If GenerateGlobalSettingsNeedToKnowReportThursday_ORIG = 0 AND GenerateGlobalSettingsNeedToKnowReportThursday = 1 Then
			CreateAuditLogEntry "Global Settings Change", "Global Settings Change", "Major", 1, "Global Settings Need To Know Report generation schedule for Thursday turned on and set to run at " & GenerateGlobalSettingsNeedToKnowReportThursdayTime
		Else
			CreateAuditLogEntry "Global Settings Change", "Global Settings Change", "Major", 1, "Global Settings Need To Know Report generation scheduled run time for Thursday changed from " & GenerateGlobalSettingsNeedToKnowReportThursdayTime_ORIG & " to " & GenerateGlobalSettingsNeedToKnowReportThursdayTime
		End If
	End If



	If Request.Form("chkNoGlobalSettingsNeedToKnowReportFriday") = "on" then GenerateGlobalSettingsNeedToKnowReportFridayMsg = "On" Else GenerateGlobalSettingsNeedToKnowReportFridayMsg = "Off"
	If GenerateGlobalSettingsNeedToKnowReportFriday_ORIG = 1 then GenerateGlobalSettingsNeedToKnowReportFridayMsgOrig = "On" Else GenerateGlobalSettingsNeedToKnowReportFridayMsgOrig = "Off"
	
	If GenerateGlobalSettingsNeedToKnowReportFriday <> GenerateGlobalSettingsNeedToKnowReportFriday_ORIG Then
		CreateAuditLogEntry "Global Settings Change", "Global Settings Change", "Major", 1, "Global Settings Need To Know Report generation schedule for Friday changed from " & GenerateGlobalSettingsNeedToKnowReportFridayMsgOrig & " to " & GenerateGlobalSettingsNeedToKnowReportFridayMsg
	End If
	
	If GenerateGlobalSettingsNeedToKnowReportFridayTime <> GenerateGlobalSettingsNeedToKnowReportFridayTime_ORIG Then
		If GenerateGlobalSettingsNeedToKnowReportFriday_ORIG = 0 AND GenerateGlobalSettingsNeedToKnowReportFriday = 1 Then
			CreateAuditLogEntry "Global Settings Change", "Global Settings Change", "Major", 1, "Global Settings Need To Know Report generation schedule for Friday turned on and set to run at " & GenerateGlobalSettingsNeedToKnowReportFridayTime
		Else
			CreateAuditLogEntry "Global Settings Change", "Global Settings Change", "Major", 1, "Global Settings Need To Know Report generation scheduled run time for Friday changed from " & GenerateGlobalSettingsNeedToKnowReportFridayTime_ORIG & " to " & GenerateGlobalSettingsNeedToKnowReportFridayTime
		End If
	End If



	If Request.Form("chkNoGlobalSettingsNeedToKnowReportSaturday") = "on" then GenerateGlobalSettingsNeedToKnowReportSaturdayMsg = "On" Else GenerateGlobalSettingsNeedToKnowReportSaturdayMsg = "Off"
	If GenerateGlobalSettingsNeedToKnowReportSaturday_ORIG = 1 then GenerateGlobalSettingsNeedToKnowReportSaturdayMsgOrig = "On" Else GenerateGlobalSettingsNeedToKnowReportSaturdayMsgOrig = "Off"
	
	If GenerateGlobalSettingsNeedToKnowReportSaturday <> GenerateGlobalSettingsNeedToKnowReportSaturday_ORIG Then
		CreateAuditLogEntry "Global Settings Change", "Global Settings Change", "Major", 1, "Global Settings Need To Know Report generation schedule for Saturday changed from " & GenerateGlobalSettingsNeedToKnowReportSaturdayMsgOrig & " to " & GenerateGlobalSettingsNeedToKnowReportSaturdayMsg
	End If
	
	If GenerateGlobalSettingsNeedToKnowReportSaturdayTime <> GenerateGlobalSettingsNeedToKnowReportSaturdayTime_ORIG Then
		If GenerateGlobalSettingsNeedToKnowReportSaturday_ORIG = 0 AND GenerateGlobalSettingsNeedToKnowReportSaturday = 1 Then
			CreateAuditLogEntry "Global Settings Change", "Global Settings Change", "Major", 1, "Global Settings Need To Know Report generation schedule for Saturday turned on and set to run at " & GenerateGlobalSettingsNeedToKnowReportSaturdayTime
		Else
			CreateAuditLogEntry "Global Settings Change", "Global Settings Change", "Major", 1, "Global Settings Need To Know Report generation scheduled run time for Saturday changed from " & GenerateGlobalSettingsNeedToKnowReportSaturdayTime_ORIG & " to " & GenerateGlobalSettingsNeedToKnowReportSaturdayTime
		End If
	End If


	If Request.Form("chkNoGlobalSettingsNeedToKnowReportIfClosed") = "on" then RunGlobalSettingsNeedToKnowReportIfClosedMsg = "On" Else RunGlobalSettingsNeedToKnowReportIfClosedMsg = "Off"
	If RunGlobalSettingsNeedToKnowReportIfClosed_ORIG = 1 then RunGlobalSettingsNeedToKnowReportIfClosedMsgOrig = "On" Else RunGlobalSettingsNeedToKnowReportIfClosedMsgOrig = "Off"
	
	If RunGlobalSettingsNeedToKnowReportIfClosed <> RunGlobalSettingsNeedToKnowReportIfClosed_ORIG Then
		CreateAuditLogEntry "Global Settings Change", "Global Settings Change", "Major", 1, "Global Settings Need To Know Report generation schedule - Don''t Run If Closed (Monday-Friday Only) -  changed from " & RunGlobalSettingsNeedToKnowReportIfClosedMsgOrig & " to " & RunGlobalSettingsNeedToKnowReportIfClosedMsg
	End If


	If Request.Form("chkNoGlobalSettingsNeedToKnowReportIfClosingEarly") = "on" then RunGlobalSettingsNeedToKnowReportIfClosingEarlyMsg = "On" Else RunGlobalSettingsNeedToKnowReportIfClosingEarlyMsg = "Off"
	If RunGlobalSettingsNeedToKnowReportIfClosingEarly_ORIG = 1 then RunGlobalSettingsNeedToKnowReportIfClosingEarlyMsgOrig = "On" Else RunGlobalSettingsNeedToKnowReportIfClosingEarlyMsgOrig = "Off"
	
	If RunGlobalSettingsNeedToKnowReportIfClosingEarly <> RunGlobalSettingsNeedToKnowReportIfClosingEarly_ORIG Then
		CreateAuditLogEntry "Global Settings Change", "Global Settings Change", "Major", 1, "Global Settings Need To Know Report generation schedule - Don''t Run If Closed (Monday-Friday Only) -  changed from " & RunGlobalSettingsNeedToKnowReportIfClosingEarlyMsgOrig & " to " & RunGlobalSettingsNeedToKnowReportIfClosingEarlyMsg
	End If


	'*********************************************************************
	'Build Array/String of Schedule Data From Request Form Field Data
	'*********************************************************************

	Schedule_GlobalSettingsNeedToKnowReportGenerationUpdated = ""
	
	Schedule_GlobalSettingsNeedToKnowReportGenerationUpdated = GenerateGlobalSettingsNeedToKnowReportSunday
	Schedule_GlobalSettingsNeedToKnowReportGenerationUpdated = Schedule_GlobalSettingsNeedToKnowReportGenerationUpdated & "," & GenerateGlobalSettingsNeedToKnowReportMonday
	Schedule_GlobalSettingsNeedToKnowReportGenerationUpdated = Schedule_GlobalSettingsNeedToKnowReportGenerationUpdated & "," & GenerateGlobalSettingsNeedToKnowReportTuesday
	Schedule_GlobalSettingsNeedToKnowReportGenerationUpdated = Schedule_GlobalSettingsNeedToKnowReportGenerationUpdated & "," & GenerateGlobalSettingsNeedToKnowReportWednesday
	Schedule_GlobalSettingsNeedToKnowReportGenerationUpdated = Schedule_GlobalSettingsNeedToKnowReportGenerationUpdated & "," & GenerateGlobalSettingsNeedToKnowReportThursday
	Schedule_GlobalSettingsNeedToKnowReportGenerationUpdated = Schedule_GlobalSettingsNeedToKnowReportGenerationUpdated & "," & GenerateGlobalSettingsNeedToKnowReportFriday
	Schedule_GlobalSettingsNeedToKnowReportGenerationUpdated = Schedule_GlobalSettingsNeedToKnowReportGenerationUpdated & "," & GenerateGlobalSettingsNeedToKnowReportSaturday
	
	Schedule_GlobalSettingsNeedToKnowReportGenerationUpdated = Schedule_GlobalSettingsNeedToKnowReportGenerationUpdated & "," & GenerateGlobalSettingsNeedToKnowReportSundayTime
	Schedule_GlobalSettingsNeedToKnowReportGenerationUpdated = Schedule_GlobalSettingsNeedToKnowReportGenerationUpdated & "," & GenerateGlobalSettingsNeedToKnowReportMondayTime
	Schedule_GlobalSettingsNeedToKnowReportGenerationUpdated = Schedule_GlobalSettingsNeedToKnowReportGenerationUpdated & "," & GenerateGlobalSettingsNeedToKnowReportTuesdayTime
	Schedule_GlobalSettingsNeedToKnowReportGenerationUpdated = Schedule_GlobalSettingsNeedToKnowReportGenerationUpdated & "," & GenerateGlobalSettingsNeedToKnowReportWednesdayTime
	Schedule_GlobalSettingsNeedToKnowReportGenerationUpdated = Schedule_GlobalSettingsNeedToKnowReportGenerationUpdated & "," & GenerateGlobalSettingsNeedToKnowReportThursdayTime
	Schedule_GlobalSettingsNeedToKnowReportGenerationUpdated = Schedule_GlobalSettingsNeedToKnowReportGenerationUpdated & "," & GenerateGlobalSettingsNeedToKnowReportFridayTime
	Schedule_GlobalSettingsNeedToKnowReportGenerationUpdated = Schedule_GlobalSettingsNeedToKnowReportGenerationUpdated & "," & GenerateGlobalSettingsNeedToKnowReportSaturdayTime

	
	Schedule_GlobalSettingsNeedToKnowReportGenerationUpdated = Schedule_GlobalSettingsNeedToKnowReportGenerationUpdated & "," & RunGlobalSettingsNeedToKnowReportIfClosed
	Schedule_GlobalSettingsNeedToKnowReportGenerationUpdated = Schedule_GlobalSettingsNeedToKnowReportGenerationUpdated & "," & RunGlobalSettingsNeedToKnowReportIfClosingEarly
	
	Response.Write("<br><br><br>Schedule_GlobalSettingsNeedToKnowReportGenerationUpdated: " & Schedule_GlobalSettingsNeedToKnowReportGenerationUpdated)

	'*********************************************************************
	'Update SQL with Array/String of Schedule Data
	'*********************************************************************
		
	SQL = "UPDATE Settings_NeedToKnow SET Schedule_GlobalSettingsNeedToKnowReportGeneration = '" & cStr(Schedule_GlobalSettingsNeedToKnowReportGenerationUpdated) & "' "
	
	Response.Write("<br><br><br>SQL: " & SQL)
	
	Set cnn8 = Server.CreateObject("ADODB.Connection")
	cnn8.open (Session("ClientCnnString"))
	Set rs = Server.CreateObject("ADODB.Recordset")
	rs.CursorLocation = 3 
	Set rs = cnn8.Execute(SQL)
	
	set rs = Nothing
	cnn8.close
	set cnn8 = Nothing


	 Response.Redirect("global-settings.asp")
	
%><!--#include file="../../../../inc/footer-main.asp"-->