<!--#include file="../../../inc/header.asp"-->

<%

	'***********************************************************
	'Get Values Of All Form Fields Posted 
	'***********************************************************
	
	GenerateProspectingWeeklyAgendaReportSunday = Request.Form("chkNoProspectingWeeklyAgendaReportSunday")
	GenerateProspectingWeeklyAgendaReportMonday = Request.Form("chkNoProspectingWeeklyAgendaReportMonday")
	GenerateProspectingWeeklyAgendaReportTuesday = Request.Form("chkNoProspectingWeeklyAgendaReportTuesday")
	GenerateProspectingWeeklyAgendaReportWednesday = Request.Form("chkNoProspectingWeeklyAgendaReportWednesday")
	GenerateProspectingWeeklyAgendaReportThursday = Request.Form("chkNoProspectingWeeklyAgendaReportThursday")
	GenerateProspectingWeeklyAgendaReportFriday = Request.Form("chkNoProspectingWeeklyAgendaReportFriday")
	GenerateProspectingWeeklyAgendaReportSaturday = Request.Form("chkNoProspectingWeeklyAgendaReportSaturday")
	
	GenerateProspectingWeeklyAgendaReportSundayTime = Request.Form("txtProspectingWeeklyAgendaReportSchedulerSundayTime")
	GenerateProspectingWeeklyAgendaReportMondayTime = Request.Form("txtProspectingWeeklyAgendaReportSchedulerMondayTime")
	GenerateProspectingWeeklyAgendaReportTuesdayTime = Request.Form("txtProspectingWeeklyAgendaReportSchedulerTuesdayTime")
	GenerateProspectingWeeklyAgendaReportWednesdayTime = Request.Form("txtProspectingWeeklyAgendaReportSchedulerWednesdayTime")
	GenerateProspectingWeeklyAgendaReportThursdayTime = Request.Form("txtProspectingWeeklyAgendaReportSchedulerThursdayTime")
	GenerateProspectingWeeklyAgendaReportFridayTime = Request.Form("txtProspectingWeeklyAgendaReportSchedulerFridayTime")
	GenerateProspectingWeeklyAgendaReportSaturdayTime = Request.Form("txtProspectingWeeklyAgendaReportSchedulerSaturdayTime")
	
	RunProspectingWeeklyAgendaReportIfClosed = Request.Form("chkNoProspectingWeeklyAgendaReportIfClosed")
	RunProspectingWeeklyAgendaReportIfClosingEarly = Request.Form("chkNoProspectingWeeklyAgendaReportIfClosingEarly")


	If Request.Form("chkNoProspectingWeeklyAgendaReportSunday") = "on" Then
		GenerateProspectingWeeklyAgendaReportSunday = 0
		GenerateProspectingWeeklyAgendaReportSundayTime = ""
	Else 
		GenerateProspectingWeeklyAgendaReportSunday = 1
	End If

	If Request.Form("chkNoProspectingWeeklyAgendaReportMonday") = "on" Then
		GenerateProspectingWeeklyAgendaReportMonday = 0
		GenerateProspectingWeeklyAgendaReportMondayTime = ""
	Else 
		GenerateProspectingWeeklyAgendaReportMonday = 1
	End If

	If Request.Form("chkNoProspectingWeeklyAgendaReportTuesday") = "on" Then
		GenerateProspectingWeeklyAgendaReportTuesday = 0
		GenerateProspectingWeeklyAgendaReportTuesdayTime = ""
	Else 
		GenerateProspectingWeeklyAgendaReportTuesday = 1
	End If

	If Request.Form("chkNoProspectingWeeklyAgendaReportWednesday") = "on" Then
		GenerateProspectingWeeklyAgendaReportWednesday = 0
		GenerateProspectingWeeklyAgendaReportWednesdayTime = ""
	Else 
		GenerateProspectingWeeklyAgendaReportWednesday = 1
	End If

	If Request.Form("chkNoProspectingWeeklyAgendaReportThursday") = "on" Then
		GenerateProspectingWeeklyAgendaReportThursday = 0
		GenerateProspectingWeeklyAgendaReportThursdayTime = ""
	Else 
		GenerateProspectingWeeklyAgendaReportThursday = 1
	End If

	If Request.Form("chkNoProspectingWeeklyAgendaReportFriday") = "on" Then
		GenerateProspectingWeeklyAgendaReportFriday = 0
		GenerateProspectingWeeklyAgendaReportFridayTime = ""
	Else 
		GenerateProspectingWeeklyAgendaReportFriday = 1
	End If

	If Request.Form("chkNoProspectingWeeklyAgendaReportSaturday") = "on" Then
		GenerateProspectingWeeklyAgendaReportSaturday = 0
		GenerateProspectingWeeklyAgendaReportSaturdayTime = ""
	Else 
		GenerateProspectingWeeklyAgendaReportSaturday = 1
	End If

	If Request.Form("chkNoProspectingWeeklyAgendaReportIfClosed") = "on" Then RunProspectingWeeklyAgendaReportIfClosed = 0 Else RunProspectingWeeklyAgendaReportIfClosed = 1
	If Request.Form("chkNoProspectingWeeklyAgendaReportIfClosingEarly") = "on" Then RunProspectingWeeklyAgendaReportIfClosingEarly = 0 Else RunProspectingWeeklyAgendaReportIfClosingEarly = 1
	
	'***********************************************************
	'Get Original Values For Audit Trail Entries
	'***********************************************************
	
	SQLProspectingSettings = "SELECT * FROM Settings_Prospecting"
	
	Set cnnProspectingSettings = Server.CreateObject("ADODB.Connection")
	cnnProspectingSettings.open (Session("ClientCnnString"))
	Set rsProspectingSettings = Server.CreateObject("ADODB.Recordset")
	rsProspectingSettings.CursorLocation = 3 
	Set rsProspectingSettings = cnnProspectingSettings.Execute(SQLProspectingSettings)
		
	If NOT rsProspectingSettings.EOF Then
	
		Schedule_ProspectingWeeklyAgendaReportGeneration = rsProspectingSettings("Schedule_ProspectingWeeklyAgendaReportGeneration")
		
		Schedule_ProspectingWeeklyAgendaReportGenerationSettings = Split(Schedule_ProspectingWeeklyAgendaReportGeneration,",")

		GenerateProspectingWeeklyAgendaReportSunday_ORIG = cInt(Schedule_ProspectingWeeklyAgendaReportGenerationSettings(0))
		GenerateProspectingWeeklyAgendaReportMonday_ORIG = cInt(Schedule_ProspectingWeeklyAgendaReportGenerationSettings(1))
		GenerateProspectingWeeklyAgendaReportTuesday_ORIG = cInt(Schedule_ProspectingWeeklyAgendaReportGenerationSettings(2))
		GenerateProspectingWeeklyAgendaReportWednesday_ORIG = cInt(Schedule_ProspectingWeeklyAgendaReportGenerationSettings(3))
		GenerateProspectingWeeklyAgendaReportThursday_ORIG = cInt(Schedule_ProspectingWeeklyAgendaReportGenerationSettings(4))
		GenerateProspectingWeeklyAgendaReportFriday_ORIG = cInt(Schedule_ProspectingWeeklyAgendaReportGenerationSettings(5))
		GenerateProspectingWeeklyAgendaReportSaturday_ORIG = cInt(Schedule_ProspectingWeeklyAgendaReportGenerationSettings(6))
		GenerateProspectingWeeklyAgendaReportSundayTime_ORIG = Schedule_ProspectingWeeklyAgendaReportGenerationSettings(7)
		GenerateProspectingWeeklyAgendaReportMondayTime_ORIG = Schedule_ProspectingWeeklyAgendaReportGenerationSettings(8)
		GenerateProspectingWeeklyAgendaReportTuesdayTime_ORIG = Schedule_ProspectingWeeklyAgendaReportGenerationSettings(9)
		GenerateProspectingWeeklyAgendaReportWednesdayTime_ORIG = Schedule_ProspectingWeeklyAgendaReportGenerationSettings(10)
		GenerateProspectingWeeklyAgendaReportThursdayTime_ORIG = Schedule_ProspectingWeeklyAgendaReportGenerationSettings(11)
		GenerateProspectingWeeklyAgendaReportFridayTime_ORIG = Schedule_ProspectingWeeklyAgendaReportGenerationSettings(12)
		GenerateProspectingWeeklyAgendaReportSaturdayTime_ORIG = Schedule_ProspectingWeeklyAgendaReportGenerationSettings(13)
		RunProspectingWeeklyAgendaReportIfClosed_ORIG = cInt(Schedule_ProspectingWeeklyAgendaReportGenerationSettings(14))
		RunProspectingWeeklyAgendaReportIfClosingEarly_ORIG = cInt(Schedule_ProspectingWeeklyAgendaReportGenerationSettings(15))
	
	End If
	
	set rsProspectingSettings = Nothing
	cnnProspectingSettings.close
	set cnnProspectingSettings = Nothing
	

	'***********************************************************
	'Perform Audit Trail Entry Inserts
	'***********************************************************

	
	If Request.Form("chkNoProspectingWeeklyAgendaReportSunday") = "on" then GenerateProspectingWeeklyAgendaReportSundayMsg = "On" Else GenerateProspectingWeeklyAgendaReportSundayMsg = "Off"
	If GenerateProspectingWeeklyAgendaReportSunday_ORIG = 1 then GenerateProspectingWeeklyAgendaReportSundayMsgOrig = "On" Else GenerateProspectingWeeklyAgendaReportSundayMsgOrig = "Off"
	
	If GenerateProspectingWeeklyAgendaReportSunday <> GenerateProspectingWeeklyAgendaReportSunday_ORIG Then
		CreateAuditLogEntry "Global Settings Change", "Global Settings Change", "Major", 1, "Prospecting Snapshot Report generation schedule for SUNDAY changed from " & GenerateProspectingWeeklyAgendaReportSundayMsgOrig & " to " & GenerateProspectingWeeklyAgendaReportSundayMsg
	End If
	
	If GenerateProspectingWeeklyAgendaReportSundayTime <> GenerateProspectingWeeklyAgendaReportSundayTime_ORIG Then
		If GenerateProspectingWeeklyAgendaReportSunday_ORIG = 0 AND GenerateProspectingWeeklyAgendaReportSunday = 1 Then
			CreateAuditLogEntry "Global Settings Change", "Global Settings Change", "Major", 1, "Prospecting Snapshot Report generation schedule for SUNDAY turned on and set to run at " & GenerateProspectingWeeklyAgendaReportSundayTime
		Else
			CreateAuditLogEntry "Global Settings Change", "Global Settings Change", "Major", 1, "Prospecting Snapshot Report generation scheduled run time for SUNDAY changed from " & GenerateProspectingWeeklyAgendaReportSundayTime_ORIG & " to " & GenerateProspectingWeeklyAgendaReportSundayTime
		End If
	End If
	
	

	If Request.Form("chkNoProspectingWeeklyAgendaReportMonday") = "on" then GenerateProspectingWeeklyAgendaReportMondayMsg = "On" Else GenerateProspectingWeeklyAgendaReportMondayMsg = "Off"
	If GenerateProspectingWeeklyAgendaReportMonday_ORIG = 1 then GenerateProspectingWeeklyAgendaReportMondayMsgOrig = "On" Else GenerateProspectingWeeklyAgendaReportMondayMsgOrig = "Off"
	
	If GenerateProspectingWeeklyAgendaReportMonday <> GenerateProspectingWeeklyAgendaReportMonday_ORIG Then
		CreateAuditLogEntry "Global Settings Change", "Global Settings Change", "Major", 1, "Prospecting Snapshot Report generation schedule for Monday changed from " & GenerateProspectingWeeklyAgendaReportMondayMsgOrig & " to " & GenerateProspectingWeeklyAgendaReportMondayMsg
	End If
	
	If GenerateProspectingWeeklyAgendaReportMondayTime <> GenerateProspectingWeeklyAgendaReportMondayTime_ORIG Then
		If GenerateProspectingWeeklyAgendaReportMonday_ORIG = 0 AND GenerateProspectingWeeklyAgendaReportMonday = 1 Then
			CreateAuditLogEntry "Global Settings Change", "Global Settings Change", "Major", 1, "Prospecting Snapshot Report generation schedule for Monday turned on and set to run at " & GenerateProspectingWeeklyAgendaReportMondayTime
		Else
			CreateAuditLogEntry "Global Settings Change", "Global Settings Change", "Major", 1, "Prospecting Snapshot Report generation scheduled run time for Monday changed from " & GenerateProspectingWeeklyAgendaReportMondayTime_ORIG & " to " & GenerateProspectingWeeklyAgendaReportMondayTime
		End If
	End If
	


	If Request.Form("chkNoProspectingWeeklyAgendaReportTuesday") = "on" then GenerateProspectingWeeklyAgendaReportTuesdayMsg = "On" Else GenerateProspectingWeeklyAgendaReportTuesdayMsg = "Off"
	If GenerateProspectingWeeklyAgendaReportTuesday_ORIG = 1 then GenerateProspectingWeeklyAgendaReportTuesdayMsgOrig = "On" Else GenerateProspectingWeeklyAgendaReportTuesdayMsgOrig = "Off"
	
	If GenerateProspectingWeeklyAgendaReportTuesday <> GenerateProspectingWeeklyAgendaReportTuesday_ORIG Then
		CreateAuditLogEntry "Global Settings Change", "Global Settings Change", "Major", 1, "Prospecting Snapshot Report generation schedule for Tuesday changed from " & GenerateProspectingWeeklyAgendaReportTuesdayMsgOrig & " to " & GenerateProspectingWeeklyAgendaReportTuesdayMsg
	End If
	
	If GenerateProspectingWeeklyAgendaReportTuesdayTime <> GenerateProspectingWeeklyAgendaReportTuesdayTime_ORIG Then
		If GenerateProspectingWeeklyAgendaReportTuesday_ORIG = 0 AND GenerateProspectingWeeklyAgendaReportTuesday = 1 Then
			CreateAuditLogEntry "Global Settings Change", "Global Settings Change", "Major", 1, "Prospecting Snapshot Report generation schedule for Tuesday turned on and set to run at " & GenerateProspectingWeeklyAgendaReportTuesdayTime
		Else
			CreateAuditLogEntry "Global Settings Change", "Global Settings Change", "Major", 1, "Prospecting Snapshot Report generation scheduled run time for Tuesday changed from " & GenerateProspectingWeeklyAgendaReportTuesdayTime_ORIG & " to " & GenerateProspectingWeeklyAgendaReportTuesdayTime
		End If
	End If



	If Request.Form("chkNoProspectingWeeklyAgendaReportWednesday") = "on" then GenerateProspectingWeeklyAgendaReportWednesdayMsg = "On" Else GenerateProspectingWeeklyAgendaReportWednesdayMsg = "Off"
	If GenerateProspectingWeeklyAgendaReportWednesday_ORIG = 1 then GenerateProspectingWeeklyAgendaReportWednesdayMsgOrig = "On" Else GenerateProspectingWeeklyAgendaReportWednesdayMsgOrig = "Off"
	
	If GenerateProspectingWeeklyAgendaReportWednesday <> GenerateProspectingWeeklyAgendaReportWednesday_ORIG Then
		CreateAuditLogEntry "Global Settings Change", "Global Settings Change", "Major", 1, "Prospecting Snapshot Report generation schedule for Wednesday changed from " & GenerateProspectingWeeklyAgendaReportWednesdayMsgOrig & " to " & GenerateProspectingWeeklyAgendaReportWednesdayMsg
	End If
	
	If GenerateProspectingWeeklyAgendaReportWednesdayTime <> GenerateProspectingWeeklyAgendaReportWednesdayTime_ORIG Then
		If GenerateProspectingWeeklyAgendaReportWednesday_ORIG = 0 AND GenerateProspectingWeeklyAgendaReportWednesday = 1 Then
			CreateAuditLogEntry "Global Settings Change", "Global Settings Change", "Major", 1, "Prospecting Snapshot Report generation schedule for Wednesday turned on and set to run at " & GenerateProspectingWeeklyAgendaReportWednesdayTime
		Else
			CreateAuditLogEntry "Global Settings Change", "Global Settings Change", "Major", 1, "Prospecting Snapshot Report generation scheduled run time for Wednesday changed from " & GenerateProspectingWeeklyAgendaReportWednesdayTime_ORIG & " to " & GenerateProspectingWeeklyAgendaReportWednesdayTime
		End If
	End If



	If Request.Form("chkNoProspectingWeeklyAgendaReportThursday") = "on" then GenerateProspectingWeeklyAgendaReportThursdayMsg = "On" Else GenerateProspectingWeeklyAgendaReportThursdayMsg = "Off"
	If GenerateProspectingWeeklyAgendaReportThursday_ORIG = 1 then GenerateProspectingWeeklyAgendaReportThursdayMsgOrig = "On" Else GenerateProspectingWeeklyAgendaReportThursdayMsgOrig = "Off"
	
	If GenerateProspectingWeeklyAgendaReportThursday <> GenerateProspectingWeeklyAgendaReportThursday_ORIG Then
		CreateAuditLogEntry "Global Settings Change", "Global Settings Change", "Major", 1, "Prospecting Snapshot Report generation schedule for Thursday changed from " & GenerateProspectingWeeklyAgendaReportThursdayMsgOrig & " to " & GenerateProspectingWeeklyAgendaReportThursdayMsg
	End If
	
	If GenerateProspectingWeeklyAgendaReportThursdayTime <> GenerateProspectingWeeklyAgendaReportThursdayTime_ORIG Then
		If GenerateProspectingWeeklyAgendaReportThursday_ORIG = 0 AND GenerateProspectingWeeklyAgendaReportThursday = 1 Then
			CreateAuditLogEntry "Global Settings Change", "Global Settings Change", "Major", 1, "Prospecting Snapshot Report generation schedule for Thursday turned on and set to run at " & GenerateProspectingWeeklyAgendaReportThursdayTime
		Else
			CreateAuditLogEntry "Global Settings Change", "Global Settings Change", "Major", 1, "Prospecting Snapshot Report generation scheduled run time for Thursday changed from " & GenerateProspectingWeeklyAgendaReportThursdayTime_ORIG & " to " & GenerateProspectingWeeklyAgendaReportThursdayTime
		End If
	End If



	If Request.Form("chkNoProspectingWeeklyAgendaReportFriday") = "on" then GenerateProspectingWeeklyAgendaReportFridayMsg = "On" Else GenerateProspectingWeeklyAgendaReportFridayMsg = "Off"
	If GenerateProspectingWeeklyAgendaReportFriday_ORIG = 1 then GenerateProspectingWeeklyAgendaReportFridayMsgOrig = "On" Else GenerateProspectingWeeklyAgendaReportFridayMsgOrig = "Off"
	
	If GenerateProspectingWeeklyAgendaReportFriday <> GenerateProspectingWeeklyAgendaReportFriday_ORIG Then
		CreateAuditLogEntry "Global Settings Change", "Global Settings Change", "Major", 1, "Prospecting Snapshot Report generation schedule for Friday changed from " & GenerateProspectingWeeklyAgendaReportFridayMsgOrig & " to " & GenerateProspectingWeeklyAgendaReportFridayMsg
	End If
	
	If GenerateProspectingWeeklyAgendaReportFridayTime <> GenerateProspectingWeeklyAgendaReportFridayTime_ORIG Then
		If GenerateProspectingWeeklyAgendaReportFriday_ORIG = 0 AND GenerateProspectingWeeklyAgendaReportFriday = 1 Then
			CreateAuditLogEntry "Global Settings Change", "Global Settings Change", "Major", 1, "Prospecting Snapshot Report generation schedule for Friday turned on and set to run at " & GenerateProspectingWeeklyAgendaReportFridayTime
		Else
			CreateAuditLogEntry "Global Settings Change", "Global Settings Change", "Major", 1, "Prospecting Snapshot Report generation scheduled run time for Friday changed from " & GenerateProspectingWeeklyAgendaReportFridayTime_ORIG & " to " & GenerateProspectingWeeklyAgendaReportFridayTime
		End If
	End If



	If Request.Form("chkNoProspectingWeeklyAgendaReportSaturday") = "on" then GenerateProspectingWeeklyAgendaReportSaturdayMsg = "On" Else GenerateProspectingWeeklyAgendaReportSaturdayMsg = "Off"
	If GenerateProspectingWeeklyAgendaReportSaturday_ORIG = 1 then GenerateProspectingWeeklyAgendaReportSaturdayMsgOrig = "On" Else GenerateProspectingWeeklyAgendaReportSaturdayMsgOrig = "Off"
	
	If GenerateProspectingWeeklyAgendaReportSaturday <> GenerateProspectingWeeklyAgendaReportSaturday_ORIG Then
		CreateAuditLogEntry "Global Settings Change", "Global Settings Change", "Major", 1, "Prospecting Snapshot Report generation schedule for Saturday changed from " & GenerateProspectingWeeklyAgendaReportSaturdayMsgOrig & " to " & GenerateProspectingWeeklyAgendaReportSaturdayMsg
	End If
	
	If GenerateProspectingWeeklyAgendaReportSaturdayTime <> GenerateProspectingWeeklyAgendaReportSaturdayTime_ORIG Then
		If GenerateProspectingWeeklyAgendaReportSaturday_ORIG = 0 AND GenerateProspectingWeeklyAgendaReportSaturday = 1 Then
			CreateAuditLogEntry "Global Settings Change", "Global Settings Change", "Major", 1, "Prospecting Snapshot Report generation schedule for Saturday turned on and set to run at " & GenerateProspectingWeeklyAgendaReportSaturdayTime
		Else
			CreateAuditLogEntry "Global Settings Change", "Global Settings Change", "Major", 1, "Prospecting Snapshot Report generation scheduled run time for Saturday changed from " & GenerateProspectingWeeklyAgendaReportSaturdayTime_ORIG & " to " & GenerateProspectingWeeklyAgendaReportSaturdayTime
		End If
	End If


	If Request.Form("chkNoProspectingWeeklyAgendaReportIfClosed") = "on" then RunProspectingWeeklyAgendaReportIfClosedMsg = "On" Else RunProspectingWeeklyAgendaReportIfClosedMsg = "Off"
	If RunProspectingWeeklyAgendaReportIfClosed_ORIG = 1 then RunProspectingWeeklyAgendaReportIfClosedMsgOrig = "On" Else RunProspectingWeeklyAgendaReportIfClosedMsgOrig = "Off"
	
	If RunProspectingWeeklyAgendaReportIfClosed <> RunProspectingWeeklyAgendaReportIfClosed_ORIG Then
		CreateAuditLogEntry "Global Settings Change", "Global Settings Change", "Major", 1, "Prospecting Snapshot Report generation schedule - Don''t Run If Closed (Monday-Friday Only) -  changed from " & RunProspectingWeeklyAgendaReportIfClosedMsgOrig & " to " & RunProspectingWeeklyAgendaReportIfClosedMsg
	End If


	If Request.Form("chkNoProspectingWeeklyAgendaReportIfClosingEarly") = "on" then RunProspectingWeeklyAgendaReportIfClosingEarlyMsg = "On" Else RunProspectingWeeklyAgendaReportIfClosingEarlyMsg = "Off"
	If RunProspectingWeeklyAgendaReportIfClosingEarly_ORIG = 1 then RunProspectingWeeklyAgendaReportIfClosingEarlyMsgOrig = "On" Else RunProspectingWeeklyAgendaReportIfClosingEarlyMsgOrig = "Off"
	
	If RunProspectingWeeklyAgendaReportIfClosingEarly <> RunProspectingWeeklyAgendaReportIfClosingEarly_ORIG Then
		CreateAuditLogEntry "Global Settings Change", "Global Settings Change", "Major", 1, "Prospecting Snapshot Report generation schedule - Don''t Run If Closed (Monday-Friday Only) -  changed from " & RunProspectingWeeklyAgendaReportIfClosingEarlyMsgOrig & " to " & RunProspectingWeeklyAgendaReportIfClosingEarlyMsg
	End If


	'*********************************************************************
	'Build Array/String of Schedule Data From Request Form Field Data
	'*********************************************************************

	Schedule_ProspectingWeeklyAgendaReportGenerationUpdated = ""
	
	Schedule_ProspectingWeeklyAgendaReportGenerationUpdated = GenerateProspectingWeeklyAgendaReportSunday
	Schedule_ProspectingWeeklyAgendaReportGenerationUpdated = Schedule_ProspectingWeeklyAgendaReportGenerationUpdated & "," & GenerateProspectingWeeklyAgendaReportMonday
	Schedule_ProspectingWeeklyAgendaReportGenerationUpdated = Schedule_ProspectingWeeklyAgendaReportGenerationUpdated & "," & GenerateProspectingWeeklyAgendaReportTuesday
	Schedule_ProspectingWeeklyAgendaReportGenerationUpdated = Schedule_ProspectingWeeklyAgendaReportGenerationUpdated & "," & GenerateProspectingWeeklyAgendaReportWednesday
	Schedule_ProspectingWeeklyAgendaReportGenerationUpdated = Schedule_ProspectingWeeklyAgendaReportGenerationUpdated & "," & GenerateProspectingWeeklyAgendaReportThursday
	Schedule_ProspectingWeeklyAgendaReportGenerationUpdated = Schedule_ProspectingWeeklyAgendaReportGenerationUpdated & "," & GenerateProspectingWeeklyAgendaReportFriday
	Schedule_ProspectingWeeklyAgendaReportGenerationUpdated = Schedule_ProspectingWeeklyAgendaReportGenerationUpdated & "," & GenerateProspectingWeeklyAgendaReportSaturday
	
	Schedule_ProspectingWeeklyAgendaReportGenerationUpdated = Schedule_ProspectingWeeklyAgendaReportGenerationUpdated & "," & GenerateProspectingWeeklyAgendaReportSundayTime
	Schedule_ProspectingWeeklyAgendaReportGenerationUpdated = Schedule_ProspectingWeeklyAgendaReportGenerationUpdated & "," & GenerateProspectingWeeklyAgendaReportMondayTime
	Schedule_ProspectingWeeklyAgendaReportGenerationUpdated = Schedule_ProspectingWeeklyAgendaReportGenerationUpdated & "," & GenerateProspectingWeeklyAgendaReportTuesdayTime
	Schedule_ProspectingWeeklyAgendaReportGenerationUpdated = Schedule_ProspectingWeeklyAgendaReportGenerationUpdated & "," & GenerateProspectingWeeklyAgendaReportWednesdayTime
	Schedule_ProspectingWeeklyAgendaReportGenerationUpdated = Schedule_ProspectingWeeklyAgendaReportGenerationUpdated & "," & GenerateProspectingWeeklyAgendaReportThursdayTime
	Schedule_ProspectingWeeklyAgendaReportGenerationUpdated = Schedule_ProspectingWeeklyAgendaReportGenerationUpdated & "," & GenerateProspectingWeeklyAgendaReportFridayTime
	Schedule_ProspectingWeeklyAgendaReportGenerationUpdated = Schedule_ProspectingWeeklyAgendaReportGenerationUpdated & "," & GenerateProspectingWeeklyAgendaReportSaturdayTime

	
	Schedule_ProspectingWeeklyAgendaReportGenerationUpdated = Schedule_ProspectingWeeklyAgendaReportGenerationUpdated & "," & RunProspectingWeeklyAgendaReportIfClosed
	Schedule_ProspectingWeeklyAgendaReportGenerationUpdated = Schedule_ProspectingWeeklyAgendaReportGenerationUpdated & "," & RunProspectingWeeklyAgendaReportIfClosingEarly
	
	Response.Write("<br><br><br>Schedule_ProspectingWeeklyAgendaReportGenerationUpdated: " & Schedule_ProspectingWeeklyAgendaReportGenerationUpdated)

	'*********************************************************************
	'Update SQL with Array/String of Schedule Data
	'*********************************************************************
		
	SQL = "UPDATE Settings_Prospecting SET Schedule_ProspectingWeeklyAgendaReportGeneration = '" & cStr(Schedule_ProspectingWeeklyAgendaReportGenerationUpdated) & "' "
	
	Response.Write("<br><br><br>SQL: " & SQL)
	
	Set cnn8 = Server.CreateObject("ADODB.Connection")
	cnn8.open (Session("ClientCnnString"))
	Set rs = Server.CreateObject("ADODB.Recordset")
	rs.CursorLocation = 3 
	Set rs = cnn8.Execute(SQL)
	
	set rs = Nothing
	cnn8.close
	set cnn8 = Nothing


	 Response.Redirect("prospecting-settings.asp")
	
%><!--#include file="../../../inc/footer-main.asp"-->