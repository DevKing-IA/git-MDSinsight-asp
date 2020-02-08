<!--#include file="../../../../inc/header.asp"-->

<%

	'***********************************************************
	'Get Values Of All Form Fields Posted 
	'***********************************************************
	
	GenerateDailyAPIActivityByPartnerReportSunday = Request.Form("chkNoDailyAPIActivityByPartnerReportSunday")
	GenerateDailyAPIActivityByPartnerReportMonday = Request.Form("chkNoDailyAPIActivityByPartnerReportMonday")
	GenerateDailyAPIActivityByPartnerReportTuesday = Request.Form("chkNoDailyAPIActivityByPartnerReportTuesday")
	GenerateDailyAPIActivityByPartnerReportWednesday = Request.Form("chkNoDailyAPIActivityByPartnerReportWednesday")
	GenerateDailyAPIActivityByPartnerReportThursday = Request.Form("chkNoDailyAPIActivityByPartnerReportThursday")
	GenerateDailyAPIActivityByPartnerReportFriday = Request.Form("chkNoDailyAPIActivityByPartnerReportFriday")
	GenerateDailyAPIActivityByPartnerReportSaturday = Request.Form("chkNoDailyAPIActivityByPartnerReportSaturday")
	
	GenerateDailyAPIActivityByPartnerReportSundayTime = Request.Form("txtDailyAPIActivityByPartnerReportSchedulerSundayTime")
	GenerateDailyAPIActivityByPartnerReportMondayTime = Request.Form("txtDailyAPIActivityByPartnerReportSchedulerMondayTime")
	GenerateDailyAPIActivityByPartnerReportTuesdayTime = Request.Form("txtDailyAPIActivityByPartnerReportSchedulerTuesdayTime")
	GenerateDailyAPIActivityByPartnerReportWednesdayTime = Request.Form("txtDailyAPIActivityByPartnerReportSchedulerWednesdayTime")
	GenerateDailyAPIActivityByPartnerReportThursdayTime = Request.Form("txtDailyAPIActivityByPartnerReportSchedulerThursdayTime")
	GenerateDailyAPIActivityByPartnerReportFridayTime = Request.Form("txtDailyAPIActivityByPartnerReportSchedulerFridayTime")
	GenerateDailyAPIActivityByPartnerReportSaturdayTime = Request.Form("txtDailyAPIActivityByPartnerReportSchedulerSaturdayTime")
	
	RunDailyAPIActivityByPartnerReportIfClosed = Request.Form("chkNoDailyAPIActivityByPartnerReportIfClosed")
	RunDailyAPIActivityByPartnerReportIfClosingEarly = Request.Form("chkNoDailyAPIActivityByPartnerReportIfClosingEarly")


	If Request.Form("chkNoDailyAPIActivityByPartnerReportSunday") = "on" Then
		GenerateDailyAPIActivityByPartnerReportSunday = 0
		GenerateDailyAPIActivityByPartnerReportSundayTime = ""
	Else 
		GenerateDailyAPIActivityByPartnerReportSunday = 1
	End If

	If Request.Form("chkNoDailyAPIActivityByPartnerReportMonday") = "on" Then
		GenerateDailyAPIActivityByPartnerReportMonday = 0
		GenerateDailyAPIActivityByPartnerReportMondayTime = ""
	Else 
		GenerateDailyAPIActivityByPartnerReportMonday = 1
	End If

	If Request.Form("chkNoDailyAPIActivityByPartnerReportTuesday") = "on" Then
		GenerateDailyAPIActivityByPartnerReportTuesday = 0
		GenerateDailyAPIActivityByPartnerReportTuesdayTime = ""
	Else 
		GenerateDailyAPIActivityByPartnerReportTuesday = 1
	End If

	If Request.Form("chkNoDailyAPIActivityByPartnerReportWednesday") = "on" Then
		GenerateDailyAPIActivityByPartnerReportWednesday = 0
		GenerateDailyAPIActivityByPartnerReportWednesdayTime = ""
	Else 
		GenerateDailyAPIActivityByPartnerReportWednesday = 1
	End If

	If Request.Form("chkNoDailyAPIActivityByPartnerReportThursday") = "on" Then
		GenerateDailyAPIActivityByPartnerReportThursday = 0
		GenerateDailyAPIActivityByPartnerReportThursdayTime = ""
	Else 
		GenerateDailyAPIActivityByPartnerReportThursday = 1
	End If

	If Request.Form("chkNoDailyAPIActivityByPartnerReportFriday") = "on" Then
		GenerateDailyAPIActivityByPartnerReportFriday = 0
		GenerateDailyAPIActivityByPartnerReportFridayTime = ""
	Else 
		GenerateDailyAPIActivityByPartnerReportFriday = 1
	End If

	If Request.Form("chkNoDailyAPIActivityByPartnerReportSaturday") = "on" Then
		GenerateDailyAPIActivityByPartnerReportSaturday = 0
		GenerateDailyAPIActivityByPartnerReportSaturdayTime = ""
	Else 
		GenerateDailyAPIActivityByPartnerReportSaturday = 1
	End If

	If Request.Form("chkNoDailyAPIActivityByPartnerReportIfClosed") = "on" Then RunDailyAPIActivityByPartnerReportIfClosed = 0 Else RunDailyAPIActivityByPartnerReportIfClosed = 1
	If Request.Form("chkNoDailyAPIActivityByPartnerReportIfClosingEarly") = "on" Then RunDailyAPIActivityByPartnerReportIfClosingEarly = 0 Else RunDailyAPIActivityByPartnerReportIfClosingEarly = 1
	
	'***********************************************************
	'Get Original Values For Audit Trail Entries
	'***********************************************************
	
	SQLPropsectingSettings = "SELECT * FROM Settings_API"
	
	Set cnnPropsectingSettings = Server.CreateObject("ADODB.Connection")
	cnnPropsectingSettings.open (Session("ClientCnnString"))
	Set rsPropsectingSettings = Server.CreateObject("ADODB.Recordset")
	rsPropsectingSettings.CursorLocation = 3 
	Set rsPropsectingSettings = cnnPropsectingSettings.Execute(SQLPropsectingSettings)
		
	If NOT rsPropsectingSettings.EOF Then
	
		Schedule_DailyAPIActivityByPartnerReportGeneration = rsPropsectingSettings("Schedule_DailyAPIActivityByPartnerReportGeneration")
		
		Schedule_DailyAPIActivityByPartnerReportGenerationSettings = Split(Schedule_DailyAPIActivityByPartnerReportGeneration,",")

		GenerateDailyAPIActivityByPartnerReportSunday_ORIG = cInt(Schedule_DailyAPIActivityByPartnerReportGenerationSettings(0))
		GenerateDailyAPIActivityByPartnerReportMonday_ORIG = cInt(Schedule_DailyAPIActivityByPartnerReportGenerationSettings(1))
		GenerateDailyAPIActivityByPartnerReportTuesday_ORIG = cInt(Schedule_DailyAPIActivityByPartnerReportGenerationSettings(2))
		GenerateDailyAPIActivityByPartnerReportWednesday_ORIG = cInt(Schedule_DailyAPIActivityByPartnerReportGenerationSettings(3))
		GenerateDailyAPIActivityByPartnerReportThursday_ORIG = cInt(Schedule_DailyAPIActivityByPartnerReportGenerationSettings(4))
		GenerateDailyAPIActivityByPartnerReportFriday_ORIG = cInt(Schedule_DailyAPIActivityByPartnerReportGenerationSettings(5))
		GenerateDailyAPIActivityByPartnerReportSaturday_ORIG = cInt(Schedule_DailyAPIActivityByPartnerReportGenerationSettings(6))
		GenerateDailyAPIActivityByPartnerReportSundayTime_ORIG = Schedule_DailyAPIActivityByPartnerReportGenerationSettings(7)
		GenerateDailyAPIActivityByPartnerReportMondayTime_ORIG = Schedule_DailyAPIActivityByPartnerReportGenerationSettings(8)
		GenerateDailyAPIActivityByPartnerReportTuesdayTime_ORIG = Schedule_DailyAPIActivityByPartnerReportGenerationSettings(9)
		GenerateDailyAPIActivityByPartnerReportWednesdayTime_ORIG = Schedule_DailyAPIActivityByPartnerReportGenerationSettings(10)
		GenerateDailyAPIActivityByPartnerReportThursdayTime_ORIG = Schedule_DailyAPIActivityByPartnerReportGenerationSettings(11)
		GenerateDailyAPIActivityByPartnerReportFridayTime_ORIG = Schedule_DailyAPIActivityByPartnerReportGenerationSettings(12)
		GenerateDailyAPIActivityByPartnerReportSaturdayTime_ORIG = Schedule_DailyAPIActivityByPartnerReportGenerationSettings(13)
		RunDailyAPIActivityByPartnerReportIfClosed_ORIG = cInt(Schedule_DailyAPIActivityByPartnerReportGenerationSettings(14))
		RunDailyAPIActivityByPartnerReportIfClosingEarly_ORIG = cInt(Schedule_DailyAPIActivityByPartnerReportGenerationSettings(15))
	
	End If
	
	set rsPropsectingSettings = Nothing
	cnnPropsectingSettings.close
	set cnnPropsectingSettings = Nothing
	

	'***********************************************************
	'Perform Audit Trail Entry Inserts
	'***********************************************************

	
	If Request.Form("chkNoDailyAPIActivityByPartnerReportSunday") = "on" then GenerateDailyAPIActivityByPartnerReportSundayMsg = "On" Else GenerateDailyAPIActivityByPartnerReportSundayMsg = "Off"
	If GenerateDailyAPIActivityByPartnerReportSunday_ORIG = 1 then GenerateDailyAPIActivityByPartnerReportSundayMsgOrig = "On" Else GenerateDailyAPIActivityByPartnerReportSundayMsgOrig = "Off"
	
	If GenerateDailyAPIActivityByPartnerReportSunday <> GenerateDailyAPIActivityByPartnerReportSunday_ORIG Then
		CreateAuditLogEntry "Global Settings Change", "Global Settings Change", "Major", 1, "Daily API Activity By Partner Report generation schedule for SUNDAY changed from " & GenerateDailyAPIActivityByPartnerReportSundayMsgOrig & " to " & GenerateDailyAPIActivityByPartnerReportSundayMsg
	End If
	
	If GenerateDailyAPIActivityByPartnerReportSundayTime <> GenerateDailyAPIActivityByPartnerReportSundayTime_ORIG Then
		If GenerateDailyAPIActivityByPartnerReportSunday_ORIG = 0 AND GenerateDailyAPIActivityByPartnerReportSunday = 1 Then
			CreateAuditLogEntry "Global Settings Change", "Global Settings Change", "Major", 1, "Daily API Activity By Partner Report generation schedule for SUNDAY turned on and set to run at " & GenerateDailyAPIActivityByPartnerReportSundayTime
		Else
			CreateAuditLogEntry "Global Settings Change", "Global Settings Change", "Major", 1, "Daily API Activity By Partner Report generation scheduled run time for SUNDAY changed from " & GenerateDailyAPIActivityByPartnerReportSundayTime_ORIG & " to " & GenerateDailyAPIActivityByPartnerReportSundayTime
		End If
	End If
	
	

	If Request.Form("chkNoDailyAPIActivityByPartnerReportMonday") = "on" then GenerateDailyAPIActivityByPartnerReportMondayMsg = "On" Else GenerateDailyAPIActivityByPartnerReportMondayMsg = "Off"
	If GenerateDailyAPIActivityByPartnerReportMonday_ORIG = 1 then GenerateDailyAPIActivityByPartnerReportMondayMsgOrig = "On" Else GenerateDailyAPIActivityByPartnerReportMondayMsgOrig = "Off"
	
	If GenerateDailyAPIActivityByPartnerReportMonday <> GenerateDailyAPIActivityByPartnerReportMonday_ORIG Then
		CreateAuditLogEntry "Global Settings Change", "Global Settings Change", "Major", 1, "Daily API Activity By Partner Report generation schedule for Monday changed from " & GenerateDailyAPIActivityByPartnerReportMondayMsgOrig & " to " & GenerateDailyAPIActivityByPartnerReportMondayMsg
	End If
	
	If GenerateDailyAPIActivityByPartnerReportMondayTime <> GenerateDailyAPIActivityByPartnerReportMondayTime_ORIG Then
		If GenerateDailyAPIActivityByPartnerReportMonday_ORIG = 0 AND GenerateDailyAPIActivityByPartnerReportMonday = 1 Then
			CreateAuditLogEntry "Global Settings Change", "Global Settings Change", "Major", 1, "Daily API Activity By Partner Report generation schedule for Monday turned on and set to run at " & GenerateDailyAPIActivityByPartnerReportMondayTime
		Else
			CreateAuditLogEntry "Global Settings Change", "Global Settings Change", "Major", 1, "Daily API Activity By Partner Report generation scheduled run time for Monday changed from " & GenerateDailyAPIActivityByPartnerReportMondayTime_ORIG & " to " & GenerateDailyAPIActivityByPartnerReportMondayTime
		End If
	End If
	


	If Request.Form("chkNoDailyAPIActivityByPartnerReportTuesday") = "on" then GenerateDailyAPIActivityByPartnerReportTuesdayMsg = "On" Else GenerateDailyAPIActivityByPartnerReportTuesdayMsg = "Off"
	If GenerateDailyAPIActivityByPartnerReportTuesday_ORIG = 1 then GenerateDailyAPIActivityByPartnerReportTuesdayMsgOrig = "On" Else GenerateDailyAPIActivityByPartnerReportTuesdayMsgOrig = "Off"
	
	If GenerateDailyAPIActivityByPartnerReportTuesday <> GenerateDailyAPIActivityByPartnerReportTuesday_ORIG Then
		CreateAuditLogEntry "Global Settings Change", "Global Settings Change", "Major", 1, "Daily API Activity By Partner Report generation schedule for Tuesday changed from " & GenerateDailyAPIActivityByPartnerReportTuesdayMsgOrig & " to " & GenerateDailyAPIActivityByPartnerReportTuesdayMsg
	End If
	
	If GenerateDailyAPIActivityByPartnerReportTuesdayTime <> GenerateDailyAPIActivityByPartnerReportTuesdayTime_ORIG Then
		If GenerateDailyAPIActivityByPartnerReportTuesday_ORIG = 0 AND GenerateDailyAPIActivityByPartnerReportTuesday = 1 Then
			CreateAuditLogEntry "Global Settings Change", "Global Settings Change", "Major", 1, "Daily API Activity By Partner Report generation schedule for Tuesday turned on and set to run at " & GenerateDailyAPIActivityByPartnerReportTuesdayTime
		Else
			CreateAuditLogEntry "Global Settings Change", "Global Settings Change", "Major", 1, "Daily API Activity By Partner Report generation scheduled run time for Tuesday changed from " & GenerateDailyAPIActivityByPartnerReportTuesdayTime_ORIG & " to " & GenerateDailyAPIActivityByPartnerReportTuesdayTime
		End If
	End If



	If Request.Form("chkNoDailyAPIActivityByPartnerReportWednesday") = "on" then GenerateDailyAPIActivityByPartnerReportWednesdayMsg = "On" Else GenerateDailyAPIActivityByPartnerReportWednesdayMsg = "Off"
	If GenerateDailyAPIActivityByPartnerReportWednesday_ORIG = 1 then GenerateDailyAPIActivityByPartnerReportWednesdayMsgOrig = "On" Else GenerateDailyAPIActivityByPartnerReportWednesdayMsgOrig = "Off"
	
	If GenerateDailyAPIActivityByPartnerReportWednesday <> GenerateDailyAPIActivityByPartnerReportWednesday_ORIG Then
		CreateAuditLogEntry "Global Settings Change", "Global Settings Change", "Major", 1, "Daily API Activity By Partner Report generation schedule for Wednesday changed from " & GenerateDailyAPIActivityByPartnerReportWednesdayMsgOrig & " to " & GenerateDailyAPIActivityByPartnerReportWednesdayMsg
	End If
	
	If GenerateDailyAPIActivityByPartnerReportWednesdayTime <> GenerateDailyAPIActivityByPartnerReportWednesdayTime_ORIG Then
		If GenerateDailyAPIActivityByPartnerReportWednesday_ORIG = 0 AND GenerateDailyAPIActivityByPartnerReportWednesday = 1 Then
			CreateAuditLogEntry "Global Settings Change", "Global Settings Change", "Major", 1, "Daily API Activity By Partner Report generation schedule for Wednesday turned on and set to run at " & GenerateDailyAPIActivityByPartnerReportWednesdayTime
		Else
			CreateAuditLogEntry "Global Settings Change", "Global Settings Change", "Major", 1, "Daily API Activity By Partner Report generation scheduled run time for Wednesday changed from " & GenerateDailyAPIActivityByPartnerReportWednesdayTime_ORIG & " to " & GenerateDailyAPIActivityByPartnerReportWednesdayTime
		End If
	End If



	If Request.Form("chkNoDailyAPIActivityByPartnerReportThursday") = "on" then GenerateDailyAPIActivityByPartnerReportThursdayMsg = "On" Else GenerateDailyAPIActivityByPartnerReportThursdayMsg = "Off"
	If GenerateDailyAPIActivityByPartnerReportThursday_ORIG = 1 then GenerateDailyAPIActivityByPartnerReportThursdayMsgOrig = "On" Else GenerateDailyAPIActivityByPartnerReportThursdayMsgOrig = "Off"
	
	If GenerateDailyAPIActivityByPartnerReportThursday <> GenerateDailyAPIActivityByPartnerReportThursday_ORIG Then
		CreateAuditLogEntry "Global Settings Change", "Global Settings Change", "Major", 1, "Daily API Activity By Partner Report generation schedule for Thursday changed from " & GenerateDailyAPIActivityByPartnerReportThursdayMsgOrig & " to " & GenerateDailyAPIActivityByPartnerReportThursdayMsg
	End If
	
	If GenerateDailyAPIActivityByPartnerReportThursdayTime <> GenerateDailyAPIActivityByPartnerReportThursdayTime_ORIG Then
		If GenerateDailyAPIActivityByPartnerReportThursday_ORIG = 0 AND GenerateDailyAPIActivityByPartnerReportThursday = 1 Then
			CreateAuditLogEntry "Global Settings Change", "Global Settings Change", "Major", 1, "Daily API Activity By Partner Report generation schedule for Thursday turned on and set to run at " & GenerateDailyAPIActivityByPartnerReportThursdayTime
		Else
			CreateAuditLogEntry "Global Settings Change", "Global Settings Change", "Major", 1, "Daily API Activity By Partner Report generation scheduled run time for Thursday changed from " & GenerateDailyAPIActivityByPartnerReportThursdayTime_ORIG & " to " & GenerateDailyAPIActivityByPartnerReportThursdayTime
		End If
	End If



	If Request.Form("chkNoDailyAPIActivityByPartnerReportFriday") = "on" then GenerateDailyAPIActivityByPartnerReportFridayMsg = "On" Else GenerateDailyAPIActivityByPartnerReportFridayMsg = "Off"
	If GenerateDailyAPIActivityByPartnerReportFriday_ORIG = 1 then GenerateDailyAPIActivityByPartnerReportFridayMsgOrig = "On" Else GenerateDailyAPIActivityByPartnerReportFridayMsgOrig = "Off"
	
	If GenerateDailyAPIActivityByPartnerReportFriday <> GenerateDailyAPIActivityByPartnerReportFriday_ORIG Then
		CreateAuditLogEntry "Global Settings Change", "Global Settings Change", "Major", 1, "Daily API Activity By Partner Report generation schedule for Friday changed from " & GenerateDailyAPIActivityByPartnerReportFridayMsgOrig & " to " & GenerateDailyAPIActivityByPartnerReportFridayMsg
	End If
	
	If GenerateDailyAPIActivityByPartnerReportFridayTime <> GenerateDailyAPIActivityByPartnerReportFridayTime_ORIG Then
		If GenerateDailyAPIActivityByPartnerReportFriday_ORIG = 0 AND GenerateDailyAPIActivityByPartnerReportFriday = 1 Then
			CreateAuditLogEntry "Global Settings Change", "Global Settings Change", "Major", 1, "Daily API Activity By Partner Report generation schedule for Friday turned on and set to run at " & GenerateDailyAPIActivityByPartnerReportFridayTime
		Else
			CreateAuditLogEntry "Global Settings Change", "Global Settings Change", "Major", 1, "Daily API Activity By Partner Report generation scheduled run time for Friday changed from " & GenerateDailyAPIActivityByPartnerReportFridayTime_ORIG & " to " & GenerateDailyAPIActivityByPartnerReportFridayTime
		End If
	End If



	If Request.Form("chkNoDailyAPIActivityByPartnerReportSaturday") = "on" then GenerateDailyAPIActivityByPartnerReportSaturdayMsg = "On" Else GenerateDailyAPIActivityByPartnerReportSaturdayMsg = "Off"
	If GenerateDailyAPIActivityByPartnerReportSaturday_ORIG = 1 then GenerateDailyAPIActivityByPartnerReportSaturdayMsgOrig = "On" Else GenerateDailyAPIActivityByPartnerReportSaturdayMsgOrig = "Off"
	
	If GenerateDailyAPIActivityByPartnerReportSaturday <> GenerateDailyAPIActivityByPartnerReportSaturday_ORIG Then
		CreateAuditLogEntry "Global Settings Change", "Global Settings Change", "Major", 1, "Daily API Activity By Partner Report generation schedule for Saturday changed from " & GenerateDailyAPIActivityByPartnerReportSaturdayMsgOrig & " to " & GenerateDailyAPIActivityByPartnerReportSaturdayMsg
	End If
	
	If GenerateDailyAPIActivityByPartnerReportSaturdayTime <> GenerateDailyAPIActivityByPartnerReportSaturdayTime_ORIG Then
		If GenerateDailyAPIActivityByPartnerReportSaturday_ORIG = 0 AND GenerateDailyAPIActivityByPartnerReportSaturday = 1 Then
			CreateAuditLogEntry "Global Settings Change", "Global Settings Change", "Major", 1, "Daily API Activity By Partner Report generation schedule for Saturday turned on and set to run at " & GenerateDailyAPIActivityByPartnerReportSaturdayTime
		Else
			CreateAuditLogEntry "Global Settings Change", "Global Settings Change", "Major", 1, "Daily API Activity By Partner Report generation scheduled run time for Saturday changed from " & GenerateDailyAPIActivityByPartnerReportSaturdayTime_ORIG & " to " & GenerateDailyAPIActivityByPartnerReportSaturdayTime
		End If
	End If


	If Request.Form("chkNoDailyAPIActivityByPartnerReportIfClosed") = "on" then RunDailyAPIActivityByPartnerReportIfClosedMsg = "On" Else RunDailyAPIActivityByPartnerReportIfClosedMsg = "Off"
	If RunDailyAPIActivityByPartnerReportIfClosed_ORIG = 1 then RunDailyAPIActivityByPartnerReportIfClosedMsgOrig = "On" Else RunDailyAPIActivityByPartnerReportIfClosedMsgOrig = "Off"
	
	If RunDailyAPIActivityByPartnerReportIfClosed <> RunDailyAPIActivityByPartnerReportIfClosed_ORIG Then
		CreateAuditLogEntry "Global Settings Change", "Global Settings Change", "Major", 1, "Daily API Activity By Partner Report generation schedule - Don''t Run If Closed (Monday-Friday Only) -  changed from " & RunDailyAPIActivityByPartnerReportIfClosedMsgOrig & " to " & RunDailyAPIActivityByPartnerReportIfClosedMsg
	End If


	If Request.Form("chkNoDailyAPIActivityByPartnerReportIfClosingEarly") = "on" then RunDailyAPIActivityByPartnerReportIfClosingEarlyMsg = "On" Else RunDailyAPIActivityByPartnerReportIfClosingEarlyMsg = "Off"
	If RunDailyAPIActivityByPartnerReportIfClosingEarly_ORIG = 1 then RunDailyAPIActivityByPartnerReportIfClosingEarlyMsgOrig = "On" Else RunDailyAPIActivityByPartnerReportIfClosingEarlyMsgOrig = "Off"
	
	If RunDailyAPIActivityByPartnerReportIfClosingEarly <> RunDailyAPIActivityByPartnerReportIfClosingEarly_ORIG Then
		CreateAuditLogEntry "Global Settings Change", "Global Settings Change", "Major", 1, "Daily API Activity By Partner Report generation schedule - Don''t Run If Closed (Monday-Friday Only) -  changed from " & RunDailyAPIActivityByPartnerReportIfClosingEarlyMsgOrig & " to " & RunDailyAPIActivityByPartnerReportIfClosingEarlyMsg
	End If


	'*********************************************************************
	'Build Array/String of Schedule Data From Request Form Field Data
	'*********************************************************************

	Schedule_DailyAPIActivityByPartnerReportGenerationUpdated = ""
	
	Schedule_DailyAPIActivityByPartnerReportGenerationUpdated = GenerateDailyAPIActivityByPartnerReportSunday
	Schedule_DailyAPIActivityByPartnerReportGenerationUpdated = Schedule_DailyAPIActivityByPartnerReportGenerationUpdated & "," & GenerateDailyAPIActivityByPartnerReportMonday
	Schedule_DailyAPIActivityByPartnerReportGenerationUpdated = Schedule_DailyAPIActivityByPartnerReportGenerationUpdated & "," & GenerateDailyAPIActivityByPartnerReportTuesday
	Schedule_DailyAPIActivityByPartnerReportGenerationUpdated = Schedule_DailyAPIActivityByPartnerReportGenerationUpdated & "," & GenerateDailyAPIActivityByPartnerReportWednesday
	Schedule_DailyAPIActivityByPartnerReportGenerationUpdated = Schedule_DailyAPIActivityByPartnerReportGenerationUpdated & "," & GenerateDailyAPIActivityByPartnerReportThursday
	Schedule_DailyAPIActivityByPartnerReportGenerationUpdated = Schedule_DailyAPIActivityByPartnerReportGenerationUpdated & "," & GenerateDailyAPIActivityByPartnerReportFriday
	Schedule_DailyAPIActivityByPartnerReportGenerationUpdated = Schedule_DailyAPIActivityByPartnerReportGenerationUpdated & "," & GenerateDailyAPIActivityByPartnerReportSaturday
	
	Schedule_DailyAPIActivityByPartnerReportGenerationUpdated = Schedule_DailyAPIActivityByPartnerReportGenerationUpdated & "," & GenerateDailyAPIActivityByPartnerReportSundayTime
	Schedule_DailyAPIActivityByPartnerReportGenerationUpdated = Schedule_DailyAPIActivityByPartnerReportGenerationUpdated & "," & GenerateDailyAPIActivityByPartnerReportMondayTime
	Schedule_DailyAPIActivityByPartnerReportGenerationUpdated = Schedule_DailyAPIActivityByPartnerReportGenerationUpdated & "," & GenerateDailyAPIActivityByPartnerReportTuesdayTime
	Schedule_DailyAPIActivityByPartnerReportGenerationUpdated = Schedule_DailyAPIActivityByPartnerReportGenerationUpdated & "," & GenerateDailyAPIActivityByPartnerReportWednesdayTime
	Schedule_DailyAPIActivityByPartnerReportGenerationUpdated = Schedule_DailyAPIActivityByPartnerReportGenerationUpdated & "," & GenerateDailyAPIActivityByPartnerReportThursdayTime
	Schedule_DailyAPIActivityByPartnerReportGenerationUpdated = Schedule_DailyAPIActivityByPartnerReportGenerationUpdated & "," & GenerateDailyAPIActivityByPartnerReportFridayTime
	Schedule_DailyAPIActivityByPartnerReportGenerationUpdated = Schedule_DailyAPIActivityByPartnerReportGenerationUpdated & "," & GenerateDailyAPIActivityByPartnerReportSaturdayTime

	
	Schedule_DailyAPIActivityByPartnerReportGenerationUpdated = Schedule_DailyAPIActivityByPartnerReportGenerationUpdated & "," & RunDailyAPIActivityByPartnerReportIfClosed
	Schedule_DailyAPIActivityByPartnerReportGenerationUpdated = Schedule_DailyAPIActivityByPartnerReportGenerationUpdated & "," & RunDailyAPIActivityByPartnerReportIfClosingEarly
	
	Response.Write("<br><br><br>Schedule_DailyAPIActivityByPartnerReportGenerationUpdated: " & Schedule_DailyAPIActivityByPartnerReportGenerationUpdated)

	'*********************************************************************
	'Update SQL with Array/String of Schedule Data
	'*********************************************************************
		
	SQL = "UPDATE Settings_API SET Schedule_DailyAPIActivityByPartnerReportGeneration = '" & cStr(Schedule_DailyAPIActivityByPartnerReportGenerationUpdated) & "' "
	
	Response.Write("<br><br><br>SQL: " & SQL)
	
	Set cnn8 = Server.CreateObject("ADODB.Connection")
	cnn8.open (Session("ClientCnnString"))
	Set rs = Server.CreateObject("ADODB.Recordset")
	rs.CursorLocation = 3 
	Set rs = cnn8.Execute(SQL)
	
	set rs = Nothing
	cnn8.close
	set cnn8 = Nothing


	 Response.Redirect("order-api.asp")
	
%><!--#include file="../../../../inc/footer-main.asp"-->