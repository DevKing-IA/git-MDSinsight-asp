<!--#include file="../../../inc/header.asp"-->

<%

	'***********************************************************
	'Get Values Of All Form Fields Posted
	'***********************************************************
	
	GenerateAutomaticCustomerAnalysisSummary1ReportSunday = Request.Form("chkNoAutomaticCustomerAnalysisSummary1ReportSunday")
	GenerateAutomaticCustomerAnalysisSummary1ReportMonday = Request.Form("chkNoAutomaticCustomerAnalysisSummary1ReportMonday")
	GenerateAutomaticCustomerAnalysisSummary1ReportTuesday = Request.Form("chkNoAutomaticCustomerAnalysisSummary1ReportTuesday")
	GenerateAutomaticCustomerAnalysisSummary1ReportWednesday = Request.Form("chkNoAutomaticCustomerAnalysisSummary1ReportWednesday")
	GenerateAutomaticCustomerAnalysisSummary1ReportThursday = Request.Form("chkNoAutomaticCustomerAnalysisSummary1ReportThursday")
	GenerateAutomaticCustomerAnalysisSummary1ReportFriday = Request.Form("chkNoAutomaticCustomerAnalysisSummary1ReportFriday")
	GenerateAutomaticCustomerAnalysisSummary1ReportSaturday = Request.Form("chkNoAutomaticCustomerAnalysisSummary1ReportSaturday")
	
	GenerateAutomaticCustomerAnalysisSummary1ReportSundayTime = Request.Form("txtAutomaticCustomerAnalysisSummary1ReportSchedulerSundayTime")
	GenerateAutomaticCustomerAnalysisSummary1ReportMondayTime = Request.Form("txtAutomaticCustomerAnalysisSummary1ReportSchedulerMondayTime")
	GenerateAutomaticCustomerAnalysisSummary1ReportTuesdayTime = Request.Form("txtAutomaticCustomerAnalysisSummary1ReportSchedulerTuesdayTime")
	GenerateAutomaticCustomerAnalysisSummary1ReportWednesdayTime = Request.Form("txtAutomaticCustomerAnalysisSummary1ReportSchedulerWednesdayTime")
	GenerateAutomaticCustomerAnalysisSummary1ReportThursdayTime = Request.Form("txtAutomaticCustomerAnalysisSummary1ReportSchedulerThursdayTime")
	GenerateAutomaticCustomerAnalysisSummary1ReportFridayTime = Request.Form("txtAutomaticCustomerAnalysisSummary1ReportSchedulerFridayTime")
	GenerateAutomaticCustomerAnalysisSummary1ReportSaturdayTime = Request.Form("txtAutomaticCustomerAnalysisSummary1ReportSchedulerSaturdayTime")
	
	RunAutomaticCustomerAnalysisSummary1ReportIfClosed = Request.Form("chkNoAutomaticCustomerAnalysisSummary1ReportIfClosed")
	RunAutomaticCustomerAnalysisSummary1ReportIfClosingEarly = Request.Form("chkNoAutomaticCustomerAnalysisSummary1ReportIfClosingEarly")


	If Request.Form("chkNoAutomaticCustomerAnalysisSummary1ReportSunday") = "on" Then
		GenerateAutomaticCustomerAnalysisSummary1ReportSunday = 0
		GenerateAutomaticCustomerAnalysisSummary1ReportSundayTime = ""
	Else 
		GenerateAutomaticCustomerAnalysisSummary1ReportSunday = 1
	End If

	If Request.Form("chkNoAutomaticCustomerAnalysisSummary1ReportMonday") = "on" Then
		GenerateAutomaticCustomerAnalysisSummary1ReportMonday = 0
		GenerateAutomaticCustomerAnalysisSummary1ReportMondayTime = ""
	Else 
		GenerateAutomaticCustomerAnalysisSummary1ReportMonday = 1
	End If

	If Request.Form("chkNoAutomaticCustomerAnalysisSummary1ReportTuesday") = "on" Then
		GenerateAutomaticCustomerAnalysisSummary1ReportTuesday = 0
		GenerateAutomaticCustomerAnalysisSummary1ReportTuesdayTime = ""
	Else 
		GenerateAutomaticCustomerAnalysisSummary1ReportTuesday = 1
	End If

	If Request.Form("chkNoAutomaticCustomerAnalysisSummary1ReportWednesday") = "on" Then
		GenerateAutomaticCustomerAnalysisSummary1ReportWednesday = 0
		GenerateAutomaticCustomerAnalysisSummary1ReportWednesdayTime = ""
	Else 
		GenerateAutomaticCustomerAnalysisSummary1ReportWednesday = 1
	End If

	If Request.Form("chkNoAutomaticCustomerAnalysisSummary1ReportThursday") = "on" Then
		GenerateAutomaticCustomerAnalysisSummary1ReportThursday = 0
		GenerateAutomaticCustomerAnalysisSummary1ReportThursdayTime = ""
	Else 
		GenerateAutomaticCustomerAnalysisSummary1ReportThursday = 1
	End If

	If Request.Form("chkNoAutomaticCustomerAnalysisSummary1ReportFriday") = "on" Then
		GenerateAutomaticCustomerAnalysisSummary1ReportFriday = 0
		GenerateAutomaticCustomerAnalysisSummary1ReportFridayTime = ""
	Else 
		GenerateAutomaticCustomerAnalysisSummary1ReportFriday = 1
	End If

	If Request.Form("chkNoAutomaticCustomerAnalysisSummary1ReportSaturday") = "on" Then
		GenerateAutomaticCustomerAnalysisSummary1ReportSaturday = 0
		GenerateAutomaticCustomerAnalysisSummary1ReportSaturdayTime = ""
	Else 
		GenerateAutomaticCustomerAnalysisSummary1ReportSaturday = 1
	End If

	If Request.Form("chkNoAutomaticCustomerAnalysisSummary1ReportIfClosed") = "on" Then RunAutomaticCustomerAnalysisSummary1ReportIfClosed = 0 Else RunAutomaticCustomerAnalysisSummary1ReportIfClosed = 1
	If Request.Form("chkNoAutomaticCustomerAnalysisSummary1ReportIfClosingEarly") = "on" Then RunAutomaticCustomerAnalysisSummary1ReportIfClosingEarly = 0 Else RunAutomaticCustomerAnalysisSummary1ReportIfClosingEarly = 1
	
	'***********************************************************
	'Get Original Values For Audit Trail Entries
	'***********************************************************
	
	SQLFieldServiceSettings = "SELECT * FROM Settings_BizIntel"
	
	Set cnnFieldServiceSettings = Server.CreateObject("ADODB.Connection")
	cnnFieldServiceSettings.open (Session("ClientCnnString"))
	Set rsFieldServiceSettings = Server.CreateObject("ADODB.Recordset")
	rsFieldServiceSettings.CursorLocation = 3 
	Set rsFieldServiceSettings = cnnFieldServiceSettings.Execute(SQLFieldServiceSettings)
		
	If NOT rsFieldServiceSettings.EOF Then
	
		Schedule_AutomaticCustomerAnalysisSummary1ReportGeneration = rsFieldServiceSettings("Schedule_AutomaticCustomerAnalysisSummary1ReportGeneration")
		
		Schedule_AutomaticCustomerAnalysisSummary1ReportGenerationSettings = Split(Schedule_AutomaticCustomerAnalysisSummary1ReportGeneration,",")

		GenerateAutomaticCustomerAnalysisSummary1ReportSunday_ORIG = cInt(Schedule_AutomaticCustomerAnalysisSummary1ReportGenerationSettings(0))
		GenerateAutomaticCustomerAnalysisSummary1ReportMonday_ORIG = cInt(Schedule_AutomaticCustomerAnalysisSummary1ReportGenerationSettings(1))
		GenerateAutomaticCustomerAnalysisSummary1ReportTuesday_ORIG = cInt(Schedule_AutomaticCustomerAnalysisSummary1ReportGenerationSettings(2))
		GenerateAutomaticCustomerAnalysisSummary1ReportWednesday_ORIG = cInt(Schedule_AutomaticCustomerAnalysisSummary1ReportGenerationSettings(3))
		GenerateAutomaticCustomerAnalysisSummary1ReportThursday_ORIG = cInt(Schedule_AutomaticCustomerAnalysisSummary1ReportGenerationSettings(4))
		GenerateAutomaticCustomerAnalysisSummary1ReportFriday_ORIG = cInt(Schedule_AutomaticCustomerAnalysisSummary1ReportGenerationSettings(5))
		GenerateAutomaticCustomerAnalysisSummary1ReportSaturday_ORIG = cInt(Schedule_AutomaticCustomerAnalysisSummary1ReportGenerationSettings(6))
		GenerateAutomaticCustomerAnalysisSummary1ReportSundayTime_ORIG = Schedule_AutomaticCustomerAnalysisSummary1ReportGenerationSettings(7)
		GenerateAutomaticCustomerAnalysisSummary1ReportMondayTime_ORIG = Schedule_AutomaticCustomerAnalysisSummary1ReportGenerationSettings(8)
		GenerateAutomaticCustomerAnalysisSummary1ReportTuesdayTime_ORIG = Schedule_AutomaticCustomerAnalysisSummary1ReportGenerationSettings(9)
		GenerateAutomaticCustomerAnalysisSummary1ReportWednesdayTime_ORIG = Schedule_AutomaticCustomerAnalysisSummary1ReportGenerationSettings(10)
		GenerateAutomaticCustomerAnalysisSummary1ReportThursdayTime_ORIG = Schedule_AutomaticCustomerAnalysisSummary1ReportGenerationSettings(11)
		GenerateAutomaticCustomerAnalysisSummary1ReportFridayTime_ORIG = Schedule_AutomaticCustomerAnalysisSummary1ReportGenerationSettings(12)
		GenerateAutomaticCustomerAnalysisSummary1ReportSaturdayTime_ORIG = Schedule_AutomaticCustomerAnalysisSummary1ReportGenerationSettings(13)
		RunAutomaticCustomerAnalysisSummary1ReportIfClosed_ORIG = cInt(Schedule_AutomaticCustomerAnalysisSummary1ReportGenerationSettings(14))
		RunAutomaticCustomerAnalysisSummary1ReportIfClosingEarly_ORIG = cInt(Schedule_AutomaticCustomerAnalysisSummary1ReportGenerationSettings(15))
	
	End If
	
	set rsFieldServiceSettings = Nothing
	cnnFieldServiceSettings.close
	set cnnFieldServiceSettings = Nothing
	

	'***********************************************************
	'Perform Audit Trail Entry Inserts
	'***********************************************************

	
	If Request.Form("chkNoAutomaticCustomerAnalysisSummary1ReportSunday") = "on" then GenerateAutomaticCustomerAnalysisSummary1ReportSundayMsg = "On" Else GenerateAutomaticCustomerAnalysisSummary1ReportSundayMsg = "Off"
	If GenerateAutomaticCustomerAnalysisSummary1ReportSunday_ORIG = 1 then GenerateAutomaticCustomerAnalysisSummary1ReportSundayMsgOrig = "On" Else GenerateAutomaticCustomerAnalysisSummary1ReportSundayMsgOrig = "Off"
	
	If GenerateAutomaticCustomerAnalysisSummary1ReportSunday <> GenerateAutomaticCustomerAnalysisSummary1ReportSunday_ORIG Then
		CreateAuditLogEntry "Global Settings Change", "Global Settings Change", "Major", 1, "Automatic Customer Analysis Summary 1 Report generation schedule for SUNDAY changed from " & GenerateAutomaticCustomerAnalysisSummary1ReportSundayMsgOrig & " to " & GenerateAutomaticCustomerAnalysisSummary1ReportSundayMsg
	End If
	
	If GenerateAutomaticCustomerAnalysisSummary1ReportSundayTime <> GenerateAutomaticCustomerAnalysisSummary1ReportSundayTime_ORIG Then
		If GenerateAutomaticCustomerAnalysisSummary1ReportSunday_ORIG = 0 AND GenerateAutomaticCustomerAnalysisSummary1ReportSunday = 1 Then
			CreateAuditLogEntry "Global Settings Change", "Global Settings Change", "Major", 1, "Automatic Customer Analysis Summary 1 Report generation schedule for SUNDAY turned on and set to run at " & GenerateAutomaticCustomerAnalysisSummary1ReportSundayTime
		Else
			CreateAuditLogEntry "Global Settings Change", "Global Settings Change", "Major", 1, "Automatic Customer Analysis Summary 1 Report generation scheduled run time for SUNDAY changed from " & GenerateAutomaticCustomerAnalysisSummary1ReportSundayTime_ORIG & " to " & GenerateAutomaticCustomerAnalysisSummary1ReportSundayTime
		End If
	End If
	
	

	If Request.Form("chkNoAutomaticCustomerAnalysisSummary1ReportMonday") = "on" then GenerateAutomaticCustomerAnalysisSummary1ReportMondayMsg = "On" Else GenerateAutomaticCustomerAnalysisSummary1ReportMondayMsg = "Off"
	If GenerateAutomaticCustomerAnalysisSummary1ReportMonday_ORIG = 1 then GenerateAutomaticCustomerAnalysisSummary1ReportMondayMsgOrig = "On" Else GenerateAutomaticCustomerAnalysisSummary1ReportMondayMsgOrig = "Off"
	
	If GenerateAutomaticCustomerAnalysisSummary1ReportMonday <> GenerateAutomaticCustomerAnalysisSummary1ReportMonday_ORIG Then
		CreateAuditLogEntry "Global Settings Change", "Global Settings Change", "Major", 1, "Automatic Customer Analysis Summary 1 Report generation schedule for Monday changed from " & GenerateAutomaticCustomerAnalysisSummary1ReportMondayMsgOrig & " to " & GenerateAutomaticCustomerAnalysisSummary1ReportMondayMsg
	End If
	
	If GenerateAutomaticCustomerAnalysisSummary1ReportMondayTime <> GenerateAutomaticCustomerAnalysisSummary1ReportMondayTime_ORIG Then
		If GenerateAutomaticCustomerAnalysisSummary1ReportMonday_ORIG = 0 AND GenerateAutomaticCustomerAnalysisSummary1ReportMonday = 1 Then
			CreateAuditLogEntry "Global Settings Change", "Global Settings Change", "Major", 1, "Automatic Customer Analysis Summary 1 Report generation schedule for Monday turned on and set to run at " & GenerateAutomaticCustomerAnalysisSummary1ReportMondayTime
		Else
			CreateAuditLogEntry "Global Settings Change", "Global Settings Change", "Major", 1, "Automatic Customer Analysis Summary 1 Report generation scheduled run time for Monday changed from " & GenerateAutomaticCustomerAnalysisSummary1ReportMondayTime_ORIG & " to " & GenerateAutomaticCustomerAnalysisSummary1ReportMondayTime
		End If
	End If
	


	If Request.Form("chkNoAutomaticCustomerAnalysisSummary1ReportTuesday") = "on" then GenerateAutomaticCustomerAnalysisSummary1ReportTuesdayMsg = "On" Else GenerateAutomaticCustomerAnalysisSummary1ReportTuesdayMsg = "Off"
	If GenerateAutomaticCustomerAnalysisSummary1ReportTuesday_ORIG = 1 then GenerateAutomaticCustomerAnalysisSummary1ReportTuesdayMsgOrig = "On" Else GenerateAutomaticCustomerAnalysisSummary1ReportTuesdayMsgOrig = "Off"
	
	If GenerateAutomaticCustomerAnalysisSummary1ReportTuesday <> GenerateAutomaticCustomerAnalysisSummary1ReportTuesday_ORIG Then
		CreateAuditLogEntry "Global Settings Change", "Global Settings Change", "Major", 1, "Automatic Customer Analysis Summary 1 Report generation schedule for Tuesday changed from " & GenerateAutomaticCustomerAnalysisSummary1ReportTuesdayMsgOrig & " to " & GenerateAutomaticCustomerAnalysisSummary1ReportTuesdayMsg
	End If
	
	If GenerateAutomaticCustomerAnalysisSummary1ReportTuesdayTime <> GenerateAutomaticCustomerAnalysisSummary1ReportTuesdayTime_ORIG Then
		If GenerateAutomaticCustomerAnalysisSummary1ReportTuesday_ORIG = 0 AND GenerateAutomaticCustomerAnalysisSummary1ReportTuesday = 1 Then
			CreateAuditLogEntry "Global Settings Change", "Global Settings Change", "Major", 1, "Automatic Customer Analysis Summary 1 Report generation schedule for Tuesday turned on and set to run at " & GenerateAutomaticCustomerAnalysisSummary1ReportTuesdayTime
		Else
			CreateAuditLogEntry "Global Settings Change", "Global Settings Change", "Major", 1, "Automatic Customer Analysis Summary 1 Report generation scheduled run time for Tuesday changed from " & GenerateAutomaticCustomerAnalysisSummary1ReportTuesdayTime_ORIG & " to " & GenerateAutomaticCustomerAnalysisSummary1ReportTuesdayTime
		End If
	End If



	If Request.Form("chkNoAutomaticCustomerAnalysisSummary1ReportWednesday") = "on" then GenerateAutomaticCustomerAnalysisSummary1ReportWednesdayMsg = "On" Else GenerateAutomaticCustomerAnalysisSummary1ReportWednesdayMsg = "Off"
	If GenerateAutomaticCustomerAnalysisSummary1ReportWednesday_ORIG = 1 then GenerateAutomaticCustomerAnalysisSummary1ReportWednesdayMsgOrig = "On" Else GenerateAutomaticCustomerAnalysisSummary1ReportWednesdayMsgOrig = "Off"
	
	If GenerateAutomaticCustomerAnalysisSummary1ReportWednesday <> GenerateAutomaticCustomerAnalysisSummary1ReportWednesday_ORIG Then
		CreateAuditLogEntry "Global Settings Change", "Global Settings Change", "Major", 1, "Automatic Customer Analysis Summary 1 Report generation schedule for Wednesday changed from " & GenerateAutomaticCustomerAnalysisSummary1ReportWednesdayMsgOrig & " to " & GenerateAutomaticCustomerAnalysisSummary1ReportWednesdayMsg
	End If
	
	If GenerateAutomaticCustomerAnalysisSummary1ReportWednesdayTime <> GenerateAutomaticCustomerAnalysisSummary1ReportWednesdayTime_ORIG Then
		If GenerateAutomaticCustomerAnalysisSummary1ReportWednesday_ORIG = 0 AND GenerateAutomaticCustomerAnalysisSummary1ReportWednesday = 1 Then
			CreateAuditLogEntry "Global Settings Change", "Global Settings Change", "Major", 1, "Automatic Customer Analysis Summary 1 Report generation schedule for Wednesday turned on and set to run at " & GenerateAutomaticCustomerAnalysisSummary1ReportWednesdayTime
		Else
			CreateAuditLogEntry "Global Settings Change", "Global Settings Change", "Major", 1, "Automatic Customer Analysis Summary 1 Report generation scheduled run time for Wednesday changed from " & GenerateAutomaticCustomerAnalysisSummary1ReportWednesdayTime_ORIG & " to " & GenerateAutomaticCustomerAnalysisSummary1ReportWednesdayTime
		End If
	End If



	If Request.Form("chkNoAutomaticCustomerAnalysisSummary1ReportThursday") = "on" then GenerateAutomaticCustomerAnalysisSummary1ReportThursdayMsg = "On" Else GenerateAutomaticCustomerAnalysisSummary1ReportThursdayMsg = "Off"
	If GenerateAutomaticCustomerAnalysisSummary1ReportThursday_ORIG = 1 then GenerateAutomaticCustomerAnalysisSummary1ReportThursdayMsgOrig = "On" Else GenerateAutomaticCustomerAnalysisSummary1ReportThursdayMsgOrig = "Off"
	
	If GenerateAutomaticCustomerAnalysisSummary1ReportThursday <> GenerateAutomaticCustomerAnalysisSummary1ReportThursday_ORIG Then
		CreateAuditLogEntry "Global Settings Change", "Global Settings Change", "Major", 1, "Automatic Customer Analysis Summary 1 Report generation schedule for Thursday changed from " & GenerateAutomaticCustomerAnalysisSummary1ReportThursdayMsgOrig & " to " & GenerateAutomaticCustomerAnalysisSummary1ReportThursdayMsg
	End If
	
	If GenerateAutomaticCustomerAnalysisSummary1ReportThursdayTime <> GenerateAutomaticCustomerAnalysisSummary1ReportThursdayTime_ORIG Then
		If GenerateAutomaticCustomerAnalysisSummary1ReportThursday_ORIG = 0 AND GenerateAutomaticCustomerAnalysisSummary1ReportThursday = 1 Then
			CreateAuditLogEntry "Global Settings Change", "Global Settings Change", "Major", 1, "Automatic Customer Analysis Summary 1 Report generation schedule for Thursday turned on and set to run at " & GenerateAutomaticCustomerAnalysisSummary1ReportThursdayTime
		Else
			CreateAuditLogEntry "Global Settings Change", "Global Settings Change", "Major", 1, "Automatic Customer Analysis Summary 1 Report generation scheduled run time for Thursday changed from " & GenerateAutomaticCustomerAnalysisSummary1ReportThursdayTime_ORIG & " to " & GenerateAutomaticCustomerAnalysisSummary1ReportThursdayTime
		End If
	End If



	If Request.Form("chkNoAutomaticCustomerAnalysisSummary1ReportFriday") = "on" then GenerateAutomaticCustomerAnalysisSummary1ReportFridayMsg = "On" Else GenerateAutomaticCustomerAnalysisSummary1ReportFridayMsg = "Off"
	If GenerateAutomaticCustomerAnalysisSummary1ReportFriday_ORIG = 1 then GenerateAutomaticCustomerAnalysisSummary1ReportFridayMsgOrig = "On" Else GenerateAutomaticCustomerAnalysisSummary1ReportFridayMsgOrig = "Off"
	
	If GenerateAutomaticCustomerAnalysisSummary1ReportFriday <> GenerateAutomaticCustomerAnalysisSummary1ReportFriday_ORIG Then
		CreateAuditLogEntry "Global Settings Change", "Global Settings Change", "Major", 1, "Automatic Customer Analysis Summary 1 Report generation schedule for Friday changed from " & GenerateAutomaticCustomerAnalysisSummary1ReportFridayMsgOrig & " to " & GenerateAutomaticCustomerAnalysisSummary1ReportFridayMsg
	End If
	
	If GenerateAutomaticCustomerAnalysisSummary1ReportFridayTime <> GenerateAutomaticCustomerAnalysisSummary1ReportFridayTime_ORIG Then
		If GenerateAutomaticCustomerAnalysisSummary1ReportFriday_ORIG = 0 AND GenerateAutomaticCustomerAnalysisSummary1ReportFriday = 1 Then
			CreateAuditLogEntry "Global Settings Change", "Global Settings Change", "Major", 1, "Automatic Customer Analysis Summary 1 Report generation schedule for Friday turned on and set to run at " & GenerateAutomaticCustomerAnalysisSummary1ReportFridayTime
		Else
			CreateAuditLogEntry "Global Settings Change", "Global Settings Change", "Major", 1, "Automatic Customer Analysis Summary 1 Report generation scheduled run time for Friday changed from " & GenerateAutomaticCustomerAnalysisSummary1ReportFridayTime_ORIG & " to " & GenerateAutomaticCustomerAnalysisSummary1ReportFridayTime
		End If
	End If



	If Request.Form("chkNoAutomaticCustomerAnalysisSummary1ReportSaturday") = "on" then GenerateAutomaticCustomerAnalysisSummary1ReportSaturdayMsg = "On" Else GenerateAutomaticCustomerAnalysisSummary1ReportSaturdayMsg = "Off"
	If GenerateAutomaticCustomerAnalysisSummary1ReportSaturday_ORIG = 1 then GenerateAutomaticCustomerAnalysisSummary1ReportSaturdayMsgOrig = "On" Else GenerateAutomaticCustomerAnalysisSummary1ReportSaturdayMsgOrig = "Off"
	
	If GenerateAutomaticCustomerAnalysisSummary1ReportSaturday <> GenerateAutomaticCustomerAnalysisSummary1ReportSaturday_ORIG Then
		CreateAuditLogEntry "Global Settings Change", "Global Settings Change", "Major", 1, "Automatic Customer Analysis Summary 1 Report generation schedule for Saturday changed from " & GenerateAutomaticCustomerAnalysisSummary1ReportSaturdayMsgOrig & " to " & GenerateAutomaticCustomerAnalysisSummary1ReportSaturdayMsg
	End If
	
	If GenerateAutomaticCustomerAnalysisSummary1ReportSaturdayTime <> GenerateAutomaticCustomerAnalysisSummary1ReportSaturdayTime_ORIG Then
		If GenerateAutomaticCustomerAnalysisSummary1ReportSaturday_ORIG = 0 AND GenerateAutomaticCustomerAnalysisSummary1ReportSaturday = 1 Then
			CreateAuditLogEntry "Global Settings Change", "Global Settings Change", "Major", 1, "Automatic Customer Analysis Summary 1 Report generation schedule for Saturday turned on and set to run at " & GenerateAutomaticCustomerAnalysisSummary1ReportSaturdayTime
		Else
			CreateAuditLogEntry "Global Settings Change", "Global Settings Change", "Major", 1, "Automatic Customer Analysis Summary 1 Report generation scheduled run time for Saturday changed from " & GenerateAutomaticCustomerAnalysisSummary1ReportSaturdayTime_ORIG & " to " & GenerateAutomaticCustomerAnalysisSummary1ReportSaturdayTime
		End If
	End If


	If Request.Form("chkNoAutomaticCustomerAnalysisSummary1ReportIfClosed") = "on" then RunAutomaticCustomerAnalysisSummary1ReportIfClosedMsg = "On" Else RunAutomaticCustomerAnalysisSummary1ReportIfClosedMsg = "Off"
	If RunAutomaticCustomerAnalysisSummary1ReportIfClosed_ORIG = 1 then RunAutomaticCustomerAnalysisSummary1ReportIfClosedMsgOrig = "On" Else RunAutomaticCustomerAnalysisSummary1ReportIfClosedMsgOrig = "Off"
	
	If RunAutomaticCustomerAnalysisSummary1ReportIfClosed <> RunAutomaticCustomerAnalysisSummary1ReportIfClosed_ORIG Then
		CreateAuditLogEntry "Global Settings Change", "Global Settings Change", "Major", 1, "Automatic Customer Analysis Summary 1 Report generation schedule - Don''t Run If Closed (Monday-Friday Only) -  changed from " & RunAutomaticCustomerAnalysisSummary1ReportIfClosedMsgOrig & " to " & RunAutomaticCustomerAnalysisSummary1ReportIfClosedMsg
	End If


	If Request.Form("chkNoAutomaticCustomerAnalysisSummary1ReportIfClosingEarly") = "on" then RunAutomaticCustomerAnalysisSummary1ReportIfClosingEarlyMsg = "On" Else RunAutomaticCustomerAnalysisSummary1ReportIfClosingEarlyMsg = "Off"
	If RunAutomaticCustomerAnalysisSummary1ReportIfClosingEarly_ORIG = 1 then RunAutomaticCustomerAnalysisSummary1ReportIfClosingEarlyMsgOrig = "On" Else RunAutomaticCustomerAnalysisSummary1ReportIfClosingEarlyMsgOrig = "Off"
	
	If RunAutomaticCustomerAnalysisSummary1ReportIfClosingEarly <> RunAutomaticCustomerAnalysisSummary1ReportIfClosingEarly_ORIG Then
		CreateAuditLogEntry "Global Settings Change", "Global Settings Change", "Major", 1, "Automatic Customer Analysis Summary 1 Report generation schedule - Don''t Run If Closed (Monday-Friday Only) -  changed from " & RunAutomaticCustomerAnalysisSummary1ReportIfClosingEarlyMsgOrig & " to " & RunAutomaticCustomerAnalysisSummary1ReportIfClosingEarlyMsg
	End If


	'*********************************************************************
	'Build Array/String of Schedule Data From Request Form Field Data
	'*********************************************************************

	Schedule_AutomaticCustomerAnalysisSummary1ReportGenerationUpdated = ""
	
	Schedule_AutomaticCustomerAnalysisSummary1ReportGenerationUpdated = GenerateAutomaticCustomerAnalysisSummary1ReportSunday
	Schedule_AutomaticCustomerAnalysisSummary1ReportGenerationUpdated = Schedule_AutomaticCustomerAnalysisSummary1ReportGenerationUpdated & "," & GenerateAutomaticCustomerAnalysisSummary1ReportMonday
	Schedule_AutomaticCustomerAnalysisSummary1ReportGenerationUpdated = Schedule_AutomaticCustomerAnalysisSummary1ReportGenerationUpdated & "," & GenerateAutomaticCustomerAnalysisSummary1ReportTuesday
	Schedule_AutomaticCustomerAnalysisSummary1ReportGenerationUpdated = Schedule_AutomaticCustomerAnalysisSummary1ReportGenerationUpdated & "," & GenerateAutomaticCustomerAnalysisSummary1ReportWednesday
	Schedule_AutomaticCustomerAnalysisSummary1ReportGenerationUpdated = Schedule_AutomaticCustomerAnalysisSummary1ReportGenerationUpdated & "," & GenerateAutomaticCustomerAnalysisSummary1ReportThursday
	Schedule_AutomaticCustomerAnalysisSummary1ReportGenerationUpdated = Schedule_AutomaticCustomerAnalysisSummary1ReportGenerationUpdated & "," & GenerateAutomaticCustomerAnalysisSummary1ReportFriday
	Schedule_AutomaticCustomerAnalysisSummary1ReportGenerationUpdated = Schedule_AutomaticCustomerAnalysisSummary1ReportGenerationUpdated & "," & GenerateAutomaticCustomerAnalysisSummary1ReportSaturday
	
	Schedule_AutomaticCustomerAnalysisSummary1ReportGenerationUpdated = Schedule_AutomaticCustomerAnalysisSummary1ReportGenerationUpdated & "," & GenerateAutomaticCustomerAnalysisSummary1ReportSundayTime
	Schedule_AutomaticCustomerAnalysisSummary1ReportGenerationUpdated = Schedule_AutomaticCustomerAnalysisSummary1ReportGenerationUpdated & "," & GenerateAutomaticCustomerAnalysisSummary1ReportMondayTime
	Schedule_AutomaticCustomerAnalysisSummary1ReportGenerationUpdated = Schedule_AutomaticCustomerAnalysisSummary1ReportGenerationUpdated & "," & GenerateAutomaticCustomerAnalysisSummary1ReportTuesdayTime
	Schedule_AutomaticCustomerAnalysisSummary1ReportGenerationUpdated = Schedule_AutomaticCustomerAnalysisSummary1ReportGenerationUpdated & "," & GenerateAutomaticCustomerAnalysisSummary1ReportWednesdayTime
	Schedule_AutomaticCustomerAnalysisSummary1ReportGenerationUpdated = Schedule_AutomaticCustomerAnalysisSummary1ReportGenerationUpdated & "," & GenerateAutomaticCustomerAnalysisSummary1ReportThursdayTime
	Schedule_AutomaticCustomerAnalysisSummary1ReportGenerationUpdated = Schedule_AutomaticCustomerAnalysisSummary1ReportGenerationUpdated & "," & GenerateAutomaticCustomerAnalysisSummary1ReportFridayTime
	Schedule_AutomaticCustomerAnalysisSummary1ReportGenerationUpdated = Schedule_AutomaticCustomerAnalysisSummary1ReportGenerationUpdated & "," & GenerateAutomaticCustomerAnalysisSummary1ReportSaturdayTime

	
	Schedule_AutomaticCustomerAnalysisSummary1ReportGenerationUpdated = Schedule_AutomaticCustomerAnalysisSummary1ReportGenerationUpdated & "," & RunAutomaticCustomerAnalysisSummary1ReportIfClosed
	Schedule_AutomaticCustomerAnalysisSummary1ReportGenerationUpdated = Schedule_AutomaticCustomerAnalysisSummary1ReportGenerationUpdated & "," & RunAutomaticCustomerAnalysisSummary1ReportIfClosingEarly
	
	Response.Write("<br><br><br>Schedule_AutomaticCustomerAnalysisSummary1ReportGenerationUpdated: " & Schedule_AutomaticCustomerAnalysisSummary1ReportGenerationUpdated)

	'*********************************************************************
	'Update SQL with Array/String of Schedule Data
	'*********************************************************************
		
	SQL = "UPDATE Settings_BizIntel SET Schedule_AutomaticCustomerAnalysisSummary1ReportGeneration = '" & cStr(Schedule_AutomaticCustomerAnalysisSummary1ReportGenerationUpdated) & "' "
	
	Response.Write("<br><br><br>SQL: " & SQL)
	
	Set cnn8 = Server.CreateObject("ADODB.Connection")
	cnn8.open (Session("ClientCnnString"))
	Set rs = Server.CreateObject("ADODB.Recordset")
	rs.CursorLocation = 3 
	Set rs = cnn8.Execute(SQL)
	
	set rs = Nothing
	cnn8.close
	set cnn8 = Nothing


	 Response.Redirect("bizintel.asp")
	
%><!--#include file="../../../inc/footer-main.asp"-->