<!--#include file="../../../inc/header.asp"-->

<%

	'***********************************************************
	'Get Values Of All Form Fields Posted
	'***********************************************************
	
	GenerateMCSActivityReportSunday = Request.Form("chkNoMCSActivityReportSunday")
	GenerateMCSActivityReportMonday = Request.Form("chkNoMCSActivityReportMonday")
	GenerateMCSActivityReportTuesday = Request.Form("chkNoMCSActivityReportTuesday")
	GenerateMCSActivityReportWednesday = Request.Form("chkNoMCSActivityReportWednesday")
	GenerateMCSActivityReportThursday = Request.Form("chkNoMCSActivityReportThursday")
	GenerateMCSActivityReportFriday = Request.Form("chkNoMCSActivityReportFriday")
	GenerateMCSActivityReportSaturday = Request.Form("chkNoMCSActivityReportSaturday")
	
	GenerateMCSActivityReportSundayTime = Request.Form("txtMCSActivityReportSchedulerSundayTime")
	GenerateMCSActivityReportMondayTime = Request.Form("txtMCSActivityReportSchedulerMondayTime")
	GenerateMCSActivityReportTuesdayTime = Request.Form("txtMCSActivityReportSchedulerTuesdayTime")
	GenerateMCSActivityReportWednesdayTime = Request.Form("txtMCSActivityReportSchedulerWednesdayTime")
	GenerateMCSActivityReportThursdayTime = Request.Form("txtMCSActivityReportSchedulerThursdayTime")
	GenerateMCSActivityReportFridayTime = Request.Form("txtMCSActivityReportSchedulerFridayTime")
	GenerateMCSActivityReportSaturdayTime = Request.Form("txtMCSActivityReportSchedulerSaturdayTime")
	
	RunMCSActivityReportIfClosed = Request.Form("chkNoMCSActivityReportIfClosed")
	RunMCSActivityReportIfClosingEarly = Request.Form("chkNoMCSActivityReportIfClosingEarly")


	If Request.Form("chkNoMCSActivityReportSunday") = "on" Then
		GenerateMCSActivityReportSunday = 0
		GenerateMCSActivityReportSundayTime = ""
	Else 
		GenerateMCSActivityReportSunday = 1
	End If

	If Request.Form("chkNoMCSActivityReportMonday") = "on" Then
		GenerateMCSActivityReportMonday = 0
		GenerateMCSActivityReportMondayTime = ""
	Else 
		GenerateMCSActivityReportMonday = 1
	End If

	If Request.Form("chkNoMCSActivityReportTuesday") = "on" Then
		GenerateMCSActivityReportTuesday = 0
		GenerateMCSActivityReportTuesdayTime = ""
	Else 
		GenerateMCSActivityReportTuesday = 1
	End If

	If Request.Form("chkNoMCSActivityReportWednesday") = "on" Then
		GenerateMCSActivityReportWednesday = 0
		GenerateMCSActivityReportWednesdayTime = ""
	Else 
		GenerateMCSActivityReportWednesday = 1
	End If

	If Request.Form("chkNoMCSActivityReportThursday") = "on" Then
		GenerateMCSActivityReportThursday = 0
		GenerateMCSActivityReportThursdayTime = ""
	Else 
		GenerateMCSActivityReportThursday = 1
	End If

	If Request.Form("chkNoMCSActivityReportFriday") = "on" Then
		GenerateMCSActivityReportFriday = 0
		GenerateMCSActivityReportFridayTime = ""
	Else 
		GenerateMCSActivityReportFriday = 1
	End If

	If Request.Form("chkNoMCSActivityReportSaturday") = "on" Then
		GenerateMCSActivityReportSaturday = 0
		GenerateMCSActivityReportSaturdayTime = ""
	Else 
		GenerateMCSActivityReportSaturday = 1
	End If

	If Request.Form("chkNoMCSActivityReportIfClosed") = "on" Then RunMCSActivityReportIfClosed = 0 Else RunMCSActivityReportIfClosed = 1
	If Request.Form("chkNoMCSActivityReportIfClosingEarly") = "on" Then RunMCSActivityReportIfClosingEarly = 0 Else RunMCSActivityReportIfClosingEarly = 1
	
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
	
		Schedule_MCSActivityReportGeneration = rsFieldServiceSettings("Schedule_MCSActivityReportGeneration")
		
		Schedule_MCSActivityReportGenerationSettings = Split(Schedule_MCSActivityReportGeneration,",")

		GenerateMCSActivityReportSunday_ORIG = cInt(Schedule_MCSActivityReportGenerationSettings(0))
		GenerateMCSActivityReportMonday_ORIG = cInt(Schedule_MCSActivityReportGenerationSettings(1))
		GenerateMCSActivityReportTuesday_ORIG = cInt(Schedule_MCSActivityReportGenerationSettings(2))
		GenerateMCSActivityReportWednesday_ORIG = cInt(Schedule_MCSActivityReportGenerationSettings(3))
		GenerateMCSActivityReportThursday_ORIG = cInt(Schedule_MCSActivityReportGenerationSettings(4))
		GenerateMCSActivityReportFriday_ORIG = cInt(Schedule_MCSActivityReportGenerationSettings(5))
		GenerateMCSActivityReportSaturday_ORIG = cInt(Schedule_MCSActivityReportGenerationSettings(6))
		GenerateMCSActivityReportSundayTime_ORIG = Schedule_MCSActivityReportGenerationSettings(7)
		GenerateMCSActivityReportMondayTime_ORIG = Schedule_MCSActivityReportGenerationSettings(8)
		GenerateMCSActivityReportTuesdayTime_ORIG = Schedule_MCSActivityReportGenerationSettings(9)
		GenerateMCSActivityReportWednesdayTime_ORIG = Schedule_MCSActivityReportGenerationSettings(10)
		GenerateMCSActivityReportThursdayTime_ORIG = Schedule_MCSActivityReportGenerationSettings(11)
		GenerateMCSActivityReportFridayTime_ORIG = Schedule_MCSActivityReportGenerationSettings(12)
		GenerateMCSActivityReportSaturdayTime_ORIG = Schedule_MCSActivityReportGenerationSettings(13)
		RunMCSActivityReportIfClosed_ORIG = cInt(Schedule_MCSActivityReportGenerationSettings(14))
		RunMCSActivityReportIfClosingEarly_ORIG = cInt(Schedule_MCSActivityReportGenerationSettings(15))
	
	End If
	
	set rsFieldServiceSettings = Nothing
	cnnFieldServiceSettings.close
	set cnnFieldServiceSettings = Nothing
	

	'***********************************************************
	'Perform Audit Trail Entry Inserts
	'***********************************************************

	
	If Request.Form("chkNoMCSActivityReportSunday") = "on" then GenerateMCSActivityReportSundayMsg = "On" Else GenerateMCSActivityReportSundayMsg = "Off"
	If GenerateMCSActivityReportSunday_ORIG = 1 then GenerateMCSActivityReportSundayMsgOrig = "On" Else GenerateMCSActivityReportSundayMsgOrig = "Off"
	
	If GenerateMCSActivityReportSunday <> GenerateMCSActivityReportSunday_ORIG Then
		CreateAuditLogEntry "Global Settings Change", "Global Settings Change", "Major", 1, "MCS Activity Report generation schedule for SUNDAY changed from " & GenerateMCSActivityReportSundayMsgOrig & " to " & GenerateMCSActivityReportSundayMsg
	End If
	
	If GenerateMCSActivityReportSundayTime <> GenerateMCSActivityReportSundayTime_ORIG Then
		If GenerateMCSActivityReportSunday_ORIG = 0 AND GenerateMCSActivityReportSunday = 1 Then
			CreateAuditLogEntry "Global Settings Change", "Global Settings Change", "Major", 1, "MCS Activity Report generation schedule for SUNDAY turned on and set to run at " & GenerateMCSActivityReportSundayTime
		Else
			CreateAuditLogEntry "Global Settings Change", "Global Settings Change", "Major", 1, "MCS Activity Report generation scheduled run time for SUNDAY changed from " & GenerateMCSActivityReportSundayTime_ORIG & " to " & GenerateMCSActivityReportSundayTime
		End If
	End If
	
	

	If Request.Form("chkNoMCSActivityReportMonday") = "on" then GenerateMCSActivityReportMondayMsg = "On" Else GenerateMCSActivityReportMondayMsg = "Off"
	If GenerateMCSActivityReportMonday_ORIG = 1 then GenerateMCSActivityReportMondayMsgOrig = "On" Else GenerateMCSActivityReportMondayMsgOrig = "Off"
	
	If GenerateMCSActivityReportMonday <> GenerateMCSActivityReportMonday_ORIG Then
		CreateAuditLogEntry "Global Settings Change", "Global Settings Change", "Major", 1, "MCS Activity Report generation schedule for Monday changed from " & GenerateMCSActivityReportMondayMsgOrig & " to " & GenerateMCSActivityReportMondayMsg
	End If
	
	If GenerateMCSActivityReportMondayTime <> GenerateMCSActivityReportMondayTime_ORIG Then
		If GenerateMCSActivityReportMonday_ORIG = 0 AND GenerateMCSActivityReportMonday = 1 Then
			CreateAuditLogEntry "Global Settings Change", "Global Settings Change", "Major", 1, "MCS Activity Report generation schedule for Monday turned on and set to run at " & GenerateMCSActivityReportMondayTime
		Else
			CreateAuditLogEntry "Global Settings Change", "Global Settings Change", "Major", 1, "MCS Activity Report generation scheduled run time for Monday changed from " & GenerateMCSActivityReportMondayTime_ORIG & " to " & GenerateMCSActivityReportMondayTime
		End If
	End If
	


	If Request.Form("chkNoMCSActivityReportTuesday") = "on" then GenerateMCSActivityReportTuesdayMsg = "On" Else GenerateMCSActivityReportTuesdayMsg = "Off"
	If GenerateMCSActivityReportTuesday_ORIG = 1 then GenerateMCSActivityReportTuesdayMsgOrig = "On" Else GenerateMCSActivityReportTuesdayMsgOrig = "Off"
	
	If GenerateMCSActivityReportTuesday <> GenerateMCSActivityReportTuesday_ORIG Then
		CreateAuditLogEntry "Global Settings Change", "Global Settings Change", "Major", 1, "MCS Activity Report generation schedule for Tuesday changed from " & GenerateMCSActivityReportTuesdayMsgOrig & " to " & GenerateMCSActivityReportTuesdayMsg
	End If
	
	If GenerateMCSActivityReportTuesdayTime <> GenerateMCSActivityReportTuesdayTime_ORIG Then
		If GenerateMCSActivityReportTuesday_ORIG = 0 AND GenerateMCSActivityReportTuesday = 1 Then
			CreateAuditLogEntry "Global Settings Change", "Global Settings Change", "Major", 1, "MCS Activity Report generation schedule for Tuesday turned on and set to run at " & GenerateMCSActivityReportTuesdayTime
		Else
			CreateAuditLogEntry "Global Settings Change", "Global Settings Change", "Major", 1, "MCS Activity Report generation scheduled run time for Tuesday changed from " & GenerateMCSActivityReportTuesdayTime_ORIG & " to " & GenerateMCSActivityReportTuesdayTime
		End If
	End If



	If Request.Form("chkNoMCSActivityReportWednesday") = "on" then GenerateMCSActivityReportWednesdayMsg = "On" Else GenerateMCSActivityReportWednesdayMsg = "Off"
	If GenerateMCSActivityReportWednesday_ORIG = 1 then GenerateMCSActivityReportWednesdayMsgOrig = "On" Else GenerateMCSActivityReportWednesdayMsgOrig = "Off"
	
	If GenerateMCSActivityReportWednesday <> GenerateMCSActivityReportWednesday_ORIG Then
		CreateAuditLogEntry "Global Settings Change", "Global Settings Change", "Major", 1, "MCS Activity Report generation schedule for Wednesday changed from " & GenerateMCSActivityReportWednesdayMsgOrig & " to " & GenerateMCSActivityReportWednesdayMsg
	End If
	
	If GenerateMCSActivityReportWednesdayTime <> GenerateMCSActivityReportWednesdayTime_ORIG Then
		If GenerateMCSActivityReportWednesday_ORIG = 0 AND GenerateMCSActivityReportWednesday = 1 Then
			CreateAuditLogEntry "Global Settings Change", "Global Settings Change", "Major", 1, "MCS Activity Report generation schedule for Wednesday turned on and set to run at " & GenerateMCSActivityReportWednesdayTime
		Else
			CreateAuditLogEntry "Global Settings Change", "Global Settings Change", "Major", 1, "MCS Activity Report generation scheduled run time for Wednesday changed from " & GenerateMCSActivityReportWednesdayTime_ORIG & " to " & GenerateMCSActivityReportWednesdayTime
		End If
	End If



	If Request.Form("chkNoMCSActivityReportThursday") = "on" then GenerateMCSActivityReportThursdayMsg = "On" Else GenerateMCSActivityReportThursdayMsg = "Off"
	If GenerateMCSActivityReportThursday_ORIG = 1 then GenerateMCSActivityReportThursdayMsgOrig = "On" Else GenerateMCSActivityReportThursdayMsgOrig = "Off"
	
	If GenerateMCSActivityReportThursday <> GenerateMCSActivityReportThursday_ORIG Then
		CreateAuditLogEntry "Global Settings Change", "Global Settings Change", "Major", 1, "MCS Activity Report generation schedule for Thursday changed from " & GenerateMCSActivityReportThursdayMsgOrig & " to " & GenerateMCSActivityReportThursdayMsg
	End If
	
	If GenerateMCSActivityReportThursdayTime <> GenerateMCSActivityReportThursdayTime_ORIG Then
		If GenerateMCSActivityReportThursday_ORIG = 0 AND GenerateMCSActivityReportThursday = 1 Then
			CreateAuditLogEntry "Global Settings Change", "Global Settings Change", "Major", 1, "MCS Activity Report generation schedule for Thursday turned on and set to run at " & GenerateMCSActivityReportThursdayTime
		Else
			CreateAuditLogEntry "Global Settings Change", "Global Settings Change", "Major", 1, "MCS Activity Report generation scheduled run time for Thursday changed from " & GenerateMCSActivityReportThursdayTime_ORIG & " to " & GenerateMCSActivityReportThursdayTime
		End If
	End If



	If Request.Form("chkNoMCSActivityReportFriday") = "on" then GenerateMCSActivityReportFridayMsg = "On" Else GenerateMCSActivityReportFridayMsg = "Off"
	If GenerateMCSActivityReportFriday_ORIG = 1 then GenerateMCSActivityReportFridayMsgOrig = "On" Else GenerateMCSActivityReportFridayMsgOrig = "Off"
	
	If GenerateMCSActivityReportFriday <> GenerateMCSActivityReportFriday_ORIG Then
		CreateAuditLogEntry "Global Settings Change", "Global Settings Change", "Major", 1, "MCS Activity Report generation schedule for Friday changed from " & GenerateMCSActivityReportFridayMsgOrig & " to " & GenerateMCSActivityReportFridayMsg
	End If
	
	If GenerateMCSActivityReportFridayTime <> GenerateMCSActivityReportFridayTime_ORIG Then
		If GenerateMCSActivityReportFriday_ORIG = 0 AND GenerateMCSActivityReportFriday = 1 Then
			CreateAuditLogEntry "Global Settings Change", "Global Settings Change", "Major", 1, "MCS Activity Report generation schedule for Friday turned on and set to run at " & GenerateMCSActivityReportFridayTime
		Else
			CreateAuditLogEntry "Global Settings Change", "Global Settings Change", "Major", 1, "MCS Activity Report generation scheduled run time for Friday changed from " & GenerateMCSActivityReportFridayTime_ORIG & " to " & GenerateMCSActivityReportFridayTime
		End If
	End If



	If Request.Form("chkNoMCSActivityReportSaturday") = "on" then GenerateMCSActivityReportSaturdayMsg = "On" Else GenerateMCSActivityReportSaturdayMsg = "Off"
	If GenerateMCSActivityReportSaturday_ORIG = 1 then GenerateMCSActivityReportSaturdayMsgOrig = "On" Else GenerateMCSActivityReportSaturdayMsgOrig = "Off"
	
	If GenerateMCSActivityReportSaturday <> GenerateMCSActivityReportSaturday_ORIG Then
		CreateAuditLogEntry "Global Settings Change", "Global Settings Change", "Major", 1, "MCS Activity Report generation schedule for Saturday changed from " & GenerateMCSActivityReportSaturdayMsgOrig & " to " & GenerateMCSActivityReportSaturdayMsg
	End If
	
	If GenerateMCSActivityReportSaturdayTime <> GenerateMCSActivityReportSaturdayTime_ORIG Then
		If GenerateMCSActivityReportSaturday_ORIG = 0 AND GenerateMCSActivityReportSaturday = 1 Then
			CreateAuditLogEntry "Global Settings Change", "Global Settings Change", "Major", 1, "MCS Activity Report generation schedule for Saturday turned on and set to run at " & GenerateMCSActivityReportSaturdayTime
		Else
			CreateAuditLogEntry "Global Settings Change", "Global Settings Change", "Major", 1, "MCS Activity Report generation scheduled run time for Saturday changed from " & GenerateMCSActivityReportSaturdayTime_ORIG & " to " & GenerateMCSActivityReportSaturdayTime
		End If
	End If


	If Request.Form("chkNoMCSActivityReportIfClosed") = "on" then RunMCSActivityReportIfClosedMsg = "On" Else RunMCSActivityReportIfClosedMsg = "Off"
	If RunMCSActivityReportIfClosed_ORIG = 1 then RunMCSActivityReportIfClosedMsgOrig = "On" Else RunMCSActivityReportIfClosedMsgOrig = "Off"
	
	If RunMCSActivityReportIfClosed <> RunMCSActivityReportIfClosed_ORIG Then
		CreateAuditLogEntry "Global Settings Change", "Global Settings Change", "Major", 1, "MCS Activity Report generation schedule - Don''t Run If Closed (Monday-Friday Only) -  changed from " & RunMCSActivityReportIfClosedMsgOrig & " to " & RunMCSActivityReportIfClosedMsg
	End If


	If Request.Form("chkNoMCSActivityReportIfClosingEarly") = "on" then RunMCSActivityReportIfClosingEarlyMsg = "On" Else RunMCSActivityReportIfClosingEarlyMsg = "Off"
	If RunMCSActivityReportIfClosingEarly_ORIG = 1 then RunMCSActivityReportIfClosingEarlyMsgOrig = "On" Else RunMCSActivityReportIfClosingEarlyMsgOrig = "Off"
	
	If RunMCSActivityReportIfClosingEarly <> RunMCSActivityReportIfClosingEarly_ORIG Then
		CreateAuditLogEntry "Global Settings Change", "Global Settings Change", "Major", 1, "MCS Activity Report generation schedule - Don''t Run If Closed (Monday-Friday Only) -  changed from " & RunMCSActivityReportIfClosingEarlyMsgOrig & " to " & RunMCSActivityReportIfClosingEarlyMsg
	End If


	'*********************************************************************
	'Build Array/String of Schedule Data From Request Form Field Data
	'*********************************************************************

	Schedule_MCSActivityReportGenerationUpdated = ""
	
	Schedule_MCSActivityReportGenerationUpdated = GenerateMCSActivityReportSunday
	Schedule_MCSActivityReportGenerationUpdated = Schedule_MCSActivityReportGenerationUpdated & "," & GenerateMCSActivityReportMonday
	Schedule_MCSActivityReportGenerationUpdated = Schedule_MCSActivityReportGenerationUpdated & "," & GenerateMCSActivityReportTuesday
	Schedule_MCSActivityReportGenerationUpdated = Schedule_MCSActivityReportGenerationUpdated & "," & GenerateMCSActivityReportWednesday
	Schedule_MCSActivityReportGenerationUpdated = Schedule_MCSActivityReportGenerationUpdated & "," & GenerateMCSActivityReportThursday
	Schedule_MCSActivityReportGenerationUpdated = Schedule_MCSActivityReportGenerationUpdated & "," & GenerateMCSActivityReportFriday
	Schedule_MCSActivityReportGenerationUpdated = Schedule_MCSActivityReportGenerationUpdated & "," & GenerateMCSActivityReportSaturday
	
	Schedule_MCSActivityReportGenerationUpdated = Schedule_MCSActivityReportGenerationUpdated & "," & GenerateMCSActivityReportSundayTime
	Schedule_MCSActivityReportGenerationUpdated = Schedule_MCSActivityReportGenerationUpdated & "," & GenerateMCSActivityReportMondayTime
	Schedule_MCSActivityReportGenerationUpdated = Schedule_MCSActivityReportGenerationUpdated & "," & GenerateMCSActivityReportTuesdayTime
	Schedule_MCSActivityReportGenerationUpdated = Schedule_MCSActivityReportGenerationUpdated & "," & GenerateMCSActivityReportWednesdayTime
	Schedule_MCSActivityReportGenerationUpdated = Schedule_MCSActivityReportGenerationUpdated & "," & GenerateMCSActivityReportThursdayTime
	Schedule_MCSActivityReportGenerationUpdated = Schedule_MCSActivityReportGenerationUpdated & "," & GenerateMCSActivityReportFridayTime
	Schedule_MCSActivityReportGenerationUpdated = Schedule_MCSActivityReportGenerationUpdated & "," & GenerateMCSActivityReportSaturdayTime

	
	Schedule_MCSActivityReportGenerationUpdated = Schedule_MCSActivityReportGenerationUpdated & "," & RunMCSActivityReportIfClosed
	Schedule_MCSActivityReportGenerationUpdated = Schedule_MCSActivityReportGenerationUpdated & "," & RunMCSActivityReportIfClosingEarly
	
	Response.Write("<br><br><br>Schedule_MCSActivityReportGenerationUpdated: " & Schedule_MCSActivityReportGenerationUpdated)

	'*********************************************************************
	'Update SQL with Array/String of Schedule Data
	'*********************************************************************
		
	SQL = "UPDATE Settings_BizIntel SET Schedule_MCSActivityReportGeneration = '" & cStr(Schedule_MCSActivityReportGenerationUpdated) & "' "
	
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