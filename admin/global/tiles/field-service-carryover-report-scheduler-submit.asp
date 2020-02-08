<!--#include file="../../../inc/header.asp"-->

<%

	'***********************************************************
	'Get Values Of All Form Fields Posted 
	'***********************************************************
	
	GenerateServiceTicketCarryoverReportSunday = Request.Form("chkNoServiceTicketCarryoverReportSunday")
	GenerateServiceTicketCarryoverReportMonday = Request.Form("chkNoServiceTicketCarryoverReportMonday")
	GenerateServiceTicketCarryoverReportTuesday = Request.Form("chkNoServiceTicketCarryoverReportTuesday")
	GenerateServiceTicketCarryoverReportWednesday = Request.Form("chkNoServiceTicketCarryoverReportWednesday")
	GenerateServiceTicketCarryoverReportThursday = Request.Form("chkNoServiceTicketCarryoverReportThursday")
	GenerateServiceTicketCarryoverReportFriday = Request.Form("chkNoServiceTicketCarryoverReportFriday")
	GenerateServiceTicketCarryoverReportSaturday = Request.Form("chkNoServiceTicketCarryoverReportSaturday")
	
	GenerateServiceTicketCarryoverReportSundayTime = Request.Form("txtServiceTicketCarryoverReportSchedulerSundayTime")
	GenerateServiceTicketCarryoverReportMondayTime = Request.Form("txtServiceTicketCarryoverReportSchedulerMondayTime")
	GenerateServiceTicketCarryoverReportTuesdayTime = Request.Form("txtServiceTicketCarryoverReportSchedulerTuesdayTime")
	GenerateServiceTicketCarryoverReportWednesdayTime = Request.Form("txtServiceTicketCarryoverReportSchedulerWednesdayTime")
	GenerateServiceTicketCarryoverReportThursdayTime = Request.Form("txtServiceTicketCarryoverReportSchedulerThursdayTime")
	GenerateServiceTicketCarryoverReportFridayTime = Request.Form("txtServiceTicketCarryoverReportSchedulerFridayTime")
	GenerateServiceTicketCarryoverReportSaturdayTime = Request.Form("txtServiceTicketCarryoverReportSchedulerSaturdayTime")
	
	RunServiceTicketCarryoverReportIfClosed = Request.Form("chkNoServiceTicketCarryoverReportIfClosed")
	RunServiceTicketCarryoverReportIfClosingEarly = Request.Form("chkNoServiceTicketCarryoverReportIfClosingEarly")


	If Request.Form("chkNoServiceTicketCarryoverReportSunday") = "on" Then
		GenerateServiceTicketCarryoverReportSunday = 0
		GenerateServiceTicketCarryoverReportSundayTime = ""
	Else 
		GenerateServiceTicketCarryoverReportSunday = 1
	End If

	If Request.Form("chkNoServiceTicketCarryoverReportMonday") = "on" Then
		GenerateServiceTicketCarryoverReportMonday = 0
		GenerateServiceTicketCarryoverReportMondayTime = ""
	Else 
		GenerateServiceTicketCarryoverReportMonday = 1
	End If

	If Request.Form("chkNoServiceTicketCarryoverReportTuesday") = "on" Then
		GenerateServiceTicketCarryoverReportTuesday = 0
		GenerateServiceTicketCarryoverReportTuesdayTime = ""
	Else 
		GenerateServiceTicketCarryoverReportTuesday = 1
	End If

	If Request.Form("chkNoServiceTicketCarryoverReportWednesday") = "on" Then
		GenerateServiceTicketCarryoverReportWednesday = 0
		GenerateServiceTicketCarryoverReportWednesdayTime = ""
	Else 
		GenerateServiceTicketCarryoverReportWednesday = 1
	End If

	If Request.Form("chkNoServiceTicketCarryoverReportThursday") = "on" Then
		GenerateServiceTicketCarryoverReportThursday = 0
		GenerateServiceTicketCarryoverReportThursdayTime = ""
	Else 
		GenerateServiceTicketCarryoverReportThursday = 1
	End If

	If Request.Form("chkNoServiceTicketCarryoverReportFriday") = "on" Then
		GenerateServiceTicketCarryoverReportFriday = 0
		GenerateServiceTicketCarryoverReportFridayTime = ""
	Else 
		GenerateServiceTicketCarryoverReportFriday = 1
	End If

	If Request.Form("chkNoServiceTicketCarryoverReportSaturday") = "on" Then
		GenerateServiceTicketCarryoverReportSaturday = 0
		GenerateServiceTicketCarryoverReportSaturdayTime = ""
	Else 
		GenerateServiceTicketCarryoverReportSaturday = 1
	End If

	If Request.Form("chkNoServiceTicketCarryoverReportIfClosed") = "on" Then RunServiceTicketCarryoverReportIfClosed = 0 Else RunServiceTicketCarryoverReportIfClosed = 1
	If Request.Form("chkNoServiceTicketCarryoverReportIfClosingEarly") = "on" Then RunServiceTicketCarryoverReportIfClosingEarly = 0 Else RunServiceTicketCarryoverReportIfClosingEarly = 1
	
	'***********************************************************
	'Get Original Values For Audit Trail Entries
	'***********************************************************
	
	SQLFieldServiceSettings = "SELECT * FROM Settings_FieldService"
	
	Set cnnFieldServiceSettings = Server.CreateObject("ADODB.Connection")
	cnnFieldServiceSettings.open (Session("ClientCnnString"))
	Set rsFieldServiceSettings = Server.CreateObject("ADODB.Recordset")
	rsFieldServiceSettings.CursorLocation = 3 
	Set rsFieldServiceSettings = cnnFieldServiceSettings.Execute(SQLFieldServiceSettings)
		
	If NOT rsFieldServiceSettings.EOF Then
	
		Schedule_ServiceTicketCarryoverReportGeneration = rsFieldServiceSettings("Schedule_ServiceTicketCarryoverReportGeneration")
		
		Schedule_ServiceTicketCarryoverReportGenerationSettings = Split(Schedule_ServiceTicketCarryoverReportGeneration,",")

		GenerateServiceTicketCarryoverReportSunday_ORIG = cInt(Schedule_ServiceTicketCarryoverReportGenerationSettings(0))
		GenerateServiceTicketCarryoverReportMonday_ORIG = cInt(Schedule_ServiceTicketCarryoverReportGenerationSettings(1))
		GenerateServiceTicketCarryoverReportTuesday_ORIG = cInt(Schedule_ServiceTicketCarryoverReportGenerationSettings(2))
		GenerateServiceTicketCarryoverReportWednesday_ORIG = cInt(Schedule_ServiceTicketCarryoverReportGenerationSettings(3))
		GenerateServiceTicketCarryoverReportThursday_ORIG = cInt(Schedule_ServiceTicketCarryoverReportGenerationSettings(4))
		GenerateServiceTicketCarryoverReportFriday_ORIG = cInt(Schedule_ServiceTicketCarryoverReportGenerationSettings(5))
		GenerateServiceTicketCarryoverReportSaturday_ORIG = cInt(Schedule_ServiceTicketCarryoverReportGenerationSettings(6))
		GenerateServiceTicketCarryoverReportSundayTime_ORIG = Schedule_ServiceTicketCarryoverReportGenerationSettings(7)
		GenerateServiceTicketCarryoverReportMondayTime_ORIG = Schedule_ServiceTicketCarryoverReportGenerationSettings(8)
		GenerateServiceTicketCarryoverReportTuesdayTime_ORIG = Schedule_ServiceTicketCarryoverReportGenerationSettings(9)
		GenerateServiceTicketCarryoverReportWednesdayTime_ORIG = Schedule_ServiceTicketCarryoverReportGenerationSettings(10)
		GenerateServiceTicketCarryoverReportThursdayTime_ORIG = Schedule_ServiceTicketCarryoverReportGenerationSettings(11)
		GenerateServiceTicketCarryoverReportFridayTime_ORIG = Schedule_ServiceTicketCarryoverReportGenerationSettings(12)
		GenerateServiceTicketCarryoverReportSaturdayTime_ORIG = Schedule_ServiceTicketCarryoverReportGenerationSettings(13)
		RunServiceTicketCarryoverReportIfClosed_ORIG = cInt(Schedule_ServiceTicketCarryoverReportGenerationSettings(14))
		RunServiceTicketCarryoverReportIfClosingEarly_ORIG = cInt(Schedule_ServiceTicketCarryoverReportGenerationSettings(15))
	
	End If
	
	set rsFieldServiceSettings = Nothing
	cnnFieldServiceSettings.close
	set cnnFieldServiceSettings = Nothing
	

	'***********************************************************
	'Perform Audit Trail Entry Inserts
	'***********************************************************

	
	If Request.Form("chkNoServiceTicketCarryoverReportSunday") = "on" then GenerateServiceTicketCarryoverReportSundayMsg = "On" Else GenerateServiceTicketCarryoverReportSundayMsg = "Off"
	If GenerateServiceTicketCarryoverReportSunday_ORIG = 1 then GenerateServiceTicketCarryoverReportSundayMsgOrig = "On" Else GenerateServiceTicketCarryoverReportSundayMsgOrig = "Off"
	
	If GenerateServiceTicketCarryoverReportSunday <> GenerateServiceTicketCarryoverReportSunday_ORIG Then
		CreateAuditLogEntry "Global Settings Change", "Global Settings Change", "Major", 1, "Service ticket carryover report schedule for SUNDAY changed from " & GenerateServiceTicketCarryoverReportSundayMsgOrig & " to " & GenerateServiceTicketCarryoverReportSundayMsg
	End If
	
	If GenerateServiceTicketCarryoverReportSundayTime <> GenerateServiceTicketCarryoverReportSundayTime_ORIG Then
		If GenerateServiceTicketCarryoverReportSunday_ORIG = 0 AND GenerateServiceTicketCarryoverReportSunday = 1 Then
			CreateAuditLogEntry "Global Settings Change", "Global Settings Change", "Major", 1, "Service ticket carryover report schedule for SUNDAY turned on and set to run at " & GenerateServiceTicketCarryoverReportSundayTime
		Else
			CreateAuditLogEntry "Global Settings Change", "Global Settings Change", "Major", 1, "Service ticket carryover report scheduled run time for SUNDAY changed from " & GenerateServiceTicketCarryoverReportSundayTime_ORIG & " to " & GenerateServiceTicketCarryoverReportSundayTime
		End If
	End If
	
	

	If Request.Form("chkNoServiceTicketCarryoverReportMonday") = "on" then GenerateServiceTicketCarryoverReportMondayMsg = "On" Else GenerateServiceTicketCarryoverReportMondayMsg = "Off"
	If GenerateServiceTicketCarryoverReportMonday_ORIG = 1 then GenerateServiceTicketCarryoverReportMondayMsgOrig = "On" Else GenerateServiceTicketCarryoverReportMondayMsgOrig = "Off"
	
	If GenerateServiceTicketCarryoverReportMonday <> GenerateServiceTicketCarryoverReportMonday_ORIG Then
		CreateAuditLogEntry "Global Settings Change", "Global Settings Change", "Major", 1, "Service ticket carryover report schedule for Monday changed from " & GenerateServiceTicketCarryoverReportMondayMsgOrig & " to " & GenerateServiceTicketCarryoverReportMondayMsg
	End If
	
	If GenerateServiceTicketCarryoverReportMondayTime <> GenerateServiceTicketCarryoverReportMondayTime_ORIG Then
		If GenerateServiceTicketCarryoverReportMonday_ORIG = 0 AND GenerateServiceTicketCarryoverReportMonday = 1 Then
			CreateAuditLogEntry "Global Settings Change", "Global Settings Change", "Major", 1, "Service ticket carryover report schedule for Monday turned on and set to run at " & GenerateServiceTicketCarryoverReportMondayTime
		Else
			CreateAuditLogEntry "Global Settings Change", "Global Settings Change", "Major", 1, "Service ticket carryover report scheduled run time for Monday changed from " & GenerateServiceTicketCarryoverReportMondayTime_ORIG & " to " & GenerateServiceTicketCarryoverReportMondayTime
		End If
	End If
	


	If Request.Form("chkNoServiceTicketCarryoverReportTuesday") = "on" then GenerateServiceTicketCarryoverReportTuesdayMsg = "On" Else GenerateServiceTicketCarryoverReportTuesdayMsg = "Off"
	If GenerateServiceTicketCarryoverReportTuesday_ORIG = 1 then GenerateServiceTicketCarryoverReportTuesdayMsgOrig = "On" Else GenerateServiceTicketCarryoverReportTuesdayMsgOrig = "Off"
	
	If GenerateServiceTicketCarryoverReportTuesday <> GenerateServiceTicketCarryoverReportTuesday_ORIG Then
		CreateAuditLogEntry "Global Settings Change", "Global Settings Change", "Major", 1, "Service ticket carryover report schedule for Tuesday changed from " & GenerateServiceTicketCarryoverReportTuesdayMsgOrig & " to " & GenerateServiceTicketCarryoverReportTuesdayMsg
	End If
	
	If GenerateServiceTicketCarryoverReportTuesdayTime <> GenerateServiceTicketCarryoverReportTuesdayTime_ORIG Then
		If GenerateServiceTicketCarryoverReportTuesday_ORIG = 0 AND GenerateServiceTicketCarryoverReportTuesday = 1 Then
			CreateAuditLogEntry "Global Settings Change", "Global Settings Change", "Major", 1, "Service ticket carryover report schedule for Tuesday turned on and set to run at " & GenerateServiceTicketCarryoverReportTuesdayTime
		Else
			CreateAuditLogEntry "Global Settings Change", "Global Settings Change", "Major", 1, "Service ticket carryover report scheduled run time for Tuesday changed from " & GenerateServiceTicketCarryoverReportTuesdayTime_ORIG & " to " & GenerateServiceTicketCarryoverReportTuesdayTime
		End If
	End If



	If Request.Form("chkNoServiceTicketCarryoverReportWednesday") = "on" then GenerateServiceTicketCarryoverReportWednesdayMsg = "On" Else GenerateServiceTicketCarryoverReportWednesdayMsg = "Off"
	If GenerateServiceTicketCarryoverReportWednesday_ORIG = 1 then GenerateServiceTicketCarryoverReportWednesdayMsgOrig = "On" Else GenerateServiceTicketCarryoverReportWednesdayMsgOrig = "Off"
	
	If GenerateServiceTicketCarryoverReportWednesday <> GenerateServiceTicketCarryoverReportWednesday_ORIG Then
		CreateAuditLogEntry "Global Settings Change", "Global Settings Change", "Major", 1, "Service ticket carryover report schedule for Wednesday changed from " & GenerateServiceTicketCarryoverReportWednesdayMsgOrig & " to " & GenerateServiceTicketCarryoverReportWednesdayMsg
	End If
	
	If GenerateServiceTicketCarryoverReportWednesdayTime <> GenerateServiceTicketCarryoverReportWednesdayTime_ORIG Then
		If GenerateServiceTicketCarryoverReportWednesday_ORIG = 0 AND GenerateServiceTicketCarryoverReportWednesday = 1 Then
			CreateAuditLogEntry "Global Settings Change", "Global Settings Change", "Major", 1, "Service ticket carryover report schedule for Wednesday turned on and set to run at " & GenerateServiceTicketCarryoverReportWednesdayTime
		Else
			CreateAuditLogEntry "Global Settings Change", "Global Settings Change", "Major", 1, "Service ticket carryover report scheduled run time for Wednesday changed from " & GenerateServiceTicketCarryoverReportWednesdayTime_ORIG & " to " & GenerateServiceTicketCarryoverReportWednesdayTime
		End If
	End If



	If Request.Form("chkNoServiceTicketCarryoverReportThursday") = "on" then GenerateServiceTicketCarryoverReportThursdayMsg = "On" Else GenerateServiceTicketCarryoverReportThursdayMsg = "Off"
	If GenerateServiceTicketCarryoverReportThursday_ORIG = 1 then GenerateServiceTicketCarryoverReportThursdayMsgOrig = "On" Else GenerateServiceTicketCarryoverReportThursdayMsgOrig = "Off"
	
	If GenerateServiceTicketCarryoverReportThursday <> GenerateServiceTicketCarryoverReportThursday_ORIG Then
		CreateAuditLogEntry "Global Settings Change", "Global Settings Change", "Major", 1, "Service ticket carryover report schedule for Thursday changed from " & GenerateServiceTicketCarryoverReportThursdayMsgOrig & " to " & GenerateServiceTicketCarryoverReportThursdayMsg
	End If
	
	If GenerateServiceTicketCarryoverReportThursdayTime <> GenerateServiceTicketCarryoverReportThursdayTime_ORIG Then
		If GenerateServiceTicketCarryoverReportThursday_ORIG = 0 AND GenerateServiceTicketCarryoverReportThursday = 1 Then
			CreateAuditLogEntry "Global Settings Change", "Global Settings Change", "Major", 1, "Service ticket carryover report schedule for Thursday turned on and set to run at " & GenerateServiceTicketCarryoverReportThursdayTime
		Else
			CreateAuditLogEntry "Global Settings Change", "Global Settings Change", "Major", 1, "Service ticket carryover report scheduled run time for Thursday changed from " & GenerateServiceTicketCarryoverReportThursdayTime_ORIG & " to " & GenerateServiceTicketCarryoverReportThursdayTime
		End If
	End If



	If Request.Form("chkNoServiceTicketCarryoverReportFriday") = "on" then GenerateServiceTicketCarryoverReportFridayMsg = "On" Else GenerateServiceTicketCarryoverReportFridayMsg = "Off"
	If GenerateServiceTicketCarryoverReportFriday_ORIG = 1 then GenerateServiceTicketCarryoverReportFridayMsgOrig = "On" Else GenerateServiceTicketCarryoverReportFridayMsgOrig = "Off"
	
	If GenerateServiceTicketCarryoverReportFriday <> GenerateServiceTicketCarryoverReportFriday_ORIG Then
		CreateAuditLogEntry "Global Settings Change", "Global Settings Change", "Major", 1, "Service ticket carryover report schedule for Friday changed from " & GenerateServiceTicketCarryoverReportFridayMsgOrig & " to " & GenerateServiceTicketCarryoverReportFridayMsg
	End If
	
	If GenerateServiceTicketCarryoverReportFridayTime <> GenerateServiceTicketCarryoverReportFridayTime_ORIG Then
		If GenerateServiceTicketCarryoverReportFriday_ORIG = 0 AND GenerateServiceTicketCarryoverReportFriday = 1 Then
			CreateAuditLogEntry "Global Settings Change", "Global Settings Change", "Major", 1, "Service ticket carryover report schedule for Friday turned on and set to run at " & GenerateServiceTicketCarryoverReportFridayTime
		Else
			CreateAuditLogEntry "Global Settings Change", "Global Settings Change", "Major", 1, "Service ticket carryover report scheduled run time for Friday changed from " & GenerateServiceTicketCarryoverReportFridayTime_ORIG & " to " & GenerateServiceTicketCarryoverReportFridayTime
		End If
	End If



	If Request.Form("chkNoServiceTicketCarryoverReportSaturday") = "on" then GenerateServiceTicketCarryoverReportSaturdayMsg = "On" Else GenerateServiceTicketCarryoverReportSaturdayMsg = "Off"
	If GenerateServiceTicketCarryoverReportSaturday_ORIG = 1 then GenerateServiceTicketCarryoverReportSaturdayMsgOrig = "On" Else GenerateServiceTicketCarryoverReportSaturdayMsgOrig = "Off"
	
	If GenerateServiceTicketCarryoverReportSaturday <> GenerateServiceTicketCarryoverReportSaturday_ORIG Then
		CreateAuditLogEntry "Global Settings Change", "Global Settings Change", "Major", 1, "Service ticket carryover report schedule for Saturday changed from " & GenerateServiceTicketCarryoverReportSaturdayMsgOrig & " to " & GenerateServiceTicketCarryoverReportSaturdayMsg
	End If
	
	If GenerateServiceTicketCarryoverReportSaturdayTime <> GenerateServiceTicketCarryoverReportSaturdayTime_ORIG Then
		If GenerateServiceTicketCarryoverReportSaturday_ORIG = 0 AND GenerateServiceTicketCarryoverReportSaturday = 1 Then
			CreateAuditLogEntry "Global Settings Change", "Global Settings Change", "Major", 1, "Service ticket carryover report schedule for Saturday turned on and set to run at " & GenerateServiceTicketCarryoverReportSaturdayTime
		Else
			CreateAuditLogEntry "Global Settings Change", "Global Settings Change", "Major", 1, "Service ticket carryover report scheduled run time for Saturday changed from " & GenerateServiceTicketCarryoverReportSaturdayTime_ORIG & " to " & GenerateServiceTicketCarryoverReportSaturdayTime
		End If
	End If


	If Request.Form("chkNoServiceTicketCarryoverReportIfClosed") = "on" then RunServiceTicketCarryoverReportIfClosedMsg = "On" Else RunServiceTicketCarryoverReportIfClosedMsg = "Off"
	If RunServiceTicketCarryoverReportIfClosed_ORIG = 1 then RunServiceTicketCarryoverReportIfClosedMsgOrig = "On" Else RunServiceTicketCarryoverReportIfClosedMsgOrig = "Off"
	
	If RunServiceTicketCarryoverReportIfClosed <> RunServiceTicketCarryoverReportIfClosed_ORIG Then
		CreateAuditLogEntry "Global Settings Change", "Global Settings Change", "Major", 1, "Service ticket carryover report schedule - Don''t Run If Closed (Monday-Friday Only) -  changed from " & RunServiceTicketCarryoverReportIfClosedMsgOrig & " to " & RunServiceTicketCarryoverReportIfClosedMsg
	End If


	If Request.Form("chkNoServiceTicketCarryoverReportIfClosingEarly") = "on" then RunServiceTicketCarryoverReportIfClosingEarlyMsg = "On" Else RunServiceTicketCarryoverReportIfClosingEarlyMsg = "Off"
	If RunServiceTicketCarryoverReportIfClosingEarly_ORIG = 1 then RunServiceTicketCarryoverReportIfClosingEarlyMsgOrig = "On" Else RunServiceTicketCarryoverReportIfClosingEarlyMsgOrig = "Off"
	
	If RunServiceTicketCarryoverReportIfClosingEarly <> RunServiceTicketCarryoverReportIfClosingEarly_ORIG Then
		CreateAuditLogEntry "Global Settings Change", "Global Settings Change", "Major", 1, "Service ticket carryover report schedule - Don''t Run If Closed (Monday-Friday Only) -  changed from " & RunServiceTicketCarryoverReportIfClosingEarlyMsgOrig & " to " & RunServiceTicketCarryoverReportIfClosingEarlyMsg
	End If


	'*********************************************************************
	'Build Array/String of Schedule Data From Request Form Field Data
	'*********************************************************************

	Schedule_ServiceTicketCarryoverReportGenerationUpdated = ""
	
	Schedule_ServiceTicketCarryoverReportGenerationUpdated = GenerateServiceTicketCarryoverReportSunday
	Schedule_ServiceTicketCarryoverReportGenerationUpdated = Schedule_ServiceTicketCarryoverReportGenerationUpdated & "," & GenerateServiceTicketCarryoverReportMonday
	Schedule_ServiceTicketCarryoverReportGenerationUpdated = Schedule_ServiceTicketCarryoverReportGenerationUpdated & "," & GenerateServiceTicketCarryoverReportTuesday
	Schedule_ServiceTicketCarryoverReportGenerationUpdated = Schedule_ServiceTicketCarryoverReportGenerationUpdated & "," & GenerateServiceTicketCarryoverReportWednesday
	Schedule_ServiceTicketCarryoverReportGenerationUpdated = Schedule_ServiceTicketCarryoverReportGenerationUpdated & "," & GenerateServiceTicketCarryoverReportThursday
	Schedule_ServiceTicketCarryoverReportGenerationUpdated = Schedule_ServiceTicketCarryoverReportGenerationUpdated & "," & GenerateServiceTicketCarryoverReportFriday
	Schedule_ServiceTicketCarryoverReportGenerationUpdated = Schedule_ServiceTicketCarryoverReportGenerationUpdated & "," & GenerateServiceTicketCarryoverReportSaturday
	
	Schedule_ServiceTicketCarryoverReportGenerationUpdated = Schedule_ServiceTicketCarryoverReportGenerationUpdated & "," & GenerateServiceTicketCarryoverReportSundayTime
	Schedule_ServiceTicketCarryoverReportGenerationUpdated = Schedule_ServiceTicketCarryoverReportGenerationUpdated & "," & GenerateServiceTicketCarryoverReportMondayTime
	Schedule_ServiceTicketCarryoverReportGenerationUpdated = Schedule_ServiceTicketCarryoverReportGenerationUpdated & "," & GenerateServiceTicketCarryoverReportTuesdayTime
	Schedule_ServiceTicketCarryoverReportGenerationUpdated = Schedule_ServiceTicketCarryoverReportGenerationUpdated & "," & GenerateServiceTicketCarryoverReportWednesdayTime
	Schedule_ServiceTicketCarryoverReportGenerationUpdated = Schedule_ServiceTicketCarryoverReportGenerationUpdated & "," & GenerateServiceTicketCarryoverReportThursdayTime
	Schedule_ServiceTicketCarryoverReportGenerationUpdated = Schedule_ServiceTicketCarryoverReportGenerationUpdated & "," & GenerateServiceTicketCarryoverReportFridayTime
	Schedule_ServiceTicketCarryoverReportGenerationUpdated = Schedule_ServiceTicketCarryoverReportGenerationUpdated & "," & GenerateServiceTicketCarryoverReportSaturdayTime

	
	Schedule_ServiceTicketCarryoverReportGenerationUpdated = Schedule_ServiceTicketCarryoverReportGenerationUpdated & "," & RunServiceTicketCarryoverReportIfClosed
	Schedule_ServiceTicketCarryoverReportGenerationUpdated = Schedule_ServiceTicketCarryoverReportGenerationUpdated & "," & RunServiceTicketCarryoverReportIfClosingEarly
	
	Response.Write("<br><br><br>Schedule_ServiceTicketCarryoverReportGenerationUpdated: " & Schedule_ServiceTicketCarryoverReportGenerationUpdated)

	'*********************************************************************
	'Update SQL with Array/String of Schedule Data
	'*********************************************************************
		
	SQL = "UPDATE Settings_FieldService SET Schedule_ServiceTicketCarryoverReportGeneration = '" & cStr(Schedule_ServiceTicketCarryoverReportGenerationUpdated) & "' "
	
	Response.Write("<br><br><br>SQL: " & SQL)
	
	Set cnn8 = Server.CreateObject("ADODB.Connection")
	cnn8.open (Session("ClientCnnString"))
	Set rs = Server.CreateObject("ADODB.Recordset")
	rs.CursorLocation = 3 
	Set rs = cnn8.Execute(SQL)
	
	set rs = Nothing
	cnn8.close
	set cnn8 = Nothing


	 Response.Redirect("field-service.asp")
	
%><!--#include file="../../../inc/footer-main.asp"-->