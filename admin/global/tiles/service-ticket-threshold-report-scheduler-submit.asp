<!--#include file="../../../inc/header.asp"-->

<%

	'***********************************************************
	'Get Values Of All Form Fields Posted 
	'***********************************************************
	
	GenerateServiceTicketthresholdReportSunday = Request.Form("chkNoServiceTicketthresholdReportSunday")
	GenerateServiceTicketthresholdReportMonday = Request.Form("chkNoServiceTicketthresholdReportMonday")
	GenerateServiceTicketthresholdReportTuesday = Request.Form("chkNoServiceTicketthresholdReportTuesday")
	GenerateServiceTicketthresholdReportWednesday = Request.Form("chkNoServiceTicketthresholdReportWednesday")
	GenerateServiceTicketthresholdReportThursday = Request.Form("chkNoServiceTicketthresholdReportThursday")
	GenerateServiceTicketthresholdReportFriday = Request.Form("chkNoServiceTicketthresholdReportFriday")
	GenerateServiceTicketthresholdReportSaturday = Request.Form("chkNoServiceTicketthresholdReportSaturday")
	
	GenerateServiceTicketthresholdReportSundayTime = Request.Form("txtServiceTicketthresholdReportSchedulerSundayTime")
	GenerateServiceTicketthresholdReportMondayTime = Request.Form("txtServiceTicketthresholdReportSchedulerMondayTime")
	GenerateServiceTicketthresholdReportTuesdayTime = Request.Form("txtServiceTicketthresholdReportSchedulerTuesdayTime")
	GenerateServiceTicketthresholdReportWednesdayTime = Request.Form("txtServiceTicketthresholdReportSchedulerWednesdayTime")
	GenerateServiceTicketthresholdReportThursdayTime = Request.Form("txtServiceTicketthresholdReportSchedulerThursdayTime")
	GenerateServiceTicketthresholdReportFridayTime = Request.Form("txtServiceTicketthresholdReportSchedulerFridayTime")
	GenerateServiceTicketthresholdReportSaturdayTime = Request.Form("txtServiceTicketthresholdReportSchedulerSaturdayTime")
	
	RunServiceTicketthresholdReportIfClosed = Request.Form("chkNoServiceTicketthresholdReportIfClosed")
	RunServiceTicketthresholdReportIfClosingEarly = Request.Form("chkNoServiceTicketthresholdReportIfClosingEarly")


	If Request.Form("chkNoServiceTicketthresholdReportSunday") = "on" Then
		GenerateServiceTicketthresholdReportSunday = 0
		GenerateServiceTicketthresholdReportSundayTime = ""
	Else 
		GenerateServiceTicketthresholdReportSunday = 1
	End If

	If Request.Form("chkNoServiceTicketthresholdReportMonday") = "on" Then
		GenerateServiceTicketthresholdReportMonday = 0
		GenerateServiceTicketthresholdReportMondayTime = ""
	Else 
		GenerateServiceTicketthresholdReportMonday = 1
	End If

	If Request.Form("chkNoServiceTicketthresholdReportTuesday") = "on" Then
		GenerateServiceTicketthresholdReportTuesday = 0
		GenerateServiceTicketthresholdReportTuesdayTime = ""
	Else 
		GenerateServiceTicketthresholdReportTuesday = 1
	End If

	If Request.Form("chkNoServiceTicketthresholdReportWednesday") = "on" Then
		GenerateServiceTicketthresholdReportWednesday = 0
		GenerateServiceTicketthresholdReportWednesdayTime = ""
	Else 
		GenerateServiceTicketthresholdReportWednesday = 1
	End If

	If Request.Form("chkNoServiceTicketthresholdReportThursday") = "on" Then
		GenerateServiceTicketthresholdReportThursday = 0
		GenerateServiceTicketthresholdReportThursdayTime = ""
	Else 
		GenerateServiceTicketthresholdReportThursday = 1
	End If

	If Request.Form("chkNoServiceTicketthresholdReportFriday") = "on" Then
		GenerateServiceTicketthresholdReportFriday = 0
		GenerateServiceTicketthresholdReportFridayTime = ""
	Else 
		GenerateServiceTicketthresholdReportFriday = 1
	End If

	If Request.Form("chkNoServiceTicketthresholdReportSaturday") = "on" Then
		GenerateServiceTicketthresholdReportSaturday = 0
		GenerateServiceTicketthresholdReportSaturdayTime = ""
	Else 
		GenerateServiceTicketthresholdReportSaturday = 1
	End If

	If Request.Form("chkNoServiceTicketthresholdReportIfClosed") = "on" Then RunServiceTicketthresholdReportIfClosed = 0 Else RunServiceTicketthresholdReportIfClosed = 1
	If Request.Form("chkNoServiceTicketthresholdReportIfClosingEarly") = "on" Then RunServiceTicketthresholdReportIfClosingEarly = 0 Else RunServiceTicketthresholdReportIfClosingEarly = 1
	
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
	
		Schedule_ServiceTicketthresholdReportGeneration = rsFieldServiceSettings("Schedule_ServiceTicketthresholdReportGeneration")
		
		Schedule_ServiceTicketthresholdReportGenerationSettings = Split(Schedule_ServiceTicketthresholdReportGeneration,",")

		GenerateServiceTicketthresholdReportSunday_ORIG = cInt(Schedule_ServiceTicketthresholdReportGenerationSettings(0))
		GenerateServiceTicketthresholdReportMonday_ORIG = cInt(Schedule_ServiceTicketthresholdReportGenerationSettings(1))
		GenerateServiceTicketthresholdReportTuesday_ORIG = cInt(Schedule_ServiceTicketthresholdReportGenerationSettings(2))
		GenerateServiceTicketthresholdReportWednesday_ORIG = cInt(Schedule_ServiceTicketthresholdReportGenerationSettings(3))
		GenerateServiceTicketthresholdReportThursday_ORIG = cInt(Schedule_ServiceTicketthresholdReportGenerationSettings(4))
		GenerateServiceTicketthresholdReportFriday_ORIG = cInt(Schedule_ServiceTicketthresholdReportGenerationSettings(5))
		GenerateServiceTicketthresholdReportSaturday_ORIG = cInt(Schedule_ServiceTicketthresholdReportGenerationSettings(6))
		GenerateServiceTicketthresholdReportSundayTime_ORIG = Schedule_ServiceTicketthresholdReportGenerationSettings(7)
		GenerateServiceTicketthresholdReportMondayTime_ORIG = Schedule_ServiceTicketthresholdReportGenerationSettings(8)
		GenerateServiceTicketthresholdReportTuesdayTime_ORIG = Schedule_ServiceTicketthresholdReportGenerationSettings(9)
		GenerateServiceTicketthresholdReportWednesdayTime_ORIG = Schedule_ServiceTicketthresholdReportGenerationSettings(10)
		GenerateServiceTicketthresholdReportThursdayTime_ORIG = Schedule_ServiceTicketthresholdReportGenerationSettings(11)
		GenerateServiceTicketthresholdReportFridayTime_ORIG = Schedule_ServiceTicketthresholdReportGenerationSettings(12)
		GenerateServiceTicketthresholdReportSaturdayTime_ORIG = Schedule_ServiceTicketthresholdReportGenerationSettings(13)
		RunServiceTicketthresholdReportIfClosed_ORIG = cInt(Schedule_ServiceTicketthresholdReportGenerationSettings(14))
		RunServiceTicketthresholdReportIfClosingEarly_ORIG = cInt(Schedule_ServiceTicketthresholdReportGenerationSettings(15))
	
	End If
	
	set rsFieldServiceSettings = Nothing
	cnnFieldServiceSettings.close
	set cnnFieldServiceSettings = Nothing
	

	'***********************************************************
	'Perform Audit Trail Entry Inserts
	'***********************************************************

	
	If Request.Form("chkNoServiceTicketthresholdReportSunday") = "on" then GenerateServiceTicketthresholdReportSundayMsg = "On" Else GenerateServiceTicketthresholdReportSundayMsg = "Off"
	If GenerateServiceTicketthresholdReportSunday_ORIG = 1 then GenerateServiceTicketthresholdReportSundayMsgOrig = "On" Else GenerateServiceTicketthresholdReportSundayMsgOrig = "Off"
	
	If GenerateServiceTicketthresholdReportSunday <> GenerateServiceTicketthresholdReportSunday_ORIG Then
		CreateAuditLogEntry "Global Settings Change", "Global Settings Change", "Major", 1, "Service ticket threshold report schedule for SUNDAY changed from " & GenerateServiceTicketthresholdReportSundayMsgOrig & " to " & GenerateServiceTicketthresholdReportSundayMsg
	End If
	
	If GenerateServiceTicketthresholdReportSundayTime <> GenerateServiceTicketthresholdReportSundayTime_ORIG Then
		If GenerateServiceTicketthresholdReportSunday_ORIG = 0 AND GenerateServiceTicketthresholdReportSunday = 1 Then
			CreateAuditLogEntry "Global Settings Change", "Global Settings Change", "Major", 1, "Service ticket threshold report schedule for SUNDAY turned on and set to run at " & GenerateServiceTicketthresholdReportSundayTime
		Else
			CreateAuditLogEntry "Global Settings Change", "Global Settings Change", "Major", 1, "Service ticket threshold report scheduled run time for SUNDAY changed from " & GenerateServiceTicketthresholdReportSundayTime_ORIG & " to " & GenerateServiceTicketthresholdReportSundayTime
		End If
	End If
	
	

	If Request.Form("chkNoServiceTicketthresholdReportMonday") = "on" then GenerateServiceTicketthresholdReportMondayMsg = "On" Else GenerateServiceTicketthresholdReportMondayMsg = "Off"
	If GenerateServiceTicketthresholdReportMonday_ORIG = 1 then GenerateServiceTicketthresholdReportMondayMsgOrig = "On" Else GenerateServiceTicketthresholdReportMondayMsgOrig = "Off"
	
	If GenerateServiceTicketthresholdReportMonday <> GenerateServiceTicketthresholdReportMonday_ORIG Then
		CreateAuditLogEntry "Global Settings Change", "Global Settings Change", "Major", 1, "Service ticket threshold report schedule for Monday changed from " & GenerateServiceTicketthresholdReportMondayMsgOrig & " to " & GenerateServiceTicketthresholdReportMondayMsg
	End If
	
	If GenerateServiceTicketthresholdReportMondayTime <> GenerateServiceTicketthresholdReportMondayTime_ORIG Then
		If GenerateServiceTicketthresholdReportMonday_ORIG = 0 AND GenerateServiceTicketthresholdReportMonday = 1 Then
			CreateAuditLogEntry "Global Settings Change", "Global Settings Change", "Major", 1, "Service ticket threshold report schedule for Monday turned on and set to run at " & GenerateServiceTicketthresholdReportMondayTime
		Else
			CreateAuditLogEntry "Global Settings Change", "Global Settings Change", "Major", 1, "Service ticket threshold report scheduled run time for Monday changed from " & GenerateServiceTicketthresholdReportMondayTime_ORIG & " to " & GenerateServiceTicketthresholdReportMondayTime
		End If
	End If
	


	If Request.Form("chkNoServiceTicketthresholdReportTuesday") = "on" then GenerateServiceTicketthresholdReportTuesdayMsg = "On" Else GenerateServiceTicketthresholdReportTuesdayMsg = "Off"
	If GenerateServiceTicketthresholdReportTuesday_ORIG = 1 then GenerateServiceTicketthresholdReportTuesdayMsgOrig = "On" Else GenerateServiceTicketthresholdReportTuesdayMsgOrig = "Off"
	
	If GenerateServiceTicketthresholdReportTuesday <> GenerateServiceTicketthresholdReportTuesday_ORIG Then
		CreateAuditLogEntry "Global Settings Change", "Global Settings Change", "Major", 1, "Service ticket threshold report schedule for Tuesday changed from " & GenerateServiceTicketthresholdReportTuesdayMsgOrig & " to " & GenerateServiceTicketthresholdReportTuesdayMsg
	End If
	
	If GenerateServiceTicketthresholdReportTuesdayTime <> GenerateServiceTicketthresholdReportTuesdayTime_ORIG Then
		If GenerateServiceTicketthresholdReportTuesday_ORIG = 0 AND GenerateServiceTicketthresholdReportTuesday = 1 Then
			CreateAuditLogEntry "Global Settings Change", "Global Settings Change", "Major", 1, "Service ticket threshold report schedule for Tuesday turned on and set to run at " & GenerateServiceTicketthresholdReportTuesdayTime
		Else
			CreateAuditLogEntry "Global Settings Change", "Global Settings Change", "Major", 1, "Service ticket threshold report scheduled run time for Tuesday changed from " & GenerateServiceTicketthresholdReportTuesdayTime_ORIG & " to " & GenerateServiceTicketthresholdReportTuesdayTime
		End If
	End If



	If Request.Form("chkNoServiceTicketthresholdReportWednesday") = "on" then GenerateServiceTicketthresholdReportWednesdayMsg = "On" Else GenerateServiceTicketthresholdReportWednesdayMsg = "Off"
	If GenerateServiceTicketthresholdReportWednesday_ORIG = 1 then GenerateServiceTicketthresholdReportWednesdayMsgOrig = "On" Else GenerateServiceTicketthresholdReportWednesdayMsgOrig = "Off"
	
	If GenerateServiceTicketthresholdReportWednesday <> GenerateServiceTicketthresholdReportWednesday_ORIG Then
		CreateAuditLogEntry "Global Settings Change", "Global Settings Change", "Major", 1, "Service ticket threshold report schedule for Wednesday changed from " & GenerateServiceTicketthresholdReportWednesdayMsgOrig & " to " & GenerateServiceTicketthresholdReportWednesdayMsg
	End If
	
	If GenerateServiceTicketthresholdReportWednesdayTime <> GenerateServiceTicketthresholdReportWednesdayTime_ORIG Then
		If GenerateServiceTicketthresholdReportWednesday_ORIG = 0 AND GenerateServiceTicketthresholdReportWednesday = 1 Then
			CreateAuditLogEntry "Global Settings Change", "Global Settings Change", "Major", 1, "Service ticket threshold report schedule for Wednesday turned on and set to run at " & GenerateServiceTicketthresholdReportWednesdayTime
		Else
			CreateAuditLogEntry "Global Settings Change", "Global Settings Change", "Major", 1, "Service ticket threshold report scheduled run time for Wednesday changed from " & GenerateServiceTicketthresholdReportWednesdayTime_ORIG & " to " & GenerateServiceTicketthresholdReportWednesdayTime
		End If
	End If



	If Request.Form("chkNoServiceTicketthresholdReportThursday") = "on" then GenerateServiceTicketthresholdReportThursdayMsg = "On" Else GenerateServiceTicketthresholdReportThursdayMsg = "Off"
	If GenerateServiceTicketthresholdReportThursday_ORIG = 1 then GenerateServiceTicketthresholdReportThursdayMsgOrig = "On" Else GenerateServiceTicketthresholdReportThursdayMsgOrig = "Off"
	
	If GenerateServiceTicketthresholdReportThursday <> GenerateServiceTicketthresholdReportThursday_ORIG Then
		CreateAuditLogEntry "Global Settings Change", "Global Settings Change", "Major", 1, "Service ticket threshold report schedule for Thursday changed from " & GenerateServiceTicketthresholdReportThursdayMsgOrig & " to " & GenerateServiceTicketthresholdReportThursdayMsg
	End If
	
	If GenerateServiceTicketthresholdReportThursdayTime <> GenerateServiceTicketthresholdReportThursdayTime_ORIG Then
		If GenerateServiceTicketthresholdReportThursday_ORIG = 0 AND GenerateServiceTicketthresholdReportThursday = 1 Then
			CreateAuditLogEntry "Global Settings Change", "Global Settings Change", "Major", 1, "Service ticket threshold report schedule for Thursday turned on and set to run at " & GenerateServiceTicketthresholdReportThursdayTime
		Else
			CreateAuditLogEntry "Global Settings Change", "Global Settings Change", "Major", 1, "Service ticket threshold report scheduled run time for Thursday changed from " & GenerateServiceTicketthresholdReportThursdayTime_ORIG & " to " & GenerateServiceTicketthresholdReportThursdayTime
		End If
	End If



	If Request.Form("chkNoServiceTicketthresholdReportFriday") = "on" then GenerateServiceTicketthresholdReportFridayMsg = "On" Else GenerateServiceTicketthresholdReportFridayMsg = "Off"
	If GenerateServiceTicketthresholdReportFriday_ORIG = 1 then GenerateServiceTicketthresholdReportFridayMsgOrig = "On" Else GenerateServiceTicketthresholdReportFridayMsgOrig = "Off"
	
	If GenerateServiceTicketthresholdReportFriday <> GenerateServiceTicketthresholdReportFriday_ORIG Then
		CreateAuditLogEntry "Global Settings Change", "Global Settings Change", "Major", 1, "Service ticket threshold report schedule for Friday changed from " & GenerateServiceTicketthresholdReportFridayMsgOrig & " to " & GenerateServiceTicketthresholdReportFridayMsg
	End If
	
	If GenerateServiceTicketthresholdReportFridayTime <> GenerateServiceTicketthresholdReportFridayTime_ORIG Then
		If GenerateServiceTicketthresholdReportFriday_ORIG = 0 AND GenerateServiceTicketthresholdReportFriday = 1 Then
			CreateAuditLogEntry "Global Settings Change", "Global Settings Change", "Major", 1, "Service ticket threshold report schedule for Friday turned on and set to run at " & GenerateServiceTicketthresholdReportFridayTime
		Else
			CreateAuditLogEntry "Global Settings Change", "Global Settings Change", "Major", 1, "Service ticket threshold report scheduled run time for Friday changed from " & GenerateServiceTicketthresholdReportFridayTime_ORIG & " to " & GenerateServiceTicketthresholdReportFridayTime
		End If
	End If



	If Request.Form("chkNoServiceTicketthresholdReportSaturday") = "on" then GenerateServiceTicketthresholdReportSaturdayMsg = "On" Else GenerateServiceTicketthresholdReportSaturdayMsg = "Off"
	If GenerateServiceTicketthresholdReportSaturday_ORIG = 1 then GenerateServiceTicketthresholdReportSaturdayMsgOrig = "On" Else GenerateServiceTicketthresholdReportSaturdayMsgOrig = "Off"
	
	If GenerateServiceTicketthresholdReportSaturday <> GenerateServiceTicketthresholdReportSaturday_ORIG Then
		CreateAuditLogEntry "Global Settings Change", "Global Settings Change", "Major", 1, "Service ticket threshold report schedule for Saturday changed from " & GenerateServiceTicketthresholdReportSaturdayMsgOrig & " to " & GenerateServiceTicketthresholdReportSaturdayMsg
	End If
	
	If GenerateServiceTicketthresholdReportSaturdayTime <> GenerateServiceTicketthresholdReportSaturdayTime_ORIG Then
		If GenerateServiceTicketthresholdReportSaturday_ORIG = 0 AND GenerateServiceTicketthresholdReportSaturday = 1 Then
			CreateAuditLogEntry "Global Settings Change", "Global Settings Change", "Major", 1, "Service ticket threshold report schedule for Saturday turned on and set to run at " & GenerateServiceTicketthresholdReportSaturdayTime
		Else
			CreateAuditLogEntry "Global Settings Change", "Global Settings Change", "Major", 1, "Service ticket threshold report scheduled run time for Saturday changed from " & GenerateServiceTicketthresholdReportSaturdayTime_ORIG & " to " & GenerateServiceTicketthresholdReportSaturdayTime
		End If
	End If


	If Request.Form("chkNoServiceTicketthresholdReportIfClosed") = "on" then RunServiceTicketthresholdReportIfClosedMsg = "On" Else RunServiceTicketthresholdReportIfClosedMsg = "Off"
	If RunServiceTicketthresholdReportIfClosed_ORIG = 1 then RunServiceTicketthresholdReportIfClosedMsgOrig = "On" Else RunServiceTicketthresholdReportIfClosedMsgOrig = "Off"
	
	If RunServiceTicketthresholdReportIfClosed <> RunServiceTicketthresholdReportIfClosed_ORIG Then
		CreateAuditLogEntry "Global Settings Change", "Global Settings Change", "Major", 1, "Service ticket threshold report schedule - Don''t Run If Closed (Monday-Friday Only) -  changed from " & RunServiceTicketthresholdReportIfClosedMsgOrig & " to " & RunServiceTicketthresholdReportIfClosedMsg
	End If


	If Request.Form("chkNoServiceTicketthresholdReportIfClosingEarly") = "on" then RunServiceTicketthresholdReportIfClosingEarlyMsg = "On" Else RunServiceTicketthresholdReportIfClosingEarlyMsg = "Off"
	If RunServiceTicketthresholdReportIfClosingEarly_ORIG = 1 then RunServiceTicketthresholdReportIfClosingEarlyMsgOrig = "On" Else RunServiceTicketthresholdReportIfClosingEarlyMsgOrig = "Off"
	
	If RunServiceTicketthresholdReportIfClosingEarly <> RunServiceTicketthresholdReportIfClosingEarly_ORIG Then
		CreateAuditLogEntry "Global Settings Change", "Global Settings Change", "Major", 1, "Service ticket threshold report schedule - Don''t Run If Closed (Monday-Friday Only) -  changed from " & RunServiceTicketthresholdReportIfClosingEarlyMsgOrig & " to " & RunServiceTicketthresholdReportIfClosingEarlyMsg
	End If


	'*********************************************************************
	'Build Array/String of Schedule Data From Request Form Field Data
	'*********************************************************************

	Schedule_ServiceTicketthresholdReportGenerationUpdated = ""
	
	Schedule_ServiceTicketthresholdReportGenerationUpdated = GenerateServiceTicketthresholdReportSunday
	Schedule_ServiceTicketthresholdReportGenerationUpdated = Schedule_ServiceTicketthresholdReportGenerationUpdated & "," & GenerateServiceTicketthresholdReportMonday
	Schedule_ServiceTicketthresholdReportGenerationUpdated = Schedule_ServiceTicketthresholdReportGenerationUpdated & "," & GenerateServiceTicketthresholdReportTuesday
	Schedule_ServiceTicketthresholdReportGenerationUpdated = Schedule_ServiceTicketthresholdReportGenerationUpdated & "," & GenerateServiceTicketthresholdReportWednesday
	Schedule_ServiceTicketthresholdReportGenerationUpdated = Schedule_ServiceTicketthresholdReportGenerationUpdated & "," & GenerateServiceTicketthresholdReportThursday
	Schedule_ServiceTicketthresholdReportGenerationUpdated = Schedule_ServiceTicketthresholdReportGenerationUpdated & "," & GenerateServiceTicketthresholdReportFriday
	Schedule_ServiceTicketthresholdReportGenerationUpdated = Schedule_ServiceTicketthresholdReportGenerationUpdated & "," & GenerateServiceTicketthresholdReportSaturday
	
	Schedule_ServiceTicketthresholdReportGenerationUpdated = Schedule_ServiceTicketthresholdReportGenerationUpdated & "," & GenerateServiceTicketthresholdReportSundayTime
	Schedule_ServiceTicketthresholdReportGenerationUpdated = Schedule_ServiceTicketthresholdReportGenerationUpdated & "," & GenerateServiceTicketthresholdReportMondayTime
	Schedule_ServiceTicketthresholdReportGenerationUpdated = Schedule_ServiceTicketthresholdReportGenerationUpdated & "," & GenerateServiceTicketthresholdReportTuesdayTime
	Schedule_ServiceTicketthresholdReportGenerationUpdated = Schedule_ServiceTicketthresholdReportGenerationUpdated & "," & GenerateServiceTicketthresholdReportWednesdayTime
	Schedule_ServiceTicketthresholdReportGenerationUpdated = Schedule_ServiceTicketthresholdReportGenerationUpdated & "," & GenerateServiceTicketthresholdReportThursdayTime
	Schedule_ServiceTicketthresholdReportGenerationUpdated = Schedule_ServiceTicketthresholdReportGenerationUpdated & "," & GenerateServiceTicketthresholdReportFridayTime
	Schedule_ServiceTicketthresholdReportGenerationUpdated = Schedule_ServiceTicketthresholdReportGenerationUpdated & "," & GenerateServiceTicketthresholdReportSaturdayTime

	
	Schedule_ServiceTicketthresholdReportGenerationUpdated = Schedule_ServiceTicketthresholdReportGenerationUpdated & "," & RunServiceTicketthresholdReportIfClosed
	Schedule_ServiceTicketthresholdReportGenerationUpdated = Schedule_ServiceTicketthresholdReportGenerationUpdated & "," & RunServiceTicketthresholdReportIfClosingEarly
	
	Response.Write("<br><br><br>Schedule_ServiceTicketthresholdReportGenerationUpdated: " & Schedule_ServiceTicketthresholdReportGenerationUpdated)

	'*********************************************************************
	'Update SQL with Array/String of Schedule Data
	'*********************************************************************
		
	SQL = "UPDATE Settings_FieldService SET Schedule_ServiceTicketthresholdReportGeneration = '" & cStr(Schedule_ServiceTicketthresholdReportGenerationUpdated) & "' "
	
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