<!--#include file="../../../inc/header.asp"-->

<%

	'***********************************************************
	'Get Values Of All Form Fields Posted
	'***********************************************************
	
	GenerateFieldServiceNotesReportSunday = Request.Form("chkNoFieldServiceNotesReportSunday")
	GenerateFieldServiceNotesReportMonday = Request.Form("chkNoFieldServiceNotesReportMonday")
	GenerateFieldServiceNotesReportTuesday = Request.Form("chkNoFieldServiceNotesReportTuesday")
	GenerateFieldServiceNotesReportWednesday = Request.Form("chkNoFieldServiceNotesReportWednesday")
	GenerateFieldServiceNotesReportThursday = Request.Form("chkNoFieldServiceNotesReportThursday")
	GenerateFieldServiceNotesReportFriday = Request.Form("chkNoFieldServiceNotesReportFriday")
	GenerateFieldServiceNotesReportSaturday = Request.Form("chkNoFieldServiceNotesReportSaturday")
	
	GenerateFieldServiceNotesReportSundayTime = Request.Form("txtFieldServiceNotesReportSchedulerSundayTime")
	GenerateFieldServiceNotesReportMondayTime = Request.Form("txtFieldServiceNotesReportSchedulerMondayTime")
	GenerateFieldServiceNotesReportTuesdayTime = Request.Form("txtFieldServiceNotesReportSchedulerTuesdayTime")
	GenerateFieldServiceNotesReportWednesdayTime = Request.Form("txtFieldServiceNotesReportSchedulerWednesdayTime")
	GenerateFieldServiceNotesReportThursdayTime = Request.Form("txtFieldServiceNotesReportSchedulerThursdayTime")
	GenerateFieldServiceNotesReportFridayTime = Request.Form("txtFieldServiceNotesReportSchedulerFridayTime")
	GenerateFieldServiceNotesReportSaturdayTime = Request.Form("txtFieldServiceNotesReportSchedulerSaturdayTime")
	
	RunFieldServiceNotesReportIfClosed = Request.Form("chkNoFieldServiceNotesReportIfClosed")
	RunFieldServiceNotesReportIfClosingEarly = Request.Form("chkNoFieldServiceNotesReportIfClosingEarly")


	If Request.Form("chkNoFieldServiceNotesReportSunday") = "on" Then
		GenerateFieldServiceNotesReportSunday = 0
		GenerateFieldServiceNotesReportSundayTime = ""
	Else 
		GenerateFieldServiceNotesReportSunday = 1
	End If

	If Request.Form("chkNoFieldServiceNotesReportMonday") = "on" Then
		GenerateFieldServiceNotesReportMonday = 0
		GenerateFieldServiceNotesReportMondayTime = ""
	Else 
		GenerateFieldServiceNotesReportMonday = 1
	End If

	If Request.Form("chkNoFieldServiceNotesReportTuesday") = "on" Then
		GenerateFieldServiceNotesReportTuesday = 0
		GenerateFieldServiceNotesReportTuesdayTime = ""
	Else 
		GenerateFieldServiceNotesReportTuesday = 1
	End If

	If Request.Form("chkNoFieldServiceNotesReportWednesday") = "on" Then
		GenerateFieldServiceNotesReportWednesday = 0
		GenerateFieldServiceNotesReportWednesdayTime = ""
	Else 
		GenerateFieldServiceNotesReportWednesday = 1
	End If

	If Request.Form("chkNoFieldServiceNotesReportThursday") = "on" Then
		GenerateFieldServiceNotesReportThursday = 0
		GenerateFieldServiceNotesReportThursdayTime = ""
	Else 
		GenerateFieldServiceNotesReportThursday = 1
	End If

	If Request.Form("chkNoFieldServiceNotesReportFriday") = "on" Then
		GenerateFieldServiceNotesReportFriday = 0
		GenerateFieldServiceNotesReportFridayTime = ""
	Else 
		GenerateFieldServiceNotesReportFriday = 1
	End If

	If Request.Form("chkNoFieldServiceNotesReportSaturday") = "on" Then
		GenerateFieldServiceNotesReportSaturday = 0
		GenerateFieldServiceNotesReportSaturdayTime = ""
	Else 
		GenerateFieldServiceNotesReportSaturday = 1
	End If

	If Request.Form("chkNoFieldServiceNotesReportIfClosed") = "on" Then RunFieldServiceNotesReportIfClosed = 0 Else RunFieldServiceNotesReportIfClosed = 1
	If Request.Form("chkNoFieldServiceNotesReportIfClosingEarly") = "on" Then RunFieldServiceNotesReportIfClosingEarly = 0 Else RunFieldServiceNotesReportIfClosingEarly = 1
	
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
	
		Schedule_FieldServiceNotesReportGeneration = rsFieldServiceSettings("Schedule_FieldServiceNotesReportGeneration")
		
		Schedule_FieldServiceNotesReportGenerationSettings = Split(Schedule_FieldServiceNotesReportGeneration,",")

		GenerateFieldServiceNotesReportSunday_ORIG = cInt(Schedule_FieldServiceNotesReportGenerationSettings(0))
		GenerateFieldServiceNotesReportMonday_ORIG = cInt(Schedule_FieldServiceNotesReportGenerationSettings(1))
		GenerateFieldServiceNotesReportTuesday_ORIG = cInt(Schedule_FieldServiceNotesReportGenerationSettings(2))
		GenerateFieldServiceNotesReportWednesday_ORIG = cInt(Schedule_FieldServiceNotesReportGenerationSettings(3))
		GenerateFieldServiceNotesReportThursday_ORIG = cInt(Schedule_FieldServiceNotesReportGenerationSettings(4))
		GenerateFieldServiceNotesReportFriday_ORIG = cInt(Schedule_FieldServiceNotesReportGenerationSettings(5))
		GenerateFieldServiceNotesReportSaturday_ORIG = cInt(Schedule_FieldServiceNotesReportGenerationSettings(6))
		GenerateFieldServiceNotesReportSundayTime_ORIG = Schedule_FieldServiceNotesReportGenerationSettings(7)
		GenerateFieldServiceNotesReportMondayTime_ORIG = Schedule_FieldServiceNotesReportGenerationSettings(8)
		GenerateFieldServiceNotesReportTuesdayTime_ORIG = Schedule_FieldServiceNotesReportGenerationSettings(9)
		GenerateFieldServiceNotesReportWednesdayTime_ORIG = Schedule_FieldServiceNotesReportGenerationSettings(10)
		GenerateFieldServiceNotesReportThursdayTime_ORIG = Schedule_FieldServiceNotesReportGenerationSettings(11)
		GenerateFieldServiceNotesReportFridayTime_ORIG = Schedule_FieldServiceNotesReportGenerationSettings(12)
		GenerateFieldServiceNotesReportSaturdayTime_ORIG = Schedule_FieldServiceNotesReportGenerationSettings(13)
		RunFieldServiceNotesReportIfClosed_ORIG = cInt(Schedule_FieldServiceNotesReportGenerationSettings(14))
		RunFieldServiceNotesReportIfClosingEarly_ORIG = cInt(Schedule_FieldServiceNotesReportGenerationSettings(15))
	
	End If
	
	set rsFieldServiceSettings = Nothing
	cnnFieldServiceSettings.close
	set cnnFieldServiceSettings = Nothing
	

	'***********************************************************
	'Perform Audit Trail Entry Inserts
	'***********************************************************

	
	If Request.Form("chkNoFieldServiceNotesReportSunday") = "on" then GenerateFieldServiceNotesReportSundayMsg = "On" Else GenerateFieldServiceNotesReportSundayMsg = "Off"
	If GenerateFieldServiceNotesReportSunday_ORIG = 1 then GenerateFieldServiceNotesReportSundayMsgOrig = "On" Else GenerateFieldServiceNotesReportSundayMsgOrig = "Off"
	
	If GenerateFieldServiceNotesReportSunday <> GenerateFieldServiceNotesReportSunday_ORIG Then
		CreateAuditLogEntry "Global Settings Change", "Global Settings Change", "Major", 1, "Field service notes report generation schedule for SUNDAY changed from " & GenerateFieldServiceNotesReportSundayMsgOrig & " to " & GenerateFieldServiceNotesReportSundayMsg
	End If
	
	If GenerateFieldServiceNotesReportSundayTime <> GenerateFieldServiceNotesReportSundayTime_ORIG Then
		If GenerateFieldServiceNotesReportSunday_ORIG = 0 AND GenerateFieldServiceNotesReportSunday = 1 Then
			CreateAuditLogEntry "Global Settings Change", "Global Settings Change", "Major", 1, "Field service notes report generation schedule for SUNDAY turned on and set to run at " & GenerateFieldServiceNotesReportSundayTime
		Else
			CreateAuditLogEntry "Global Settings Change", "Global Settings Change", "Major", 1, "Field service notes report generation scheduled run time for SUNDAY changed from " & GenerateFieldServiceNotesReportSundayTime_ORIG & " to " & GenerateFieldServiceNotesReportSundayTime
		End If
	End If
	
	

	If Request.Form("chkNoFieldServiceNotesReportMonday") = "on" then GenerateFieldServiceNotesReportMondayMsg = "On" Else GenerateFieldServiceNotesReportMondayMsg = "Off"
	If GenerateFieldServiceNotesReportMonday_ORIG = 1 then GenerateFieldServiceNotesReportMondayMsgOrig = "On" Else GenerateFieldServiceNotesReportMondayMsgOrig = "Off"
	
	If GenerateFieldServiceNotesReportMonday <> GenerateFieldServiceNotesReportMonday_ORIG Then
		CreateAuditLogEntry "Global Settings Change", "Global Settings Change", "Major", 1, "Field service notes report generation schedule for Monday changed from " & GenerateFieldServiceNotesReportMondayMsgOrig & " to " & GenerateFieldServiceNotesReportMondayMsg
	End If
	
	If GenerateFieldServiceNotesReportMondayTime <> GenerateFieldServiceNotesReportMondayTime_ORIG Then
		If GenerateFieldServiceNotesReportMonday_ORIG = 0 AND GenerateFieldServiceNotesReportMonday = 1 Then
			CreateAuditLogEntry "Global Settings Change", "Global Settings Change", "Major", 1, "Field service notes report generation schedule for Monday turned on and set to run at " & GenerateFieldServiceNotesReportMondayTime
		Else
			CreateAuditLogEntry "Global Settings Change", "Global Settings Change", "Major", 1, "Field service notes report generation scheduled run time for Monday changed from " & GenerateFieldServiceNotesReportMondayTime_ORIG & " to " & GenerateFieldServiceNotesReportMondayTime
		End If
	End If
	


	If Request.Form("chkNoFieldServiceNotesReportTuesday") = "on" then GenerateFieldServiceNotesReportTuesdayMsg = "On" Else GenerateFieldServiceNotesReportTuesdayMsg = "Off"
	If GenerateFieldServiceNotesReportTuesday_ORIG = 1 then GenerateFieldServiceNotesReportTuesdayMsgOrig = "On" Else GenerateFieldServiceNotesReportTuesdayMsgOrig = "Off"
	
	If GenerateFieldServiceNotesReportTuesday <> GenerateFieldServiceNotesReportTuesday_ORIG Then
		CreateAuditLogEntry "Global Settings Change", "Global Settings Change", "Major", 1, "Field service notes report generation schedule for Tuesday changed from " & GenerateFieldServiceNotesReportTuesdayMsgOrig & " to " & GenerateFieldServiceNotesReportTuesdayMsg
	End If
	
	If GenerateFieldServiceNotesReportTuesdayTime <> GenerateFieldServiceNotesReportTuesdayTime_ORIG Then
		If GenerateFieldServiceNotesReportTuesday_ORIG = 0 AND GenerateFieldServiceNotesReportTuesday = 1 Then
			CreateAuditLogEntry "Global Settings Change", "Global Settings Change", "Major", 1, "Field service notes report generation schedule for Tuesday turned on and set to run at " & GenerateFieldServiceNotesReportTuesdayTime
		Else
			CreateAuditLogEntry "Global Settings Change", "Global Settings Change", "Major", 1, "Field service notes report generation scheduled run time for Tuesday changed from " & GenerateFieldServiceNotesReportTuesdayTime_ORIG & " to " & GenerateFieldServiceNotesReportTuesdayTime
		End If
	End If



	If Request.Form("chkNoFieldServiceNotesReportWednesday") = "on" then GenerateFieldServiceNotesReportWednesdayMsg = "On" Else GenerateFieldServiceNotesReportWednesdayMsg = "Off"
	If GenerateFieldServiceNotesReportWednesday_ORIG = 1 then GenerateFieldServiceNotesReportWednesdayMsgOrig = "On" Else GenerateFieldServiceNotesReportWednesdayMsgOrig = "Off"
	
	If GenerateFieldServiceNotesReportWednesday <> GenerateFieldServiceNotesReportWednesday_ORIG Then
		CreateAuditLogEntry "Global Settings Change", "Global Settings Change", "Major", 1, "Field service notes report generation schedule for Wednesday changed from " & GenerateFieldServiceNotesReportWednesdayMsgOrig & " to " & GenerateFieldServiceNotesReportWednesdayMsg
	End If
	
	If GenerateFieldServiceNotesReportWednesdayTime <> GenerateFieldServiceNotesReportWednesdayTime_ORIG Then
		If GenerateFieldServiceNotesReportWednesday_ORIG = 0 AND GenerateFieldServiceNotesReportWednesday = 1 Then
			CreateAuditLogEntry "Global Settings Change", "Global Settings Change", "Major", 1, "Field service notes report generation schedule for Wednesday turned on and set to run at " & GenerateFieldServiceNotesReportWednesdayTime
		Else
			CreateAuditLogEntry "Global Settings Change", "Global Settings Change", "Major", 1, "Field service notes report generation scheduled run time for Wednesday changed from " & GenerateFieldServiceNotesReportWednesdayTime_ORIG & " to " & GenerateFieldServiceNotesReportWednesdayTime
		End If
	End If



	If Request.Form("chkNoFieldServiceNotesReportThursday") = "on" then GenerateFieldServiceNotesReportThursdayMsg = "On" Else GenerateFieldServiceNotesReportThursdayMsg = "Off"
	If GenerateFieldServiceNotesReportThursday_ORIG = 1 then GenerateFieldServiceNotesReportThursdayMsgOrig = "On" Else GenerateFieldServiceNotesReportThursdayMsgOrig = "Off"
	
	If GenerateFieldServiceNotesReportThursday <> GenerateFieldServiceNotesReportThursday_ORIG Then
		CreateAuditLogEntry "Global Settings Change", "Global Settings Change", "Major", 1, "Field service notes report generation schedule for Thursday changed from " & GenerateFieldServiceNotesReportThursdayMsgOrig & " to " & GenerateFieldServiceNotesReportThursdayMsg
	End If
	
	If GenerateFieldServiceNotesReportThursdayTime <> GenerateFieldServiceNotesReportThursdayTime_ORIG Then
		If GenerateFieldServiceNotesReportThursday_ORIG = 0 AND GenerateFieldServiceNotesReportThursday = 1 Then
			CreateAuditLogEntry "Global Settings Change", "Global Settings Change", "Major", 1, "Field service notes report generation schedule for Thursday turned on and set to run at " & GenerateFieldServiceNotesReportThursdayTime
		Else
			CreateAuditLogEntry "Global Settings Change", "Global Settings Change", "Major", 1, "Field service notes report generation scheduled run time for Thursday changed from " & GenerateFieldServiceNotesReportThursdayTime_ORIG & " to " & GenerateFieldServiceNotesReportThursdayTime
		End If
	End If



	If Request.Form("chkNoFieldServiceNotesReportFriday") = "on" then GenerateFieldServiceNotesReportFridayMsg = "On" Else GenerateFieldServiceNotesReportFridayMsg = "Off"
	If GenerateFieldServiceNotesReportFriday_ORIG = 1 then GenerateFieldServiceNotesReportFridayMsgOrig = "On" Else GenerateFieldServiceNotesReportFridayMsgOrig = "Off"
	
	If GenerateFieldServiceNotesReportFriday <> GenerateFieldServiceNotesReportFriday_ORIG Then
		CreateAuditLogEntry "Global Settings Change", "Global Settings Change", "Major", 1, "Field service notes report generation schedule for Friday changed from " & GenerateFieldServiceNotesReportFridayMsgOrig & " to " & GenerateFieldServiceNotesReportFridayMsg
	End If
	
	If GenerateFieldServiceNotesReportFridayTime <> GenerateFieldServiceNotesReportFridayTime_ORIG Then
		If GenerateFieldServiceNotesReportFriday_ORIG = 0 AND GenerateFieldServiceNotesReportFriday = 1 Then
			CreateAuditLogEntry "Global Settings Change", "Global Settings Change", "Major", 1, "Field service notes report generation schedule for Friday turned on and set to run at " & GenerateFieldServiceNotesReportFridayTime
		Else
			CreateAuditLogEntry "Global Settings Change", "Global Settings Change", "Major", 1, "Field service notes report generation scheduled run time for Friday changed from " & GenerateFieldServiceNotesReportFridayTime_ORIG & " to " & GenerateFieldServiceNotesReportFridayTime
		End If
	End If



	If Request.Form("chkNoFieldServiceNotesReportSaturday") = "on" then GenerateFieldServiceNotesReportSaturdayMsg = "On" Else GenerateFieldServiceNotesReportSaturdayMsg = "Off"
	If GenerateFieldServiceNotesReportSaturday_ORIG = 1 then GenerateFieldServiceNotesReportSaturdayMsgOrig = "On" Else GenerateFieldServiceNotesReportSaturdayMsgOrig = "Off"
	
	If GenerateFieldServiceNotesReportSaturday <> GenerateFieldServiceNotesReportSaturday_ORIG Then
		CreateAuditLogEntry "Global Settings Change", "Global Settings Change", "Major", 1, "Field service notes report generation schedule for Saturday changed from " & GenerateFieldServiceNotesReportSaturdayMsgOrig & " to " & GenerateFieldServiceNotesReportSaturdayMsg
	End If
	
	If GenerateFieldServiceNotesReportSaturdayTime <> GenerateFieldServiceNotesReportSaturdayTime_ORIG Then
		If GenerateFieldServiceNotesReportSaturday_ORIG = 0 AND GenerateFieldServiceNotesReportSaturday = 1 Then
			CreateAuditLogEntry "Global Settings Change", "Global Settings Change", "Major", 1, "Field service notes report generation schedule for Saturday turned on and set to run at " & GenerateFieldServiceNotesReportSaturdayTime
		Else
			CreateAuditLogEntry "Global Settings Change", "Global Settings Change", "Major", 1, "Field service notes report generation scheduled run time for Saturday changed from " & GenerateFieldServiceNotesReportSaturdayTime_ORIG & " to " & GenerateFieldServiceNotesReportSaturdayTime
		End If
	End If


	If Request.Form("chkNoFieldServiceNotesReportIfClosed") = "on" then RunFieldServiceNotesReportIfClosedMsg = "On" Else RunFieldServiceNotesReportIfClosedMsg = "Off"
	If RunFieldServiceNotesReportIfClosed_ORIG = 1 then RunFieldServiceNotesReportIfClosedMsgOrig = "On" Else RunFieldServiceNotesReportIfClosedMsgOrig = "Off"
	
	If RunFieldServiceNotesReportIfClosed <> RunFieldServiceNotesReportIfClosed_ORIG Then
		CreateAuditLogEntry "Global Settings Change", "Global Settings Change", "Major", 1, "Field service notes report generation schedule - Don''t Run If Closed (Monday-Friday Only) -  changed from " & RunFieldServiceNotesReportIfClosedMsgOrig & " to " & RunFieldServiceNotesReportIfClosedMsg
	End If


	If Request.Form("chkNoFieldServiceNotesReportIfClosingEarly") = "on" then RunFieldServiceNotesReportIfClosingEarlyMsg = "On" Else RunFieldServiceNotesReportIfClosingEarlyMsg = "Off"
	If RunFieldServiceNotesReportIfClosingEarly_ORIG = 1 then RunFieldServiceNotesReportIfClosingEarlyMsgOrig = "On" Else RunFieldServiceNotesReportIfClosingEarlyMsgOrig = "Off"
	
	If RunFieldServiceNotesReportIfClosingEarly <> RunFieldServiceNotesReportIfClosingEarly_ORIG Then
		CreateAuditLogEntry "Global Settings Change", "Global Settings Change", "Major", 1, "Field service notes report generation schedule - Don''t Run If Closed (Monday-Friday Only) -  changed from " & RunFieldServiceNotesReportIfClosingEarlyMsgOrig & " to " & RunFieldServiceNotesReportIfClosingEarlyMsg
	End If


	'*********************************************************************
	'Build Array/String of Schedule Data From Request Form Field Data
	'*********************************************************************

	Schedule_FieldServiceNotesReportGenerationUpdated = ""
	
	Schedule_FieldServiceNotesReportGenerationUpdated = GenerateFieldServiceNotesReportSunday
	Schedule_FieldServiceNotesReportGenerationUpdated = Schedule_FieldServiceNotesReportGenerationUpdated & "," & GenerateFieldServiceNotesReportMonday
	Schedule_FieldServiceNotesReportGenerationUpdated = Schedule_FieldServiceNotesReportGenerationUpdated & "," & GenerateFieldServiceNotesReportTuesday
	Schedule_FieldServiceNotesReportGenerationUpdated = Schedule_FieldServiceNotesReportGenerationUpdated & "," & GenerateFieldServiceNotesReportWednesday
	Schedule_FieldServiceNotesReportGenerationUpdated = Schedule_FieldServiceNotesReportGenerationUpdated & "," & GenerateFieldServiceNotesReportThursday
	Schedule_FieldServiceNotesReportGenerationUpdated = Schedule_FieldServiceNotesReportGenerationUpdated & "," & GenerateFieldServiceNotesReportFriday
	Schedule_FieldServiceNotesReportGenerationUpdated = Schedule_FieldServiceNotesReportGenerationUpdated & "," & GenerateFieldServiceNotesReportSaturday
	
	Schedule_FieldServiceNotesReportGenerationUpdated = Schedule_FieldServiceNotesReportGenerationUpdated & "," & GenerateFieldServiceNotesReportSundayTime
	Schedule_FieldServiceNotesReportGenerationUpdated = Schedule_FieldServiceNotesReportGenerationUpdated & "," & GenerateFieldServiceNotesReportMondayTime
	Schedule_FieldServiceNotesReportGenerationUpdated = Schedule_FieldServiceNotesReportGenerationUpdated & "," & GenerateFieldServiceNotesReportTuesdayTime
	Schedule_FieldServiceNotesReportGenerationUpdated = Schedule_FieldServiceNotesReportGenerationUpdated & "," & GenerateFieldServiceNotesReportWednesdayTime
	Schedule_FieldServiceNotesReportGenerationUpdated = Schedule_FieldServiceNotesReportGenerationUpdated & "," & GenerateFieldServiceNotesReportThursdayTime
	Schedule_FieldServiceNotesReportGenerationUpdated = Schedule_FieldServiceNotesReportGenerationUpdated & "," & GenerateFieldServiceNotesReportFridayTime
	Schedule_FieldServiceNotesReportGenerationUpdated = Schedule_FieldServiceNotesReportGenerationUpdated & "," & GenerateFieldServiceNotesReportSaturdayTime

	
	Schedule_FieldServiceNotesReportGenerationUpdated = Schedule_FieldServiceNotesReportGenerationUpdated & "," & RunFieldServiceNotesReportIfClosed
	Schedule_FieldServiceNotesReportGenerationUpdated = Schedule_FieldServiceNotesReportGenerationUpdated & "," & RunFieldServiceNotesReportIfClosingEarly
	
	Response.Write("<br><br><br>Schedule_FieldServiceNotesReportGenerationUpdated: " & Schedule_FieldServiceNotesReportGenerationUpdated)

	'*********************************************************************
	'Update SQL with Array/String of Schedule Data
	'*********************************************************************
		
	SQL = "UPDATE Settings_FieldService SET Schedule_FieldServiceNotesReportGeneration = '" & cStr(Schedule_FieldServiceNotesReportGenerationUpdated) & "' "
	
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