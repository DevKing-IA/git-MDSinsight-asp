<!--#include file="../../../inc/header.asp"-->

<%

	'***********************************************************
	'Get Values Of All Form Fields Posted
	'***********************************************************
	
	GenerateFilterTicketSunday = Request.Form("chkNoAutoFilterTicketGenSunday")
	GenerateFilterTicketMonday = Request.Form("chkNoAutoFilterTicketGenMonday")
	GenerateFilterTicketTuesday = Request.Form("chkNoAutoFilterTicketGenTuesday")
	GenerateFilterTicketWednesday = Request.Form("chkNoAutoFilterTicketGenWednesday")
	GenerateFilterTicketThursday = Request.Form("chkNoAutoFilterTicketGenThursday")
	GenerateFilterTicketFriday = Request.Form("chkNoAutoFilterTicketGenFriday")
	GenerateFilterTicketSaturday = Request.Form("chkNoAutoFilterTicketGenSaturday")
	
	GenerateFilterTicketSundayTime = Request.Form("txtAutoFilterGenSchedulerSundayTime")
	GenerateFilterTicketMondayTime = Request.Form("txtAutoFilterGenSchedulerMondayTime")
	GenerateFilterTicketTuesdayTime = Request.Form("txtAutoFilterGenSchedulerTuesdayTime")
	GenerateFilterTicketWednesdayTime = Request.Form("txtAutoFilterGenSchedulerWednesdayTime")
	GenerateFilterTicketThursdayTime = Request.Form("txtAutoFilterGenSchedulerThursdayTime")
	GenerateFilterTicketFridayTime = Request.Form("txtAutoFilterGenSchedulerFridayTime")
	GenerateFilterTicketSaturdayTime = Request.Form("txtAutoFilterGenSchedulerSaturdayTime")
	
	RunFilterTicketAutoGenIfClosed = Request.Form("chkNoAutoFilterTicketGenIfClosed")
	RunFilterTicketAutoGenIfClosingEarly = Request.Form("chkNoAutoFilterTicketGenIfClosingEarly")


	If Request.Form("chkNoAutoFilterTicketGenSunday") = "on" Then
		GenerateFilterTicketSunday = 0
		GenerateFilterTicketSundayTime = ""
	Else 
		GenerateFilterTicketSunday = 1
	End If

	If Request.Form("chkNoAutoFilterTicketGenMonday") = "on" Then
		GenerateFilterTicketMonday = 0
		GenerateFilterTicketMondayTime = ""
	Else 
		GenerateFilterTicketMonday = 1
	End If

	If Request.Form("chkNoAutoFilterTicketGenTuesday") = "on" Then
		GenerateFilterTicketTuesday = 0
		GenerateFilterTicketTuesdayTime = ""
	Else 
		GenerateFilterTicketTuesday = 1
	End If

	If Request.Form("chkNoAutoFilterTicketGenWednesday") = "on" Then
		GenerateFilterTicketWednesday = 0
		GenerateFilterTicketWednesdayTime = ""
	Else 
		GenerateFilterTicketWednesday = 1
	End If

	If Request.Form("chkNoAutoFilterTicketGenThursday") = "on" Then
		GenerateFilterTicketThursday = 0
		GenerateFilterTicketThursdayTime = ""
	Else 
		GenerateFilterTicketThursday = 1
	End If

	If Request.Form("chkNoAutoFilterTicketGenFriday") = "on" Then
		GenerateFilterTicketFriday = 0
		GenerateFilterTicketFridayTime = ""
	Else 
		GenerateFilterTicketFriday = 1
	End If

	If Request.Form("chkNoAutoFilterTicketGenSaturday") = "on" Then
		GenerateFilterTicketSaturday = 0
		GenerateFilterTicketSaturdayTime = ""
	Else 
		GenerateFilterTicketSaturday = 1
	End If

	If Request.Form("chkNoAutoFilterTicketGenIfClosed") = "on" Then RunFilterTicketAutoGenIfClosed = 0 Else RunFilterTicketAutoGenIfClosed = 1
	If Request.Form("chkNoAutoFilterTicketGenIfClosingEarly") = "on" Then RunFilterTicketAutoGenIfClosingEarly = 0 Else RunFilterTicketAutoGenIfClosingEarly = 1
	
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
	
		Schedule_FilterGeneration = rsFieldServiceSettings("Schedule_FilterGeneration")
		
		Schedule_FilterGenerationSettings = Split(Schedule_FilterGeneration,",")

		GenerateFilterTicketSunday_ORIG = cInt(Schedule_FilterGenerationSettings(0))
		GenerateFilterTicketMonday_ORIG = cInt(Schedule_FilterGenerationSettings(1))
		GenerateFilterTicketTuesday_ORIG = cInt(Schedule_FilterGenerationSettings(2))
		GenerateFilterTicketWednesday_ORIG = cInt(Schedule_FilterGenerationSettings(3))
		GenerateFilterTicketThursday_ORIG = cInt(Schedule_FilterGenerationSettings(4))
		GenerateFilterTicketFriday_ORIG = cInt(Schedule_FilterGenerationSettings(5))
		GenerateFilterTicketSaturday_ORIG = cInt(Schedule_FilterGenerationSettings(6))
		GenerateFilterTicketSundayTime_ORIG = Schedule_FilterGenerationSettings(7)
		GenerateFilterTicketMondayTime_ORIG = Schedule_FilterGenerationSettings(8)
		GenerateFilterTicketTuesdayTime_ORIG = Schedule_FilterGenerationSettings(9)
		GenerateFilterTicketWednesdayTime_ORIG = Schedule_FilterGenerationSettings(10)
		GenerateFilterTicketThursdayTime_ORIG = Schedule_FilterGenerationSettings(11)
		GenerateFilterTicketFridayTime_ORIG = Schedule_FilterGenerationSettings(12)
		GenerateFilterTicketSaturdayTime_ORIG = Schedule_FilterGenerationSettings(13)
		RunFilterTicketAutoGenIfClosed_ORIG = cInt(Schedule_FilterGenerationSettings(14))
		RunFilterTicketAutoGenIfClosingEarly_ORIG = cInt(Schedule_FilterGenerationSettings(15))
	
	End If
	
	set rsFieldServiceSettings = Nothing
	cnnFieldServiceSettings.close
	set cnnFieldServiceSettings = Nothing
	

	'***********************************************************
	'Perform Audit Trail Entry Inserts
	'***********************************************************

	
	If Request.Form("chkNoAutoFilterTicketGenSunday") = "on" then GenerateFilterTicketSundayMsg = "On" Else GenerateFilterTicketSundayMsg = "Off"
	If GenerateFilterTicketSunday_ORIG = 1 then GenerateFilterTicketSundayMsgOrig = "On" Else GenerateFilterTicketSundayMsgOrig = "Off"
	
	If GenerateFilterTicketSunday <> GenerateFilterTicketSunday_ORIG Then
		CreateAuditLogEntry "Global Settings Change", "Global Settings Change", "Major", 1, "Automated filter change ticket generation schedule for SUNDAY changed from " & GenerateFilterTicketSundayMsgOrig & " to " & GenerateFilterTicketSundayMsg
	End If
	
	If GenerateFilterTicketSundayTime <> GenerateFilterTicketSundayTime_ORIG Then
		If GenerateFilterTicketSunday_ORIG = 0 AND GenerateFilterTicketSunday = 1 Then
			CreateAuditLogEntry "Global Settings Change", "Global Settings Change", "Major", 1, "Automated filter change ticket generation schedule for SUNDAY turned on and set to run at " & GenerateFilterTicketSundayTime
		Else
			CreateAuditLogEntry "Global Settings Change", "Global Settings Change", "Major", 1, "Automated filter change ticket generation scheduled run time for SUNDAY changed from " & GenerateFilterTicketSundayTime_ORIG & " to " & GenerateFilterTicketSundayTime
		End If
	End If
	
	

	If Request.Form("chkNoAutoFilterTicketGenMonday") = "on" then GenerateFilterTicketMondayMsg = "On" Else GenerateFilterTicketMondayMsg = "Off"
	If GenerateFilterTicketMonday_ORIG = 1 then GenerateFilterTicketMondayMsgOrig = "On" Else GenerateFilterTicketMondayMsgOrig = "Off"
	
	If GenerateFilterTicketMonday <> GenerateFilterTicketMonday_ORIG Then
		CreateAuditLogEntry "Global Settings Change", "Global Settings Change", "Major", 1, "Automated filter change ticket generation schedule for Monday changed from " & GenerateFilterTicketMondayMsgOrig & " to " & GenerateFilterTicketMondayMsg
	End If
	
	If GenerateFilterTicketMondayTime <> GenerateFilterTicketMondayTime_ORIG Then
		If GenerateFilterTicketMonday_ORIG = 0 AND GenerateFilterTicketMonday = 1 Then
			CreateAuditLogEntry "Global Settings Change", "Global Settings Change", "Major", 1, "Automated filter change ticket generation schedule for Monday turned on and set to run at " & GenerateFilterTicketMondayTime
		Else
			CreateAuditLogEntry "Global Settings Change", "Global Settings Change", "Major", 1, "Automated filter change ticket generation scheduled run time for Monday changed from " & GenerateFilterTicketMondayTime_ORIG & " to " & GenerateFilterTicketMondayTime
		End If
	End If
	


	If Request.Form("chkNoAutoFilterTicketGenTuesday") = "on" then GenerateFilterTicketTuesdayMsg = "On" Else GenerateFilterTicketTuesdayMsg = "Off"
	If GenerateFilterTicketTuesday_ORIG = 1 then GenerateFilterTicketTuesdayMsgOrig = "On" Else GenerateFilterTicketTuesdayMsgOrig = "Off"
	
	If GenerateFilterTicketTuesday <> GenerateFilterTicketTuesday_ORIG Then
		CreateAuditLogEntry "Global Settings Change", "Global Settings Change", "Major", 1, "Automated filter change ticket generation schedule for Tuesday changed from " & GenerateFilterTicketTuesdayMsgOrig & " to " & GenerateFilterTicketTuesdayMsg
	End If
	
	If GenerateFilterTicketTuesdayTime <> GenerateFilterTicketTuesdayTime_ORIG Then
		If GenerateFilterTicketTuesday_ORIG = 0 AND GenerateFilterTicketTuesday = 1 Then
			CreateAuditLogEntry "Global Settings Change", "Global Settings Change", "Major", 1, "Automated filter change ticket generation schedule for Tuesday turned on and set to run at " & GenerateFilterTicketTuesdayTime
		Else
			CreateAuditLogEntry "Global Settings Change", "Global Settings Change", "Major", 1, "Automated filter change ticket generation scheduled run time for Tuesday changed from " & GenerateFilterTicketTuesdayTime_ORIG & " to " & GenerateFilterTicketTuesdayTime
		End If
	End If



	If Request.Form("chkNoAutoFilterTicketGenWednesday") = "on" then GenerateFilterTicketWednesdayMsg = "On" Else GenerateFilterTicketWednesdayMsg = "Off"
	If GenerateFilterTicketWednesday_ORIG = 1 then GenerateFilterTicketWednesdayMsgOrig = "On" Else GenerateFilterTicketWednesdayMsgOrig = "Off"
	
	If GenerateFilterTicketWednesday <> GenerateFilterTicketWednesday_ORIG Then
		CreateAuditLogEntry "Global Settings Change", "Global Settings Change", "Major", 1, "Automated filter change ticket generation schedule for Wednesday changed from " & GenerateFilterTicketWednesdayMsgOrig & " to " & GenerateFilterTicketWednesdayMsg
	End If
	
	If GenerateFilterTicketWednesdayTime <> GenerateFilterTicketWednesdayTime_ORIG Then
		If GenerateFilterTicketWednesday_ORIG = 0 AND GenerateFilterTicketWednesday = 1 Then
			CreateAuditLogEntry "Global Settings Change", "Global Settings Change", "Major", 1, "Automated filter change ticket generation schedule for Wednesday turned on and set to run at " & GenerateFilterTicketWednesdayTime
		Else
			CreateAuditLogEntry "Global Settings Change", "Global Settings Change", "Major", 1, "Automated filter change ticket generation scheduled run time for Wednesday changed from " & GenerateFilterTicketWednesdayTime_ORIG & " to " & GenerateFilterTicketWednesdayTime
		End If
	End If



	If Request.Form("chkNoAutoFilterTicketGenThursday") = "on" then GenerateFilterTicketThursdayMsg = "On" Else GenerateFilterTicketThursdayMsg = "Off"
	If GenerateFilterTicketThursday_ORIG = 1 then GenerateFilterTicketThursdayMsgOrig = "On" Else GenerateFilterTicketThursdayMsgOrig = "Off"
	
	If GenerateFilterTicketThursday <> GenerateFilterTicketThursday_ORIG Then
		CreateAuditLogEntry "Global Settings Change", "Global Settings Change", "Major", 1, "Automated filter change ticket generation schedule for Thursday changed from " & GenerateFilterTicketThursdayMsgOrig & " to " & GenerateFilterTicketThursdayMsg
	End If
	
	If GenerateFilterTicketThursdayTime <> GenerateFilterTicketThursdayTime_ORIG Then
		If GenerateFilterTicketThursday_ORIG = 0 AND GenerateFilterTicketThursday = 1 Then
			CreateAuditLogEntry "Global Settings Change", "Global Settings Change", "Major", 1, "Automated filter change ticket generation schedule for Thursday turned on and set to run at " & GenerateFilterTicketThursdayTime
		Else
			CreateAuditLogEntry "Global Settings Change", "Global Settings Change", "Major", 1, "Automated filter change ticket generation scheduled run time for Thursday changed from " & GenerateFilterTicketThursdayTime_ORIG & " to " & GenerateFilterTicketThursdayTime
		End If
	End If



	If Request.Form("chkNoAutoFilterTicketGenFriday") = "on" then GenerateFilterTicketFridayMsg = "On" Else GenerateFilterTicketFridayMsg = "Off"
	If GenerateFilterTicketFriday_ORIG = 1 then GenerateFilterTicketFridayMsgOrig = "On" Else GenerateFilterTicketFridayMsgOrig = "Off"
	
	If GenerateFilterTicketFriday <> GenerateFilterTicketFriday_ORIG Then
		CreateAuditLogEntry "Global Settings Change", "Global Settings Change", "Major", 1, "Automated filter change ticket generation schedule for Friday changed from " & GenerateFilterTicketFridayMsgOrig & " to " & GenerateFilterTicketFridayMsg
	End If
	
	If GenerateFilterTicketFridayTime <> GenerateFilterTicketFridayTime_ORIG Then
		If GenerateFilterTicketFriday_ORIG = 0 AND GenerateFilterTicketFriday = 1 Then
			CreateAuditLogEntry "Global Settings Change", "Global Settings Change", "Major", 1, "Automated filter change ticket generation schedule for Friday turned on and set to run at " & GenerateFilterTicketFridayTime
		Else
			CreateAuditLogEntry "Global Settings Change", "Global Settings Change", "Major", 1, "Automated filter change ticket generation scheduled run time for Friday changed from " & GenerateFilterTicketFridayTime_ORIG & " to " & GenerateFilterTicketFridayTime
		End If
	End If



	If Request.Form("chkNoAutoFilterTicketGenSaturday") = "on" then GenerateFilterTicketSaturdayMsg = "On" Else GenerateFilterTicketSaturdayMsg = "Off"
	If GenerateFilterTicketSaturday_ORIG = 1 then GenerateFilterTicketSaturdayMsgOrig = "On" Else GenerateFilterTicketSaturdayMsgOrig = "Off"
	
	If GenerateFilterTicketSaturday <> GenerateFilterTicketSaturday_ORIG Then
		CreateAuditLogEntry "Global Settings Change", "Global Settings Change", "Major", 1, "Automated filter change ticket generation schedule for Saturday changed from " & GenerateFilterTicketSaturdayMsgOrig & " to " & GenerateFilterTicketSaturdayMsg
	End If
	
	If GenerateFilterTicketSaturdayTime <> GenerateFilterTicketSaturdayTime_ORIG Then
		If GenerateFilterTicketSaturday_ORIG = 0 AND GenerateFilterTicketSaturday = 1 Then
			CreateAuditLogEntry "Global Settings Change", "Global Settings Change", "Major", 1, "Automated filter change ticket generation schedule for Saturday turned on and set to run at " & GenerateFilterTicketSaturdayTime
		Else
			CreateAuditLogEntry "Global Settings Change", "Global Settings Change", "Major", 1, "Automated filter change ticket generation scheduled run time for Saturday changed from " & GenerateFilterTicketSaturdayTime_ORIG & " to " & GenerateFilterTicketSaturdayTime
		End If
	End If


	If Request.Form("chkNoAutoFilterTicketGenIfClosed") = "on" then RunFilterTicketAutoGenIfClosedMsg = "On" Else RunFilterTicketAutoGenIfClosedMsg = "Off"
	If RunFilterTicketAutoGenIfClosed_ORIG = 1 then RunFilterTicketAutoGenIfClosedMsgOrig = "On" Else RunFilterTicketAutoGenIfClosedMsgOrig = "Off"
	
	If RunFilterTicketAutoGenIfClosed <> RunFilterTicketAutoGenIfClosed_ORIG Then
		CreateAuditLogEntry "Global Settings Change", "Global Settings Change", "Major", 1, "Automated filter change ticket generation schedule - Don''t Run If Closed (Monday-Friday Only) -  changed from " & RunFilterTicketAutoGenIfClosedMsgOrig & " to " & RunFilterTicketAutoGenIfClosedMsg
	End If


	If Request.Form("chkNoAutoFilterTicketGenIfClosingEarly") = "on" then RunFilterTicketAutoGenIfClosingEarlyMsg = "On" Else RunFilterTicketAutoGenIfClosingEarlyMsg = "Off"
	If RunFilterTicketAutoGenIfClosingEarly_ORIG = 1 then RunFilterTicketAutoGenIfClosingEarlyMsgOrig = "On" Else RunFilterTicketAutoGenIfClosingEarlyMsgOrig = "Off"
	
	If RunFilterTicketAutoGenIfClosingEarly <> RunFilterTicketAutoGenIfClosingEarly_ORIG Then
		CreateAuditLogEntry "Global Settings Change", "Global Settings Change", "Major", 1, "Automated filter change ticket generation schedule - Don''t Run If Closed (Monday-Friday Only) -  changed from " & RunFilterTicketAutoGenIfClosingEarlyMsgOrig & " to " & RunFilterTicketAutoGenIfClosingEarlyMsg
	End If


	'*********************************************************************
	'Build Array/String of Schedule Data From Request Form Field Data
	'*********************************************************************

	Schedule_FilterGenerationUpdated = ""
	
	Schedule_FilterGenerationUpdated = GenerateFilterTicketSunday
	Schedule_FilterGenerationUpdated = Schedule_FilterGenerationUpdated & "," & GenerateFilterTicketMonday
	Schedule_FilterGenerationUpdated = Schedule_FilterGenerationUpdated & "," & GenerateFilterTicketTuesday
	Schedule_FilterGenerationUpdated = Schedule_FilterGenerationUpdated & "," & GenerateFilterTicketWednesday
	Schedule_FilterGenerationUpdated = Schedule_FilterGenerationUpdated & "," & GenerateFilterTicketThursday
	Schedule_FilterGenerationUpdated = Schedule_FilterGenerationUpdated & "," & GenerateFilterTicketFriday
	Schedule_FilterGenerationUpdated = Schedule_FilterGenerationUpdated & "," & GenerateFilterTicketSaturday
	
	Schedule_FilterGenerationUpdated = Schedule_FilterGenerationUpdated & "," & GenerateFilterTicketSundayTime
	Schedule_FilterGenerationUpdated = Schedule_FilterGenerationUpdated & "," & GenerateFilterTicketMondayTime
	Schedule_FilterGenerationUpdated = Schedule_FilterGenerationUpdated & "," & GenerateFilterTicketTuesdayTime
	Schedule_FilterGenerationUpdated = Schedule_FilterGenerationUpdated & "," & GenerateFilterTicketWednesdayTime
	Schedule_FilterGenerationUpdated = Schedule_FilterGenerationUpdated & "," & GenerateFilterTicketThursdayTime
	Schedule_FilterGenerationUpdated = Schedule_FilterGenerationUpdated & "," & GenerateFilterTicketFridayTime
	Schedule_FilterGenerationUpdated = Schedule_FilterGenerationUpdated & "," & GenerateFilterTicketSaturdayTime

	
	Schedule_FilterGenerationUpdated = Schedule_FilterGenerationUpdated & "," & RunFilterTicketAutoGenIfClosed
	Schedule_FilterGenerationUpdated = Schedule_FilterGenerationUpdated & "," & RunFilterTicketAutoGenIfClosingEarly
	
	Response.Write("<br><br><br>Schedule_FilterGenerationUpdated: " & Schedule_FilterGenerationUpdated)

	'*********************************************************************
	'Update SQL with Array/String of Schedule Data
	'*********************************************************************
		
	SQL = "UPDATE Settings_FieldService SET Schedule_FilterGeneration = '" & cStr(Schedule_FilterGenerationUpdated) & "' "
	
	Response.Write("<br><br><br>SQL: " & SQL)
	
	Set cnn8 = Server.CreateObject("ADODB.Connection")
	cnn8.open (Session("ClientCnnString"))
	Set rs = Server.CreateObject("ADODB.Recordset")
	rs.CursorLocation = 3 
	Set rs = cnn8.Execute(SQL)
	
	set rs = Nothing
	cnn8.close
	set cnn8 = Nothing


	 Response.Redirect("filter-changes.asp")
	
%><!--#include file="../../../inc/footer-main.asp"-->