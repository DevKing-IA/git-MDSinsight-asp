<!--#include file="../../../../inc/header.asp"-->

<%

	'***********************************************************
	'Get Values Of All Form Fields Posted 
	'***********************************************************
	
	GenerateEquipmentNeedToKnowReportSunday = Request.Form("chkNoEquipmentNeedToKnowReportSunday")
	GenerateEquipmentNeedToKnowReportMonday = Request.Form("chkNoEquipmentNeedToKnowReportMonday")
	GenerateEquipmentNeedToKnowReportTuesday = Request.Form("chkNoEquipmentNeedToKnowReportTuesday")
	GenerateEquipmentNeedToKnowReportWednesday = Request.Form("chkNoEquipmentNeedToKnowReportWednesday")
	GenerateEquipmentNeedToKnowReportThursday = Request.Form("chkNoEquipmentNeedToKnowReportThursday")
	GenerateEquipmentNeedToKnowReportFriday = Request.Form("chkNoEquipmentNeedToKnowReportFriday")
	GenerateEquipmentNeedToKnowReportSaturday = Request.Form("chkNoEquipmentNeedToKnowReportSaturday")
	
	GenerateEquipmentNeedToKnowReportSundayTime = Request.Form("txtEquipmentNeedToKnowReportSchedulerSundayTime")
	GenerateEquipmentNeedToKnowReportMondayTime = Request.Form("txtEquipmentNeedToKnowReportSchedulerMondayTime")
	GenerateEquipmentNeedToKnowReportTuesdayTime = Request.Form("txtEquipmentNeedToKnowReportSchedulerTuesdayTime")
	GenerateEquipmentNeedToKnowReportWednesdayTime = Request.Form("txtEquipmentNeedToKnowReportSchedulerWednesdayTime")
	GenerateEquipmentNeedToKnowReportThursdayTime = Request.Form("txtEquipmentNeedToKnowReportSchedulerThursdayTime")
	GenerateEquipmentNeedToKnowReportFridayTime = Request.Form("txtEquipmentNeedToKnowReportSchedulerFridayTime")
	GenerateEquipmentNeedToKnowReportSaturdayTime = Request.Form("txtEquipmentNeedToKnowReportSchedulerSaturdayTime")
	
	RunEquipmentNeedToKnowReportIfClosed = Request.Form("chkNoEquipmentNeedToKnowReportIfClosed")
	RunEquipmentNeedToKnowReportIfClosingEarly = Request.Form("chkNoEquipmentNeedToKnowReportIfClosingEarly")


	If Request.Form("chkNoEquipmentNeedToKnowReportSunday") = "on" Then
		GenerateEquipmentNeedToKnowReportSunday = 0
		GenerateEquipmentNeedToKnowReportSundayTime = ""
	Else 
		GenerateEquipmentNeedToKnowReportSunday = 1
	End If

	If Request.Form("chkNoEquipmentNeedToKnowReportMonday") = "on" Then
		GenerateEquipmentNeedToKnowReportMonday = 0
		GenerateEquipmentNeedToKnowReportMondayTime = ""
	Else 
		GenerateEquipmentNeedToKnowReportMonday = 1
	End If

	If Request.Form("chkNoEquipmentNeedToKnowReportTuesday") = "on" Then
		GenerateEquipmentNeedToKnowReportTuesday = 0
		GenerateEquipmentNeedToKnowReportTuesdayTime = ""
	Else 
		GenerateEquipmentNeedToKnowReportTuesday = 1
	End If

	If Request.Form("chkNoEquipmentNeedToKnowReportWednesday") = "on" Then
		GenerateEquipmentNeedToKnowReportWednesday = 0
		GenerateEquipmentNeedToKnowReportWednesdayTime = ""
	Else 
		GenerateEquipmentNeedToKnowReportWednesday = 1
	End If

	If Request.Form("chkNoEquipmentNeedToKnowReportThursday") = "on" Then
		GenerateEquipmentNeedToKnowReportThursday = 0
		GenerateEquipmentNeedToKnowReportThursdayTime = ""
	Else 
		GenerateEquipmentNeedToKnowReportThursday = 1
	End If

	If Request.Form("chkNoEquipmentNeedToKnowReportFriday") = "on" Then
		GenerateEquipmentNeedToKnowReportFriday = 0
		GenerateEquipmentNeedToKnowReportFridayTime = ""
	Else 
		GenerateEquipmentNeedToKnowReportFriday = 1
	End If

	If Request.Form("chkNoEquipmentNeedToKnowReportSaturday") = "on" Then
		GenerateEquipmentNeedToKnowReportSaturday = 0
		GenerateEquipmentNeedToKnowReportSaturdayTime = ""
	Else 
		GenerateEquipmentNeedToKnowReportSaturday = 1
	End If

	If Request.Form("chkNoEquipmentNeedToKnowReportIfClosed") = "on" Then RunEquipmentNeedToKnowReportIfClosed = 0 Else RunEquipmentNeedToKnowReportIfClosed = 1
	If Request.Form("chkNoEquipmentNeedToKnowReportIfClosingEarly") = "on" Then RunEquipmentNeedToKnowReportIfClosingEarly = 0 Else RunEquipmentNeedToKnowReportIfClosingEarly = 1
	
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
	
		Schedule_EquipmentNeedToKnowReportGeneration = rsPropsectingSettings("Schedule_EquipmentNeedToKnowReportGeneration")
		
		Schedule_EquipmentNeedToKnowReportGenerationSettings = Split(Schedule_EquipmentNeedToKnowReportGeneration,",")

		GenerateEquipmentNeedToKnowReportSunday_ORIG = cInt(Schedule_EquipmentNeedToKnowReportGenerationSettings(0))
		GenerateEquipmentNeedToKnowReportMonday_ORIG = cInt(Schedule_EquipmentNeedToKnowReportGenerationSettings(1))
		GenerateEquipmentNeedToKnowReportTuesday_ORIG = cInt(Schedule_EquipmentNeedToKnowReportGenerationSettings(2))
		GenerateEquipmentNeedToKnowReportWednesday_ORIG = cInt(Schedule_EquipmentNeedToKnowReportGenerationSettings(3))
		GenerateEquipmentNeedToKnowReportThursday_ORIG = cInt(Schedule_EquipmentNeedToKnowReportGenerationSettings(4))
		GenerateEquipmentNeedToKnowReportFriday_ORIG = cInt(Schedule_EquipmentNeedToKnowReportGenerationSettings(5))
		GenerateEquipmentNeedToKnowReportSaturday_ORIG = cInt(Schedule_EquipmentNeedToKnowReportGenerationSettings(6))
		GenerateEquipmentNeedToKnowReportSundayTime_ORIG = Schedule_EquipmentNeedToKnowReportGenerationSettings(7)
		GenerateEquipmentNeedToKnowReportMondayTime_ORIG = Schedule_EquipmentNeedToKnowReportGenerationSettings(8)
		GenerateEquipmentNeedToKnowReportTuesdayTime_ORIG = Schedule_EquipmentNeedToKnowReportGenerationSettings(9)
		GenerateEquipmentNeedToKnowReportWednesdayTime_ORIG = Schedule_EquipmentNeedToKnowReportGenerationSettings(10)
		GenerateEquipmentNeedToKnowReportThursdayTime_ORIG = Schedule_EquipmentNeedToKnowReportGenerationSettings(11)
		GenerateEquipmentNeedToKnowReportFridayTime_ORIG = Schedule_EquipmentNeedToKnowReportGenerationSettings(12)
		GenerateEquipmentNeedToKnowReportSaturdayTime_ORIG = Schedule_EquipmentNeedToKnowReportGenerationSettings(13)
		RunEquipmentNeedToKnowReportIfClosed_ORIG = cInt(Schedule_EquipmentNeedToKnowReportGenerationSettings(14))
		RunEquipmentNeedToKnowReportIfClosingEarly_ORIG = cInt(Schedule_EquipmentNeedToKnowReportGenerationSettings(15))
	
	End If
	
	set rsPropsectingSettings = Nothing
	cnnPropsectingSettings.close
	set cnnPropsectingSettings = Nothing
	

	'***********************************************************
	'Perform Audit Trail Entry Inserts
	'***********************************************************

	
	If Request.Form("chkNoEquipmentNeedToKnowReportSunday") = "on" then GenerateEquipmentNeedToKnowReportSundayMsg = "On" Else GenerateEquipmentNeedToKnowReportSundayMsg = "Off"
	If GenerateEquipmentNeedToKnowReportSunday_ORIG = 1 then GenerateEquipmentNeedToKnowReportSundayMsgOrig = "On" Else GenerateEquipmentNeedToKnowReportSundayMsgOrig = "Off"
	
	If GenerateEquipmentNeedToKnowReportSunday <> GenerateEquipmentNeedToKnowReportSunday_ORIG Then
		CreateAuditLogEntry "Global Settings Change", "Global Settings Change", "Major", 1, "Equipment Need To Know Report generation schedule for SUNDAY changed from " & GenerateEquipmentNeedToKnowReportSundayMsgOrig & " to " & GenerateEquipmentNeedToKnowReportSundayMsg
	End If
	
	If GenerateEquipmentNeedToKnowReportSundayTime <> GenerateEquipmentNeedToKnowReportSundayTime_ORIG Then
		If GenerateEquipmentNeedToKnowReportSunday_ORIG = 0 AND GenerateEquipmentNeedToKnowReportSunday = 1 Then
			CreateAuditLogEntry "Global Settings Change", "Global Settings Change", "Major", 1, "Equipment Need To Know Report generation schedule for SUNDAY turned on and set to run at " & GenerateEquipmentNeedToKnowReportSundayTime
		Else
			CreateAuditLogEntry "Global Settings Change", "Global Settings Change", "Major", 1, "Equipment Need To Know Report generation scheduled run time for SUNDAY changed from " & GenerateEquipmentNeedToKnowReportSundayTime_ORIG & " to " & GenerateEquipmentNeedToKnowReportSundayTime
		End If
	End If
	
	

	If Request.Form("chkNoEquipmentNeedToKnowReportMonday") = "on" then GenerateEquipmentNeedToKnowReportMondayMsg = "On" Else GenerateEquipmentNeedToKnowReportMondayMsg = "Off"
	If GenerateEquipmentNeedToKnowReportMonday_ORIG = 1 then GenerateEquipmentNeedToKnowReportMondayMsgOrig = "On" Else GenerateEquipmentNeedToKnowReportMondayMsgOrig = "Off"
	
	If GenerateEquipmentNeedToKnowReportMonday <> GenerateEquipmentNeedToKnowReportMonday_ORIG Then
		CreateAuditLogEntry "Global Settings Change", "Global Settings Change", "Major", 1, "Equipment Need To Know Report generation schedule for Monday changed from " & GenerateEquipmentNeedToKnowReportMondayMsgOrig & " to " & GenerateEquipmentNeedToKnowReportMondayMsg
	End If
	
	If GenerateEquipmentNeedToKnowReportMondayTime <> GenerateEquipmentNeedToKnowReportMondayTime_ORIG Then
		If GenerateEquipmentNeedToKnowReportMonday_ORIG = 0 AND GenerateEquipmentNeedToKnowReportMonday = 1 Then
			CreateAuditLogEntry "Global Settings Change", "Global Settings Change", "Major", 1, "Equipment Need To Know Report generation schedule for Monday turned on and set to run at " & GenerateEquipmentNeedToKnowReportMondayTime
		Else
			CreateAuditLogEntry "Global Settings Change", "Global Settings Change", "Major", 1, "Equipment Need To Know Report generation scheduled run time for Monday changed from " & GenerateEquipmentNeedToKnowReportMondayTime_ORIG & " to " & GenerateEquipmentNeedToKnowReportMondayTime
		End If
	End If
	


	If Request.Form("chkNoEquipmentNeedToKnowReportTuesday") = "on" then GenerateEquipmentNeedToKnowReportTuesdayMsg = "On" Else GenerateEquipmentNeedToKnowReportTuesdayMsg = "Off"
	If GenerateEquipmentNeedToKnowReportTuesday_ORIG = 1 then GenerateEquipmentNeedToKnowReportTuesdayMsgOrig = "On" Else GenerateEquipmentNeedToKnowReportTuesdayMsgOrig = "Off"
	
	If GenerateEquipmentNeedToKnowReportTuesday <> GenerateEquipmentNeedToKnowReportTuesday_ORIG Then
		CreateAuditLogEntry "Global Settings Change", "Global Settings Change", "Major", 1, "Equipment Need To Know Report generation schedule for Tuesday changed from " & GenerateEquipmentNeedToKnowReportTuesdayMsgOrig & " to " & GenerateEquipmentNeedToKnowReportTuesdayMsg
	End If
	
	If GenerateEquipmentNeedToKnowReportTuesdayTime <> GenerateEquipmentNeedToKnowReportTuesdayTime_ORIG Then
		If GenerateEquipmentNeedToKnowReportTuesday_ORIG = 0 AND GenerateEquipmentNeedToKnowReportTuesday = 1 Then
			CreateAuditLogEntry "Global Settings Change", "Global Settings Change", "Major", 1, "Equipment Need To Know Report generation schedule for Tuesday turned on and set to run at " & GenerateEquipmentNeedToKnowReportTuesdayTime
		Else
			CreateAuditLogEntry "Global Settings Change", "Global Settings Change", "Major", 1, "Equipment Need To Know Report generation scheduled run time for Tuesday changed from " & GenerateEquipmentNeedToKnowReportTuesdayTime_ORIG & " to " & GenerateEquipmentNeedToKnowReportTuesdayTime
		End If
	End If



	If Request.Form("chkNoEquipmentNeedToKnowReportWednesday") = "on" then GenerateEquipmentNeedToKnowReportWednesdayMsg = "On" Else GenerateEquipmentNeedToKnowReportWednesdayMsg = "Off"
	If GenerateEquipmentNeedToKnowReportWednesday_ORIG = 1 then GenerateEquipmentNeedToKnowReportWednesdayMsgOrig = "On" Else GenerateEquipmentNeedToKnowReportWednesdayMsgOrig = "Off"
	
	If GenerateEquipmentNeedToKnowReportWednesday <> GenerateEquipmentNeedToKnowReportWednesday_ORIG Then
		CreateAuditLogEntry "Global Settings Change", "Global Settings Change", "Major", 1, "Equipment Need To Know Report generation schedule for Wednesday changed from " & GenerateEquipmentNeedToKnowReportWednesdayMsgOrig & " to " & GenerateEquipmentNeedToKnowReportWednesdayMsg
	End If
	
	If GenerateEquipmentNeedToKnowReportWednesdayTime <> GenerateEquipmentNeedToKnowReportWednesdayTime_ORIG Then
		If GenerateEquipmentNeedToKnowReportWednesday_ORIG = 0 AND GenerateEquipmentNeedToKnowReportWednesday = 1 Then
			CreateAuditLogEntry "Global Settings Change", "Global Settings Change", "Major", 1, "Equipment Need To Know Report generation schedule for Wednesday turned on and set to run at " & GenerateEquipmentNeedToKnowReportWednesdayTime
		Else
			CreateAuditLogEntry "Global Settings Change", "Global Settings Change", "Major", 1, "Equipment Need To Know Report generation scheduled run time for Wednesday changed from " & GenerateEquipmentNeedToKnowReportWednesdayTime_ORIG & " to " & GenerateEquipmentNeedToKnowReportWednesdayTime
		End If
	End If



	If Request.Form("chkNoEquipmentNeedToKnowReportThursday") = "on" then GenerateEquipmentNeedToKnowReportThursdayMsg = "On" Else GenerateEquipmentNeedToKnowReportThursdayMsg = "Off"
	If GenerateEquipmentNeedToKnowReportThursday_ORIG = 1 then GenerateEquipmentNeedToKnowReportThursdayMsgOrig = "On" Else GenerateEquipmentNeedToKnowReportThursdayMsgOrig = "Off"
	
	If GenerateEquipmentNeedToKnowReportThursday <> GenerateEquipmentNeedToKnowReportThursday_ORIG Then
		CreateAuditLogEntry "Global Settings Change", "Global Settings Change", "Major", 1, "Equipment Need To Know Report generation schedule for Thursday changed from " & GenerateEquipmentNeedToKnowReportThursdayMsgOrig & " to " & GenerateEquipmentNeedToKnowReportThursdayMsg
	End If
	
	If GenerateEquipmentNeedToKnowReportThursdayTime <> GenerateEquipmentNeedToKnowReportThursdayTime_ORIG Then
		If GenerateEquipmentNeedToKnowReportThursday_ORIG = 0 AND GenerateEquipmentNeedToKnowReportThursday = 1 Then
			CreateAuditLogEntry "Global Settings Change", "Global Settings Change", "Major", 1, "Equipment Need To Know Report generation schedule for Thursday turned on and set to run at " & GenerateEquipmentNeedToKnowReportThursdayTime
		Else
			CreateAuditLogEntry "Global Settings Change", "Global Settings Change", "Major", 1, "Equipment Need To Know Report generation scheduled run time for Thursday changed from " & GenerateEquipmentNeedToKnowReportThursdayTime_ORIG & " to " & GenerateEquipmentNeedToKnowReportThursdayTime
		End If
	End If



	If Request.Form("chkNoEquipmentNeedToKnowReportFriday") = "on" then GenerateEquipmentNeedToKnowReportFridayMsg = "On" Else GenerateEquipmentNeedToKnowReportFridayMsg = "Off"
	If GenerateEquipmentNeedToKnowReportFriday_ORIG = 1 then GenerateEquipmentNeedToKnowReportFridayMsgOrig = "On" Else GenerateEquipmentNeedToKnowReportFridayMsgOrig = "Off"
	
	If GenerateEquipmentNeedToKnowReportFriday <> GenerateEquipmentNeedToKnowReportFriday_ORIG Then
		CreateAuditLogEntry "Global Settings Change", "Global Settings Change", "Major", 1, "Equipment Need To Know Report generation schedule for Friday changed from " & GenerateEquipmentNeedToKnowReportFridayMsgOrig & " to " & GenerateEquipmentNeedToKnowReportFridayMsg
	End If
	
	If GenerateEquipmentNeedToKnowReportFridayTime <> GenerateEquipmentNeedToKnowReportFridayTime_ORIG Then
		If GenerateEquipmentNeedToKnowReportFriday_ORIG = 0 AND GenerateEquipmentNeedToKnowReportFriday = 1 Then
			CreateAuditLogEntry "Global Settings Change", "Global Settings Change", "Major", 1, "Equipment Need To Know Report generation schedule for Friday turned on and set to run at " & GenerateEquipmentNeedToKnowReportFridayTime
		Else
			CreateAuditLogEntry "Global Settings Change", "Global Settings Change", "Major", 1, "Equipment Need To Know Report generation scheduled run time for Friday changed from " & GenerateEquipmentNeedToKnowReportFridayTime_ORIG & " to " & GenerateEquipmentNeedToKnowReportFridayTime
		End If
	End If



	If Request.Form("chkNoEquipmentNeedToKnowReportSaturday") = "on" then GenerateEquipmentNeedToKnowReportSaturdayMsg = "On" Else GenerateEquipmentNeedToKnowReportSaturdayMsg = "Off"
	If GenerateEquipmentNeedToKnowReportSaturday_ORIG = 1 then GenerateEquipmentNeedToKnowReportSaturdayMsgOrig = "On" Else GenerateEquipmentNeedToKnowReportSaturdayMsgOrig = "Off"
	
	If GenerateEquipmentNeedToKnowReportSaturday <> GenerateEquipmentNeedToKnowReportSaturday_ORIG Then
		CreateAuditLogEntry "Global Settings Change", "Global Settings Change", "Major", 1, "Equipment Need To Know Report generation schedule for Saturday changed from " & GenerateEquipmentNeedToKnowReportSaturdayMsgOrig & " to " & GenerateEquipmentNeedToKnowReportSaturdayMsg
	End If
	
	If GenerateEquipmentNeedToKnowReportSaturdayTime <> GenerateEquipmentNeedToKnowReportSaturdayTime_ORIG Then
		If GenerateEquipmentNeedToKnowReportSaturday_ORIG = 0 AND GenerateEquipmentNeedToKnowReportSaturday = 1 Then
			CreateAuditLogEntry "Global Settings Change", "Global Settings Change", "Major", 1, "Equipment Need To Know Report generation schedule for Saturday turned on and set to run at " & GenerateEquipmentNeedToKnowReportSaturdayTime
		Else
			CreateAuditLogEntry "Global Settings Change", "Global Settings Change", "Major", 1, "Equipment Need To Know Report generation scheduled run time for Saturday changed from " & GenerateEquipmentNeedToKnowReportSaturdayTime_ORIG & " to " & GenerateEquipmentNeedToKnowReportSaturdayTime
		End If
	End If


	If Request.Form("chkNoEquipmentNeedToKnowReportIfClosed") = "on" then RunEquipmentNeedToKnowReportIfClosedMsg = "On" Else RunEquipmentNeedToKnowReportIfClosedMsg = "Off"
	If RunEquipmentNeedToKnowReportIfClosed_ORIG = 1 then RunEquipmentNeedToKnowReportIfClosedMsgOrig = "On" Else RunEquipmentNeedToKnowReportIfClosedMsgOrig = "Off"
	
	If RunEquipmentNeedToKnowReportIfClosed <> RunEquipmentNeedToKnowReportIfClosed_ORIG Then
		CreateAuditLogEntry "Global Settings Change", "Global Settings Change", "Major", 1, "Equipment Need To Know Report generation schedule - Don''t Run If Closed (Monday-Friday Only) -  changed from " & RunEquipmentNeedToKnowReportIfClosedMsgOrig & " to " & RunEquipmentNeedToKnowReportIfClosedMsg
	End If


	If Request.Form("chkNoEquipmentNeedToKnowReportIfClosingEarly") = "on" then RunEquipmentNeedToKnowReportIfClosingEarlyMsg = "On" Else RunEquipmentNeedToKnowReportIfClosingEarlyMsg = "Off"
	If RunEquipmentNeedToKnowReportIfClosingEarly_ORIG = 1 then RunEquipmentNeedToKnowReportIfClosingEarlyMsgOrig = "On" Else RunEquipmentNeedToKnowReportIfClosingEarlyMsgOrig = "Off"
	
	If RunEquipmentNeedToKnowReportIfClosingEarly <> RunEquipmentNeedToKnowReportIfClosingEarly_ORIG Then
		CreateAuditLogEntry "Global Settings Change", "Global Settings Change", "Major", 1, "Equipment Need To Know Report generation schedule - Don''t Run If Closed (Monday-Friday Only) -  changed from " & RunEquipmentNeedToKnowReportIfClosingEarlyMsgOrig & " to " & RunEquipmentNeedToKnowReportIfClosingEarlyMsg
	End If


	'*********************************************************************
	'Build Array/String of Schedule Data From Request Form Field Data
	'*********************************************************************

	Schedule_EquipmentNeedToKnowReportGenerationUpdated = ""
	
	Schedule_EquipmentNeedToKnowReportGenerationUpdated = GenerateEquipmentNeedToKnowReportSunday
	Schedule_EquipmentNeedToKnowReportGenerationUpdated = Schedule_EquipmentNeedToKnowReportGenerationUpdated & "," & GenerateEquipmentNeedToKnowReportMonday
	Schedule_EquipmentNeedToKnowReportGenerationUpdated = Schedule_EquipmentNeedToKnowReportGenerationUpdated & "," & GenerateEquipmentNeedToKnowReportTuesday
	Schedule_EquipmentNeedToKnowReportGenerationUpdated = Schedule_EquipmentNeedToKnowReportGenerationUpdated & "," & GenerateEquipmentNeedToKnowReportWednesday
	Schedule_EquipmentNeedToKnowReportGenerationUpdated = Schedule_EquipmentNeedToKnowReportGenerationUpdated & "," & GenerateEquipmentNeedToKnowReportThursday
	Schedule_EquipmentNeedToKnowReportGenerationUpdated = Schedule_EquipmentNeedToKnowReportGenerationUpdated & "," & GenerateEquipmentNeedToKnowReportFriday
	Schedule_EquipmentNeedToKnowReportGenerationUpdated = Schedule_EquipmentNeedToKnowReportGenerationUpdated & "," & GenerateEquipmentNeedToKnowReportSaturday
	
	Schedule_EquipmentNeedToKnowReportGenerationUpdated = Schedule_EquipmentNeedToKnowReportGenerationUpdated & "," & GenerateEquipmentNeedToKnowReportSundayTime
	Schedule_EquipmentNeedToKnowReportGenerationUpdated = Schedule_EquipmentNeedToKnowReportGenerationUpdated & "," & GenerateEquipmentNeedToKnowReportMondayTime
	Schedule_EquipmentNeedToKnowReportGenerationUpdated = Schedule_EquipmentNeedToKnowReportGenerationUpdated & "," & GenerateEquipmentNeedToKnowReportTuesdayTime
	Schedule_EquipmentNeedToKnowReportGenerationUpdated = Schedule_EquipmentNeedToKnowReportGenerationUpdated & "," & GenerateEquipmentNeedToKnowReportWednesdayTime
	Schedule_EquipmentNeedToKnowReportGenerationUpdated = Schedule_EquipmentNeedToKnowReportGenerationUpdated & "," & GenerateEquipmentNeedToKnowReportThursdayTime
	Schedule_EquipmentNeedToKnowReportGenerationUpdated = Schedule_EquipmentNeedToKnowReportGenerationUpdated & "," & GenerateEquipmentNeedToKnowReportFridayTime
	Schedule_EquipmentNeedToKnowReportGenerationUpdated = Schedule_EquipmentNeedToKnowReportGenerationUpdated & "," & GenerateEquipmentNeedToKnowReportSaturdayTime

	
	Schedule_EquipmentNeedToKnowReportGenerationUpdated = Schedule_EquipmentNeedToKnowReportGenerationUpdated & "," & RunEquipmentNeedToKnowReportIfClosed
	Schedule_EquipmentNeedToKnowReportGenerationUpdated = Schedule_EquipmentNeedToKnowReportGenerationUpdated & "," & RunEquipmentNeedToKnowReportIfClosingEarly
	
	Response.Write("<br><br><br>Schedule_EquipmentNeedToKnowReportGenerationUpdated: " & Schedule_EquipmentNeedToKnowReportGenerationUpdated)

	'*********************************************************************
	'Update SQL with Array/String of Schedule Data
	'*********************************************************************
		
	SQL = "UPDATE Settings_NeedToKnow SET Schedule_EquipmentNeedToKnowReportGeneration = '" & cStr(Schedule_EquipmentNeedToKnowReportGenerationUpdated) & "' "
	
	Response.Write("<br><br><br>SQL: " & SQL)
	
	Set cnn8 = Server.CreateObject("ADODB.Connection")
	cnn8.open (Session("ClientCnnString"))
	Set rs = Server.CreateObject("ADODB.Recordset")
	rs.CursorLocation = 3 
	Set rs = cnn8.Execute(SQL)
	
	set rs = Nothing
	cnn8.close
	set cnn8 = Nothing


	 Response.Redirect("equipment.asp")
	
%><!--#include file="../../../../inc/footer-main.asp"-->