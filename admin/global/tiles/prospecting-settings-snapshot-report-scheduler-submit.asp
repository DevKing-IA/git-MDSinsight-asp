<!--#include file="../../../inc/header.asp"-->

<%

	'***********************************************************
	'Get Values Of All Form Fields Posted 
	'***********************************************************
	
	GeneratePropsectingSnapshotReportSunday = Request.Form("chkNoPropsectingSnapshotReportSunday")
	GeneratePropsectingSnapshotReportMonday = Request.Form("chkNoPropsectingSnapshotReportMonday")
	GeneratePropsectingSnapshotReportTuesday = Request.Form("chkNoPropsectingSnapshotReportTuesday")
	GeneratePropsectingSnapshotReportWednesday = Request.Form("chkNoPropsectingSnapshotReportWednesday")
	GeneratePropsectingSnapshotReportThursday = Request.Form("chkNoPropsectingSnapshotReportThursday")
	GeneratePropsectingSnapshotReportFriday = Request.Form("chkNoPropsectingSnapshotReportFriday")
	GeneratePropsectingSnapshotReportSaturday = Request.Form("chkNoPropsectingSnapshotReportSaturday")
	
	GeneratePropsectingSnapshotReportSundayTime = Request.Form("txtPropsectingSnapshotReportSchedulerSundayTime")
	GeneratePropsectingSnapshotReportMondayTime = Request.Form("txtPropsectingSnapshotReportSchedulerMondayTime")
	GeneratePropsectingSnapshotReportTuesdayTime = Request.Form("txtPropsectingSnapshotReportSchedulerTuesdayTime")
	GeneratePropsectingSnapshotReportWednesdayTime = Request.Form("txtPropsectingSnapshotReportSchedulerWednesdayTime")
	GeneratePropsectingSnapshotReportThursdayTime = Request.Form("txtPropsectingSnapshotReportSchedulerThursdayTime")
	GeneratePropsectingSnapshotReportFridayTime = Request.Form("txtPropsectingSnapshotReportSchedulerFridayTime")
	GeneratePropsectingSnapshotReportSaturdayTime = Request.Form("txtPropsectingSnapshotReportSchedulerSaturdayTime")
	
	RunPropsectingSnapshotReportIfClosed = Request.Form("chkNoPropsectingSnapshotReportIfClosed")
	RunPropsectingSnapshotReportIfClosingEarly = Request.Form("chkNoPropsectingSnapshotReportIfClosingEarly")


	If Request.Form("chkNoPropsectingSnapshotReportSunday") = "on" Then
		GeneratePropsectingSnapshotReportSunday = 0
		GeneratePropsectingSnapshotReportSundayTime = ""
	Else 
		GeneratePropsectingSnapshotReportSunday = 1
	End If

	If Request.Form("chkNoPropsectingSnapshotReportMonday") = "on" Then
		GeneratePropsectingSnapshotReportMonday = 0
		GeneratePropsectingSnapshotReportMondayTime = ""
	Else 
		GeneratePropsectingSnapshotReportMonday = 1
	End If

	If Request.Form("chkNoPropsectingSnapshotReportTuesday") = "on" Then
		GeneratePropsectingSnapshotReportTuesday = 0
		GeneratePropsectingSnapshotReportTuesdayTime = ""
	Else 
		GeneratePropsectingSnapshotReportTuesday = 1
	End If

	If Request.Form("chkNoPropsectingSnapshotReportWednesday") = "on" Then
		GeneratePropsectingSnapshotReportWednesday = 0
		GeneratePropsectingSnapshotReportWednesdayTime = ""
	Else 
		GeneratePropsectingSnapshotReportWednesday = 1
	End If

	If Request.Form("chkNoPropsectingSnapshotReportThursday") = "on" Then
		GeneratePropsectingSnapshotReportThursday = 0
		GeneratePropsectingSnapshotReportThursdayTime = ""
	Else 
		GeneratePropsectingSnapshotReportThursday = 1
	End If

	If Request.Form("chkNoPropsectingSnapshotReportFriday") = "on" Then
		GeneratePropsectingSnapshotReportFriday = 0
		GeneratePropsectingSnapshotReportFridayTime = ""
	Else 
		GeneratePropsectingSnapshotReportFriday = 1
	End If

	If Request.Form("chkNoPropsectingSnapshotReportSaturday") = "on" Then
		GeneratePropsectingSnapshotReportSaturday = 0
		GeneratePropsectingSnapshotReportSaturdayTime = ""
	Else 
		GeneratePropsectingSnapshotReportSaturday = 1
	End If

	If Request.Form("chkNoPropsectingSnapshotReportIfClosed") = "on" Then RunPropsectingSnapshotReportIfClosed = 0 Else RunPropsectingSnapshotReportIfClosed = 1
	If Request.Form("chkNoPropsectingSnapshotReportIfClosingEarly") = "on" Then RunPropsectingSnapshotReportIfClosingEarly = 0 Else RunPropsectingSnapshotReportIfClosingEarly = 1
	
	'***********************************************************
	'Get Original Values For Audit Trail Entries
	'***********************************************************
	
	SQLPropsectingSettings = "SELECT * FROM Settings_Prospecting"
	
	Set cnnPropsectingSettings = Server.CreateObject("ADODB.Connection")
	cnnPropsectingSettings.open (Session("ClientCnnString"))
	Set rsPropsectingSettings = Server.CreateObject("ADODB.Recordset")
	rsPropsectingSettings.CursorLocation = 3 
	Set rsPropsectingSettings = cnnPropsectingSettings.Execute(SQLPropsectingSettings)
		
	If NOT rsPropsectingSettings.EOF Then
	
		Schedule_PropsectingSnapshotReportGeneration = rsPropsectingSettings("Schedule_ProspectingSnapshotReportGeneration")
		
		Schedule_PropsectingSnapshotReportGenerationSettings = Split(Schedule_PropsectingSnapshotReportGeneration,",")

		GeneratePropsectingSnapshotReportSunday_ORIG = cInt(Schedule_PropsectingSnapshotReportGenerationSettings(0))
		GeneratePropsectingSnapshotReportMonday_ORIG = cInt(Schedule_PropsectingSnapshotReportGenerationSettings(1))
		GeneratePropsectingSnapshotReportTuesday_ORIG = cInt(Schedule_PropsectingSnapshotReportGenerationSettings(2))
		GeneratePropsectingSnapshotReportWednesday_ORIG = cInt(Schedule_PropsectingSnapshotReportGenerationSettings(3))
		GeneratePropsectingSnapshotReportThursday_ORIG = cInt(Schedule_PropsectingSnapshotReportGenerationSettings(4))
		GeneratePropsectingSnapshotReportFriday_ORIG = cInt(Schedule_PropsectingSnapshotReportGenerationSettings(5))
		GeneratePropsectingSnapshotReportSaturday_ORIG = cInt(Schedule_PropsectingSnapshotReportGenerationSettings(6))
		GeneratePropsectingSnapshotReportSundayTime_ORIG = Schedule_PropsectingSnapshotReportGenerationSettings(7)
		GeneratePropsectingSnapshotReportMondayTime_ORIG = Schedule_PropsectingSnapshotReportGenerationSettings(8)
		GeneratePropsectingSnapshotReportTuesdayTime_ORIG = Schedule_PropsectingSnapshotReportGenerationSettings(9)
		GeneratePropsectingSnapshotReportWednesdayTime_ORIG = Schedule_PropsectingSnapshotReportGenerationSettings(10)
		GeneratePropsectingSnapshotReportThursdayTime_ORIG = Schedule_PropsectingSnapshotReportGenerationSettings(11)
		GeneratePropsectingSnapshotReportFridayTime_ORIG = Schedule_PropsectingSnapshotReportGenerationSettings(12)
		GeneratePropsectingSnapshotReportSaturdayTime_ORIG = Schedule_PropsectingSnapshotReportGenerationSettings(13)
		RunPropsectingSnapshotReportIfClosed_ORIG = cInt(Schedule_PropsectingSnapshotReportGenerationSettings(14))
		RunPropsectingSnapshotReportIfClosingEarly_ORIG = cInt(Schedule_PropsectingSnapshotReportGenerationSettings(15))
	
	End If
	
	set rsPropsectingSettings = Nothing
	cnnPropsectingSettings.close
	set cnnPropsectingSettings = Nothing
	

	'***********************************************************
	'Perform Audit Trail Entry Inserts
	'***********************************************************

	
	If Request.Form("chkNoPropsectingSnapshotReportSunday") = "on" then GeneratePropsectingSnapshotReportSundayMsg = "On" Else GeneratePropsectingSnapshotReportSundayMsg = "Off"
	If GeneratePropsectingSnapshotReportSunday_ORIG = 1 then GeneratePropsectingSnapshotReportSundayMsgOrig = "On" Else GeneratePropsectingSnapshotReportSundayMsgOrig = "Off"
	
	If GeneratePropsectingSnapshotReportSunday <> GeneratePropsectingSnapshotReportSunday_ORIG Then
		CreateAuditLogEntry "Global Settings Change", "Global Settings Change", "Major", 1, "Propsecting Snapshot Report generation schedule for SUNDAY changed from " & GeneratePropsectingSnapshotReportSundayMsgOrig & " to " & GeneratePropsectingSnapshotReportSundayMsg
	End If
	
	If GeneratePropsectingSnapshotReportSundayTime <> GeneratePropsectingSnapshotReportSundayTime_ORIG Then
		If GeneratePropsectingSnapshotReportSunday_ORIG = 0 AND GeneratePropsectingSnapshotReportSunday = 1 Then
			CreateAuditLogEntry "Global Settings Change", "Global Settings Change", "Major", 1, "Propsecting Snapshot Report generation schedule for SUNDAY turned on and set to run at " & GeneratePropsectingSnapshotReportSundayTime
		Else
			CreateAuditLogEntry "Global Settings Change", "Global Settings Change", "Major", 1, "Propsecting Snapshot Report generation scheduled run time for SUNDAY changed from " & GeneratePropsectingSnapshotReportSundayTime_ORIG & " to " & GeneratePropsectingSnapshotReportSundayTime
		End If
	End If
	
	

	If Request.Form("chkNoPropsectingSnapshotReportMonday") = "on" then GeneratePropsectingSnapshotReportMondayMsg = "On" Else GeneratePropsectingSnapshotReportMondayMsg = "Off"
	If GeneratePropsectingSnapshotReportMonday_ORIG = 1 then GeneratePropsectingSnapshotReportMondayMsgOrig = "On" Else GeneratePropsectingSnapshotReportMondayMsgOrig = "Off"
	
	If GeneratePropsectingSnapshotReportMonday <> GeneratePropsectingSnapshotReportMonday_ORIG Then
		CreateAuditLogEntry "Global Settings Change", "Global Settings Change", "Major", 1, "Propsecting Snapshot Report generation schedule for Monday changed from " & GeneratePropsectingSnapshotReportMondayMsgOrig & " to " & GeneratePropsectingSnapshotReportMondayMsg
	End If
	
	If GeneratePropsectingSnapshotReportMondayTime <> GeneratePropsectingSnapshotReportMondayTime_ORIG Then
		If GeneratePropsectingSnapshotReportMonday_ORIG = 0 AND GeneratePropsectingSnapshotReportMonday = 1 Then
			CreateAuditLogEntry "Global Settings Change", "Global Settings Change", "Major", 1, "Propsecting Snapshot Report generation schedule for Monday turned on and set to run at " & GeneratePropsectingSnapshotReportMondayTime
		Else
			CreateAuditLogEntry "Global Settings Change", "Global Settings Change", "Major", 1, "Propsecting Snapshot Report generation scheduled run time for Monday changed from " & GeneratePropsectingSnapshotReportMondayTime_ORIG & " to " & GeneratePropsectingSnapshotReportMondayTime
		End If
	End If
	


	If Request.Form("chkNoPropsectingSnapshotReportTuesday") = "on" then GeneratePropsectingSnapshotReportTuesdayMsg = "On" Else GeneratePropsectingSnapshotReportTuesdayMsg = "Off"
	If GeneratePropsectingSnapshotReportTuesday_ORIG = 1 then GeneratePropsectingSnapshotReportTuesdayMsgOrig = "On" Else GeneratePropsectingSnapshotReportTuesdayMsgOrig = "Off"
	
	If GeneratePropsectingSnapshotReportTuesday <> GeneratePropsectingSnapshotReportTuesday_ORIG Then
		CreateAuditLogEntry "Global Settings Change", "Global Settings Change", "Major", 1, "Propsecting Snapshot Report generation schedule for Tuesday changed from " & GeneratePropsectingSnapshotReportTuesdayMsgOrig & " to " & GeneratePropsectingSnapshotReportTuesdayMsg
	End If
	
	If GeneratePropsectingSnapshotReportTuesdayTime <> GeneratePropsectingSnapshotReportTuesdayTime_ORIG Then
		If GeneratePropsectingSnapshotReportTuesday_ORIG = 0 AND GeneratePropsectingSnapshotReportTuesday = 1 Then
			CreateAuditLogEntry "Global Settings Change", "Global Settings Change", "Major", 1, "Propsecting Snapshot Report generation schedule for Tuesday turned on and set to run at " & GeneratePropsectingSnapshotReportTuesdayTime
		Else
			CreateAuditLogEntry "Global Settings Change", "Global Settings Change", "Major", 1, "Propsecting Snapshot Report generation scheduled run time for Tuesday changed from " & GeneratePropsectingSnapshotReportTuesdayTime_ORIG & " to " & GeneratePropsectingSnapshotReportTuesdayTime
		End If
	End If



	If Request.Form("chkNoPropsectingSnapshotReportWednesday") = "on" then GeneratePropsectingSnapshotReportWednesdayMsg = "On" Else GeneratePropsectingSnapshotReportWednesdayMsg = "Off"
	If GeneratePropsectingSnapshotReportWednesday_ORIG = 1 then GeneratePropsectingSnapshotReportWednesdayMsgOrig = "On" Else GeneratePropsectingSnapshotReportWednesdayMsgOrig = "Off"
	
	If GeneratePropsectingSnapshotReportWednesday <> GeneratePropsectingSnapshotReportWednesday_ORIG Then
		CreateAuditLogEntry "Global Settings Change", "Global Settings Change", "Major", 1, "Propsecting Snapshot Report generation schedule for Wednesday changed from " & GeneratePropsectingSnapshotReportWednesdayMsgOrig & " to " & GeneratePropsectingSnapshotReportWednesdayMsg
	End If
	
	If GeneratePropsectingSnapshotReportWednesdayTime <> GeneratePropsectingSnapshotReportWednesdayTime_ORIG Then
		If GeneratePropsectingSnapshotReportWednesday_ORIG = 0 AND GeneratePropsectingSnapshotReportWednesday = 1 Then
			CreateAuditLogEntry "Global Settings Change", "Global Settings Change", "Major", 1, "Propsecting Snapshot Report generation schedule for Wednesday turned on and set to run at " & GeneratePropsectingSnapshotReportWednesdayTime
		Else
			CreateAuditLogEntry "Global Settings Change", "Global Settings Change", "Major", 1, "Propsecting Snapshot Report generation scheduled run time for Wednesday changed from " & GeneratePropsectingSnapshotReportWednesdayTime_ORIG & " to " & GeneratePropsectingSnapshotReportWednesdayTime
		End If
	End If



	If Request.Form("chkNoPropsectingSnapshotReportThursday") = "on" then GeneratePropsectingSnapshotReportThursdayMsg = "On" Else GeneratePropsectingSnapshotReportThursdayMsg = "Off"
	If GeneratePropsectingSnapshotReportThursday_ORIG = 1 then GeneratePropsectingSnapshotReportThursdayMsgOrig = "On" Else GeneratePropsectingSnapshotReportThursdayMsgOrig = "Off"
	
	If GeneratePropsectingSnapshotReportThursday <> GeneratePropsectingSnapshotReportThursday_ORIG Then
		CreateAuditLogEntry "Global Settings Change", "Global Settings Change", "Major", 1, "Propsecting Snapshot Report generation schedule for Thursday changed from " & GeneratePropsectingSnapshotReportThursdayMsgOrig & " to " & GeneratePropsectingSnapshotReportThursdayMsg
	End If
	
	If GeneratePropsectingSnapshotReportThursdayTime <> GeneratePropsectingSnapshotReportThursdayTime_ORIG Then
		If GeneratePropsectingSnapshotReportThursday_ORIG = 0 AND GeneratePropsectingSnapshotReportThursday = 1 Then
			CreateAuditLogEntry "Global Settings Change", "Global Settings Change", "Major", 1, "Propsecting Snapshot Report generation schedule for Thursday turned on and set to run at " & GeneratePropsectingSnapshotReportThursdayTime
		Else
			CreateAuditLogEntry "Global Settings Change", "Global Settings Change", "Major", 1, "Propsecting Snapshot Report generation scheduled run time for Thursday changed from " & GeneratePropsectingSnapshotReportThursdayTime_ORIG & " to " & GeneratePropsectingSnapshotReportThursdayTime
		End If
	End If



	If Request.Form("chkNoPropsectingSnapshotReportFriday") = "on" then GeneratePropsectingSnapshotReportFridayMsg = "On" Else GeneratePropsectingSnapshotReportFridayMsg = "Off"
	If GeneratePropsectingSnapshotReportFriday_ORIG = 1 then GeneratePropsectingSnapshotReportFridayMsgOrig = "On" Else GeneratePropsectingSnapshotReportFridayMsgOrig = "Off"
	
	If GeneratePropsectingSnapshotReportFriday <> GeneratePropsectingSnapshotReportFriday_ORIG Then
		CreateAuditLogEntry "Global Settings Change", "Global Settings Change", "Major", 1, "Propsecting Snapshot Report generation schedule for Friday changed from " & GeneratePropsectingSnapshotReportFridayMsgOrig & " to " & GeneratePropsectingSnapshotReportFridayMsg
	End If
	
	If GeneratePropsectingSnapshotReportFridayTime <> GeneratePropsectingSnapshotReportFridayTime_ORIG Then
		If GeneratePropsectingSnapshotReportFriday_ORIG = 0 AND GeneratePropsectingSnapshotReportFriday = 1 Then
			CreateAuditLogEntry "Global Settings Change", "Global Settings Change", "Major", 1, "Propsecting Snapshot Report generation schedule for Friday turned on and set to run at " & GeneratePropsectingSnapshotReportFridayTime
		Else
			CreateAuditLogEntry "Global Settings Change", "Global Settings Change", "Major", 1, "Propsecting Snapshot Report generation scheduled run time for Friday changed from " & GeneratePropsectingSnapshotReportFridayTime_ORIG & " to " & GeneratePropsectingSnapshotReportFridayTime
		End If
	End If



	If Request.Form("chkNoPropsectingSnapshotReportSaturday") = "on" then GeneratePropsectingSnapshotReportSaturdayMsg = "On" Else GeneratePropsectingSnapshotReportSaturdayMsg = "Off"
	If GeneratePropsectingSnapshotReportSaturday_ORIG = 1 then GeneratePropsectingSnapshotReportSaturdayMsgOrig = "On" Else GeneratePropsectingSnapshotReportSaturdayMsgOrig = "Off"
	
	If GeneratePropsectingSnapshotReportSaturday <> GeneratePropsectingSnapshotReportSaturday_ORIG Then
		CreateAuditLogEntry "Global Settings Change", "Global Settings Change", "Major", 1, "Propsecting Snapshot Report generation schedule for Saturday changed from " & GeneratePropsectingSnapshotReportSaturdayMsgOrig & " to " & GeneratePropsectingSnapshotReportSaturdayMsg
	End If
	
	If GeneratePropsectingSnapshotReportSaturdayTime <> GeneratePropsectingSnapshotReportSaturdayTime_ORIG Then
		If GeneratePropsectingSnapshotReportSaturday_ORIG = 0 AND GeneratePropsectingSnapshotReportSaturday = 1 Then
			CreateAuditLogEntry "Global Settings Change", "Global Settings Change", "Major", 1, "Propsecting Snapshot Report generation schedule for Saturday turned on and set to run at " & GeneratePropsectingSnapshotReportSaturdayTime
		Else
			CreateAuditLogEntry "Global Settings Change", "Global Settings Change", "Major", 1, "Propsecting Snapshot Report generation scheduled run time for Saturday changed from " & GeneratePropsectingSnapshotReportSaturdayTime_ORIG & " to " & GeneratePropsectingSnapshotReportSaturdayTime
		End If
	End If


	If Request.Form("chkNoPropsectingSnapshotReportIfClosed") = "on" then RunPropsectingSnapshotReportIfClosedMsg = "On" Else RunPropsectingSnapshotReportIfClosedMsg = "Off"
	If RunPropsectingSnapshotReportIfClosed_ORIG = 1 then RunPropsectingSnapshotReportIfClosedMsgOrig = "On" Else RunPropsectingSnapshotReportIfClosedMsgOrig = "Off"
	
	If RunPropsectingSnapshotReportIfClosed <> RunPropsectingSnapshotReportIfClosed_ORIG Then
		CreateAuditLogEntry "Global Settings Change", "Global Settings Change", "Major", 1, "Propsecting Snapshot Report generation schedule - Don''t Run If Closed (Monday-Friday Only) -  changed from " & RunPropsectingSnapshotReportIfClosedMsgOrig & " to " & RunPropsectingSnapshotReportIfClosedMsg
	End If


	If Request.Form("chkNoPropsectingSnapshotReportIfClosingEarly") = "on" then RunPropsectingSnapshotReportIfClosingEarlyMsg = "On" Else RunPropsectingSnapshotReportIfClosingEarlyMsg = "Off"
	If RunPropsectingSnapshotReportIfClosingEarly_ORIG = 1 then RunPropsectingSnapshotReportIfClosingEarlyMsgOrig = "On" Else RunPropsectingSnapshotReportIfClosingEarlyMsgOrig = "Off"
	
	If RunPropsectingSnapshotReportIfClosingEarly <> RunPropsectingSnapshotReportIfClosingEarly_ORIG Then
		CreateAuditLogEntry "Global Settings Change", "Global Settings Change", "Major", 1, "Propsecting Snapshot Report generation schedule - Don''t Run If Closed (Monday-Friday Only) -  changed from " & RunPropsectingSnapshotReportIfClosingEarlyMsgOrig & " to " & RunPropsectingSnapshotReportIfClosingEarlyMsg
	End If


	'*********************************************************************
	'Build Array/String of Schedule Data From Request Form Field Data
	'*********************************************************************

	Schedule_PropsectingSnapshotReportGenerationUpdated = ""
	
	Schedule_PropsectingSnapshotReportGenerationUpdated = GeneratePropsectingSnapshotReportSunday
	Schedule_PropsectingSnapshotReportGenerationUpdated = Schedule_PropsectingSnapshotReportGenerationUpdated & "," & GeneratePropsectingSnapshotReportMonday
	Schedule_PropsectingSnapshotReportGenerationUpdated = Schedule_PropsectingSnapshotReportGenerationUpdated & "," & GeneratePropsectingSnapshotReportTuesday
	Schedule_PropsectingSnapshotReportGenerationUpdated = Schedule_PropsectingSnapshotReportGenerationUpdated & "," & GeneratePropsectingSnapshotReportWednesday
	Schedule_PropsectingSnapshotReportGenerationUpdated = Schedule_PropsectingSnapshotReportGenerationUpdated & "," & GeneratePropsectingSnapshotReportThursday
	Schedule_PropsectingSnapshotReportGenerationUpdated = Schedule_PropsectingSnapshotReportGenerationUpdated & "," & GeneratePropsectingSnapshotReportFriday
	Schedule_PropsectingSnapshotReportGenerationUpdated = Schedule_PropsectingSnapshotReportGenerationUpdated & "," & GeneratePropsectingSnapshotReportSaturday
	
	Schedule_PropsectingSnapshotReportGenerationUpdated = Schedule_PropsectingSnapshotReportGenerationUpdated & "," & GeneratePropsectingSnapshotReportSundayTime
	Schedule_PropsectingSnapshotReportGenerationUpdated = Schedule_PropsectingSnapshotReportGenerationUpdated & "," & GeneratePropsectingSnapshotReportMondayTime
	Schedule_PropsectingSnapshotReportGenerationUpdated = Schedule_PropsectingSnapshotReportGenerationUpdated & "," & GeneratePropsectingSnapshotReportTuesdayTime
	Schedule_PropsectingSnapshotReportGenerationUpdated = Schedule_PropsectingSnapshotReportGenerationUpdated & "," & GeneratePropsectingSnapshotReportWednesdayTime
	Schedule_PropsectingSnapshotReportGenerationUpdated = Schedule_PropsectingSnapshotReportGenerationUpdated & "," & GeneratePropsectingSnapshotReportThursdayTime
	Schedule_PropsectingSnapshotReportGenerationUpdated = Schedule_PropsectingSnapshotReportGenerationUpdated & "," & GeneratePropsectingSnapshotReportFridayTime
	Schedule_PropsectingSnapshotReportGenerationUpdated = Schedule_PropsectingSnapshotReportGenerationUpdated & "," & GeneratePropsectingSnapshotReportSaturdayTime

	
	Schedule_PropsectingSnapshotReportGenerationUpdated = Schedule_PropsectingSnapshotReportGenerationUpdated & "," & RunPropsectingSnapshotReportIfClosed
	Schedule_PropsectingSnapshotReportGenerationUpdated = Schedule_PropsectingSnapshotReportGenerationUpdated & "," & RunPropsectingSnapshotReportIfClosingEarly
	
	Response.Write("<br><br><br>Schedule_PropsectingSnapshotReportGenerationUpdated: " & Schedule_PropsectingSnapshotReportGenerationUpdated)

	'*********************************************************************
	'Update SQL with Array/String of Schedule Data
	'*********************************************************************
		
	SQL = "UPDATE Settings_Prospecting SET Schedule_ProspectingSnapshotReportGeneration = '" & cStr(Schedule_PropsectingSnapshotReportGenerationUpdated) & "' "
	
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