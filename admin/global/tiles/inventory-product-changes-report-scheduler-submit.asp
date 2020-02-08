<!--#include file="../../../inc/header.asp"-->

<%

	'***********************************************************
	'Get Values Of All Form Fields Posted
	'***********************************************************
	
	GenerateInventoryProductChangesReportSunday = Request.Form("chkNoInventoryProductChangesReportSunday")
	GenerateInventoryProductChangesReportMonday = Request.Form("chkNoInventoryProductChangesReportMonday")
	GenerateInventoryProductChangesReportTuesday = Request.Form("chkNoInventoryProductChangesReportTuesday")
	GenerateInventoryProductChangesReportWednesday = Request.Form("chkNoInventoryProductChangesReportWednesday")
	GenerateInventoryProductChangesReportThursday = Request.Form("chkNoInventoryProductChangesReportThursday")
	GenerateInventoryProductChangesReportFriday = Request.Form("chkNoInventoryProductChangesReportFriday")
	GenerateInventoryProductChangesReportSaturday = Request.Form("chkNoInventoryProductChangesReportSaturday")
	
	GenerateInventoryProductChangesReportSundayTime = Request.Form("txtInventoryProductChangesReportSchedulerSundayTime")
	GenerateInventoryProductChangesReportMondayTime = Request.Form("txtInventoryProductChangesReportSchedulerMondayTime")
	GenerateInventoryProductChangesReportTuesdayTime = Request.Form("txtInventoryProductChangesReportSchedulerTuesdayTime")
	GenerateInventoryProductChangesReportWednesdayTime = Request.Form("txtInventoryProductChangesReportSchedulerWednesdayTime")
	GenerateInventoryProductChangesReportThursdayTime = Request.Form("txtInventoryProductChangesReportSchedulerThursdayTime")
	GenerateInventoryProductChangesReportFridayTime = Request.Form("txtInventoryProductChangesReportSchedulerFridayTime")
	GenerateInventoryProductChangesReportSaturdayTime = Request.Form("txtInventoryProductChangesReportSchedulerSaturdayTime")
	
	RunInventoryProductChangesReportIfClosed = Request.Form("chkNoInventoryProductChangesReportIfClosed")
	RunInventoryProductChangesReportIfClosingEarly = Request.Form("chkNoInventoryProductChangesReportIfClosingEarly")


	If Request.Form("chkNoInventoryProductChangesReportSunday") = "on" Then
		GenerateInventoryProductChangesReportSunday = 0
		GenerateInventoryProductChangesReportSundayTime = ""
	Else 
		GenerateInventoryProductChangesReportSunday = 1
	End If

	If Request.Form("chkNoInventoryProductChangesReportMonday") = "on" Then
		GenerateInventoryProductChangesReportMonday = 0
		GenerateInventoryProductChangesReportMondayTime = ""
	Else 
		GenerateInventoryProductChangesReportMonday = 1
	End If

	If Request.Form("chkNoInventoryProductChangesReportTuesday") = "on" Then
		GenerateInventoryProductChangesReportTuesday = 0
		GenerateInventoryProductChangesReportTuesdayTime = ""
	Else 
		GenerateInventoryProductChangesReportTuesday = 1
	End If

	If Request.Form("chkNoInventoryProductChangesReportWednesday") = "on" Then
		GenerateInventoryProductChangesReportWednesday = 0
		GenerateInventoryProductChangesReportWednesdayTime = ""
	Else 
		GenerateInventoryProductChangesReportWednesday = 1
	End If

	If Request.Form("chkNoInventoryProductChangesReportThursday") = "on" Then
		GenerateInventoryProductChangesReportThursday = 0
		GenerateInventoryProductChangesReportThursdayTime = ""
	Else 
		GenerateInventoryProductChangesReportThursday = 1
	End If

	If Request.Form("chkNoInventoryProductChangesReportFriday") = "on" Then
		GenerateInventoryProductChangesReportFriday = 0
		GenerateInventoryProductChangesReportFridayTime = ""
	Else 
		GenerateInventoryProductChangesReportFriday = 1
	End If

	If Request.Form("chkNoInventoryProductChangesReportSaturday") = "on" Then
		GenerateInventoryProductChangesReportSaturday = 0
		GenerateInventoryProductChangesReportSaturdayTime = ""
	Else 
		GenerateInventoryProductChangesReportSaturday = 1
	End If

	If Request.Form("chkNoInventoryProductChangesReportIfClosed") = "on" Then RunInventoryProductChangesReportIfClosed = 0 Else RunInventoryProductChangesReportIfClosed = 1
	If Request.Form("chkNoInventoryProductChangesReportIfClosingEarly") = "on" Then RunInventoryProductChangesReportIfClosingEarly = 0 Else RunInventoryProductChangesReportIfClosingEarly = 1
	
	'***********************************************************
	'Get Original Values For Audit Trail Entries
	'***********************************************************
	
	SQLFieldServiceSettings = "SELECT * FROM Settings_InventoryControl"
	
	Set cnnFieldServiceSettings = Server.CreateObject("ADODB.Connection")
	cnnFieldServiceSettings.open (Session("ClientCnnString"))
	Set rsFieldServiceSettings = Server.CreateObject("ADODB.Recordset")
	rsFieldServiceSettings.CursorLocation = 3 
	Set rsFieldServiceSettings = cnnFieldServiceSettings.Execute(SQLFieldServiceSettings)
		
	If NOT rsFieldServiceSettings.EOF Then
	
		Schedule_InventoryProductChangesReportGeneration = rsFieldServiceSettings("Schedule_InventoryProductChangesReportGeneration")
		
		Schedule_InventoryProductChangesReportGenerationSettings = Split(Schedule_InventoryProductChangesReportGeneration,",")

		GenerateInventoryProductChangesReportSunday_ORIG = cInt(Schedule_InventoryProductChangesReportGenerationSettings(0))
		GenerateInventoryProductChangesReportMonday_ORIG = cInt(Schedule_InventoryProductChangesReportGenerationSettings(1))
		GenerateInventoryProductChangesReportTuesday_ORIG = cInt(Schedule_InventoryProductChangesReportGenerationSettings(2))
		GenerateInventoryProductChangesReportWednesday_ORIG = cInt(Schedule_InventoryProductChangesReportGenerationSettings(3))
		GenerateInventoryProductChangesReportThursday_ORIG = cInt(Schedule_InventoryProductChangesReportGenerationSettings(4))
		GenerateInventoryProductChangesReportFriday_ORIG = cInt(Schedule_InventoryProductChangesReportGenerationSettings(5))
		GenerateInventoryProductChangesReportSaturday_ORIG = cInt(Schedule_InventoryProductChangesReportGenerationSettings(6))
		GenerateInventoryProductChangesReportSundayTime_ORIG = Schedule_InventoryProductChangesReportGenerationSettings(7)
		GenerateInventoryProductChangesReportMondayTime_ORIG = Schedule_InventoryProductChangesReportGenerationSettings(8)
		GenerateInventoryProductChangesReportTuesdayTime_ORIG = Schedule_InventoryProductChangesReportGenerationSettings(9)
		GenerateInventoryProductChangesReportWednesdayTime_ORIG = Schedule_InventoryProductChangesReportGenerationSettings(10)
		GenerateInventoryProductChangesReportThursdayTime_ORIG = Schedule_InventoryProductChangesReportGenerationSettings(11)
		GenerateInventoryProductChangesReportFridayTime_ORIG = Schedule_InventoryProductChangesReportGenerationSettings(12)
		GenerateInventoryProductChangesReportSaturdayTime_ORIG = Schedule_InventoryProductChangesReportGenerationSettings(13)
		RunInventoryProductChangesReportIfClosed_ORIG = cInt(Schedule_InventoryProductChangesReportGenerationSettings(14))
		RunInventoryProductChangesReportIfClosingEarly_ORIG = cInt(Schedule_InventoryProductChangesReportGenerationSettings(15))
	
	End If
	
	set rsFieldServiceSettings = Nothing
	cnnFieldServiceSettings.close
	set cnnFieldServiceSettings = Nothing
	

	'***********************************************************
	'Perform Audit Trail Entry Inserts
	'***********************************************************

	
	If Request.Form("chkNoInventoryProductChangesReportSunday") = "on" then GenerateInventoryProductChangesReportSundayMsg = "On" Else GenerateInventoryProductChangesReportSundayMsg = "Off"
	If GenerateInventoryProductChangesReportSunday_ORIG = 1 then GenerateInventoryProductChangesReportSundayMsgOrig = "On" Else GenerateInventoryProductChangesReportSundayMsgOrig = "Off"
	
	If GenerateInventoryProductChangesReportSunday <> GenerateInventoryProductChangesReportSunday_ORIG Then
		CreateAuditLogEntry "Global Settings Change", "Global Settings Change", "Major", 1, "Daily Inventory API Activity By Partner Report generation schedule for SUNDAY changed from " & GenerateInventoryProductChangesReportSundayMsgOrig & " to " & GenerateInventoryProductChangesReportSundayMsg
	End If
	
	If GenerateInventoryProductChangesReportSundayTime <> GenerateInventoryProductChangesReportSundayTime_ORIG Then
		If GenerateInventoryProductChangesReportSunday_ORIG = 0 AND GenerateInventoryProductChangesReportSunday = 1 Then
			CreateAuditLogEntry "Global Settings Change", "Global Settings Change", "Major", 1, "Daily Inventory API Activity By Partner Report generation schedule for SUNDAY turned on and set to run at " & GenerateInventoryProductChangesReportSundayTime
		Else
			CreateAuditLogEntry "Global Settings Change", "Global Settings Change", "Major", 1, "Daily Inventory API Activity By Partner Report generation scheduled run time for SUNDAY changed from " & GenerateInventoryProductChangesReportSundayTime_ORIG & " to " & GenerateInventoryProductChangesReportSundayTime
		End If
	End If
	
	

	If Request.Form("chkNoInventoryProductChangesReportMonday") = "on" then GenerateInventoryProductChangesReportMondayMsg = "On" Else GenerateInventoryProductChangesReportMondayMsg = "Off"
	If GenerateInventoryProductChangesReportMonday_ORIG = 1 then GenerateInventoryProductChangesReportMondayMsgOrig = "On" Else GenerateInventoryProductChangesReportMondayMsgOrig = "Off"
	
	If GenerateInventoryProductChangesReportMonday <> GenerateInventoryProductChangesReportMonday_ORIG Then
		CreateAuditLogEntry "Global Settings Change", "Global Settings Change", "Major", 1, "Daily Inventory API Activity By Partner Report generation schedule for Monday changed from " & GenerateInventoryProductChangesReportMondayMsgOrig & " to " & GenerateInventoryProductChangesReportMondayMsg
	End If
	
	If GenerateInventoryProductChangesReportMondayTime <> GenerateInventoryProductChangesReportMondayTime_ORIG Then
		If GenerateInventoryProductChangesReportMonday_ORIG = 0 AND GenerateInventoryProductChangesReportMonday = 1 Then
			CreateAuditLogEntry "Global Settings Change", "Global Settings Change", "Major", 1, "Daily Inventory API Activity By Partner Report generation schedule for Monday turned on and set to run at " & GenerateInventoryProductChangesReportMondayTime
		Else
			CreateAuditLogEntry "Global Settings Change", "Global Settings Change", "Major", 1, "Daily Inventory API Activity By Partner Report generation scheduled run time for Monday changed from " & GenerateInventoryProductChangesReportMondayTime_ORIG & " to " & GenerateInventoryProductChangesReportMondayTime
		End If
	End If
	


	If Request.Form("chkNoInventoryProductChangesReportTuesday") = "on" then GenerateInventoryProductChangesReportTuesdayMsg = "On" Else GenerateInventoryProductChangesReportTuesdayMsg = "Off"
	If GenerateInventoryProductChangesReportTuesday_ORIG = 1 then GenerateInventoryProductChangesReportTuesdayMsgOrig = "On" Else GenerateInventoryProductChangesReportTuesdayMsgOrig = "Off"
	
	If GenerateInventoryProductChangesReportTuesday <> GenerateInventoryProductChangesReportTuesday_ORIG Then
		CreateAuditLogEntry "Global Settings Change", "Global Settings Change", "Major", 1, "Daily Inventory API Activity By Partner Report generation schedule for Tuesday changed from " & GenerateInventoryProductChangesReportTuesdayMsgOrig & " to " & GenerateInventoryProductChangesReportTuesdayMsg
	End If
	
	If GenerateInventoryProductChangesReportTuesdayTime <> GenerateInventoryProductChangesReportTuesdayTime_ORIG Then
		If GenerateInventoryProductChangesReportTuesday_ORIG = 0 AND GenerateInventoryProductChangesReportTuesday = 1 Then
			CreateAuditLogEntry "Global Settings Change", "Global Settings Change", "Major", 1, "Daily Inventory API Activity By Partner Report generation schedule for Tuesday turned on and set to run at " & GenerateInventoryProductChangesReportTuesdayTime
		Else
			CreateAuditLogEntry "Global Settings Change", "Global Settings Change", "Major", 1, "Daily Inventory API Activity By Partner Report generation scheduled run time for Tuesday changed from " & GenerateInventoryProductChangesReportTuesdayTime_ORIG & " to " & GenerateInventoryProductChangesReportTuesdayTime
		End If
	End If



	If Request.Form("chkNoInventoryProductChangesReportWednesday") = "on" then GenerateInventoryProductChangesReportWednesdayMsg = "On" Else GenerateInventoryProductChangesReportWednesdayMsg = "Off"
	If GenerateInventoryProductChangesReportWednesday_ORIG = 1 then GenerateInventoryProductChangesReportWednesdayMsgOrig = "On" Else GenerateInventoryProductChangesReportWednesdayMsgOrig = "Off"
	
	If GenerateInventoryProductChangesReportWednesday <> GenerateInventoryProductChangesReportWednesday_ORIG Then
		CreateAuditLogEntry "Global Settings Change", "Global Settings Change", "Major", 1, "Daily Inventory API Activity By Partner Report generation schedule for Wednesday changed from " & GenerateInventoryProductChangesReportWednesdayMsgOrig & " to " & GenerateInventoryProductChangesReportWednesdayMsg
	End If
	
	If GenerateInventoryProductChangesReportWednesdayTime <> GenerateInventoryProductChangesReportWednesdayTime_ORIG Then
		If GenerateInventoryProductChangesReportWednesday_ORIG = 0 AND GenerateInventoryProductChangesReportWednesday = 1 Then
			CreateAuditLogEntry "Global Settings Change", "Global Settings Change", "Major", 1, "Daily Inventory API Activity By Partner Report generation schedule for Wednesday turned on and set to run at " & GenerateInventoryProductChangesReportWednesdayTime
		Else
			CreateAuditLogEntry "Global Settings Change", "Global Settings Change", "Major", 1, "Daily Inventory API Activity By Partner Report generation scheduled run time for Wednesday changed from " & GenerateInventoryProductChangesReportWednesdayTime_ORIG & " to " & GenerateInventoryProductChangesReportWednesdayTime
		End If
	End If



	If Request.Form("chkNoInventoryProductChangesReportThursday") = "on" then GenerateInventoryProductChangesReportThursdayMsg = "On" Else GenerateInventoryProductChangesReportThursdayMsg = "Off"
	If GenerateInventoryProductChangesReportThursday_ORIG = 1 then GenerateInventoryProductChangesReportThursdayMsgOrig = "On" Else GenerateInventoryProductChangesReportThursdayMsgOrig = "Off"
	
	If GenerateInventoryProductChangesReportThursday <> GenerateInventoryProductChangesReportThursday_ORIG Then
		CreateAuditLogEntry "Global Settings Change", "Global Settings Change", "Major", 1, "Daily Inventory API Activity By Partner Report generation schedule for Thursday changed from " & GenerateInventoryProductChangesReportThursdayMsgOrig & " to " & GenerateInventoryProductChangesReportThursdayMsg
	End If
	
	If GenerateInventoryProductChangesReportThursdayTime <> GenerateInventoryProductChangesReportThursdayTime_ORIG Then
		If GenerateInventoryProductChangesReportThursday_ORIG = 0 AND GenerateInventoryProductChangesReportThursday = 1 Then
			CreateAuditLogEntry "Global Settings Change", "Global Settings Change", "Major", 1, "Daily Inventory API Activity By Partner Report generation schedule for Thursday turned on and set to run at " & GenerateInventoryProductChangesReportThursdayTime
		Else
			CreateAuditLogEntry "Global Settings Change", "Global Settings Change", "Major", 1, "Daily Inventory API Activity By Partner Report generation scheduled run time for Thursday changed from " & GenerateInventoryProductChangesReportThursdayTime_ORIG & " to " & GenerateInventoryProductChangesReportThursdayTime
		End If
	End If



	If Request.Form("chkNoInventoryProductChangesReportFriday") = "on" then GenerateInventoryProductChangesReportFridayMsg = "On" Else GenerateInventoryProductChangesReportFridayMsg = "Off"
	If GenerateInventoryProductChangesReportFriday_ORIG = 1 then GenerateInventoryProductChangesReportFridayMsgOrig = "On" Else GenerateInventoryProductChangesReportFridayMsgOrig = "Off"
	
	If GenerateInventoryProductChangesReportFriday <> GenerateInventoryProductChangesReportFriday_ORIG Then
		CreateAuditLogEntry "Global Settings Change", "Global Settings Change", "Major", 1, "Daily Inventory API Activity By Partner Report generation schedule for Friday changed from " & GenerateInventoryProductChangesReportFridayMsgOrig & " to " & GenerateInventoryProductChangesReportFridayMsg
	End If
	
	If GenerateInventoryProductChangesReportFridayTime <> GenerateInventoryProductChangesReportFridayTime_ORIG Then
		If GenerateInventoryProductChangesReportFriday_ORIG = 0 AND GenerateInventoryProductChangesReportFriday = 1 Then
			CreateAuditLogEntry "Global Settings Change", "Global Settings Change", "Major", 1, "Daily Inventory API Activity By Partner Report generation schedule for Friday turned on and set to run at " & GenerateInventoryProductChangesReportFridayTime
		Else
			CreateAuditLogEntry "Global Settings Change", "Global Settings Change", "Major", 1, "Daily Inventory API Activity By Partner Report generation scheduled run time for Friday changed from " & GenerateInventoryProductChangesReportFridayTime_ORIG & " to " & GenerateInventoryProductChangesReportFridayTime
		End If
	End If



	If Request.Form("chkNoInventoryProductChangesReportSaturday") = "on" then GenerateInventoryProductChangesReportSaturdayMsg = "On" Else GenerateInventoryProductChangesReportSaturdayMsg = "Off"
	If GenerateInventoryProductChangesReportSaturday_ORIG = 1 then GenerateInventoryProductChangesReportSaturdayMsgOrig = "On" Else GenerateInventoryProductChangesReportSaturdayMsgOrig = "Off"
	
	If GenerateInventoryProductChangesReportSaturday <> GenerateInventoryProductChangesReportSaturday_ORIG Then
		CreateAuditLogEntry "Global Settings Change", "Global Settings Change", "Major", 1, "Daily Inventory API Activity By Partner Report generation schedule for Saturday changed from " & GenerateInventoryProductChangesReportSaturdayMsgOrig & " to " & GenerateInventoryProductChangesReportSaturdayMsg
	End If
	
	If GenerateInventoryProductChangesReportSaturdayTime <> GenerateInventoryProductChangesReportSaturdayTime_ORIG Then
		If GenerateInventoryProductChangesReportSaturday_ORIG = 0 AND GenerateInventoryProductChangesReportSaturday = 1 Then
			CreateAuditLogEntry "Global Settings Change", "Global Settings Change", "Major", 1, "Daily Inventory API Activity By Partner Report generation schedule for Saturday turned on and set to run at " & GenerateInventoryProductChangesReportSaturdayTime
		Else
			CreateAuditLogEntry "Global Settings Change", "Global Settings Change", "Major", 1, "Daily Inventory API Activity By Partner Report generation scheduled run time for Saturday changed from " & GenerateInventoryProductChangesReportSaturdayTime_ORIG & " to " & GenerateInventoryProductChangesReportSaturdayTime
		End If
	End If


	If Request.Form("chkNoInventoryProductChangesReportIfClosed") = "on" then RunInventoryProductChangesReportIfClosedMsg = "On" Else RunInventoryProductChangesReportIfClosedMsg = "Off"
	If RunInventoryProductChangesReportIfClosed_ORIG = 1 then RunInventoryProductChangesReportIfClosedMsgOrig = "On" Else RunInventoryProductChangesReportIfClosedMsgOrig = "Off"
	
	If RunInventoryProductChangesReportIfClosed <> RunInventoryProductChangesReportIfClosed_ORIG Then
		CreateAuditLogEntry "Global Settings Change", "Global Settings Change", "Major", 1, "Daily Inventory API Activity By Partner Report generation schedule - Don''t Run If Closed (Monday-Friday Only) -  changed from " & RunInventoryProductChangesReportIfClosedMsgOrig & " to " & RunInventoryProductChangesReportIfClosedMsg
	End If


	If Request.Form("chkNoInventoryProductChangesReportIfClosingEarly") = "on" then RunInventoryProductChangesReportIfClosingEarlyMsg = "On" Else RunInventoryProductChangesReportIfClosingEarlyMsg = "Off"
	If RunInventoryProductChangesReportIfClosingEarly_ORIG = 1 then RunInventoryProductChangesReportIfClosingEarlyMsgOrig = "On" Else RunInventoryProductChangesReportIfClosingEarlyMsgOrig = "Off"
	
	If RunInventoryProductChangesReportIfClosingEarly <> RunInventoryProductChangesReportIfClosingEarly_ORIG Then
		CreateAuditLogEntry "Global Settings Change", "Global Settings Change", "Major", 1, "Daily Inventory API Activity By Partner Report generation schedule - Don''t Run If Closed (Monday-Friday Only) -  changed from " & RunInventoryProductChangesReportIfClosingEarlyMsgOrig & " to " & RunInventoryProductChangesReportIfClosingEarlyMsg
	End If


	'*********************************************************************
	'Build Array/String of Schedule Data From Request Form Field Data
	'*********************************************************************

	Schedule_InventoryProductChangesReportGenerationUpdated = ""
	
	Schedule_InventoryProductChangesReportGenerationUpdated = GenerateInventoryProductChangesReportSunday
	Schedule_InventoryProductChangesReportGenerationUpdated = Schedule_InventoryProductChangesReportGenerationUpdated & "," & GenerateInventoryProductChangesReportMonday
	Schedule_InventoryProductChangesReportGenerationUpdated = Schedule_InventoryProductChangesReportGenerationUpdated & "," & GenerateInventoryProductChangesReportTuesday
	Schedule_InventoryProductChangesReportGenerationUpdated = Schedule_InventoryProductChangesReportGenerationUpdated & "," & GenerateInventoryProductChangesReportWednesday
	Schedule_InventoryProductChangesReportGenerationUpdated = Schedule_InventoryProductChangesReportGenerationUpdated & "," & GenerateInventoryProductChangesReportThursday
	Schedule_InventoryProductChangesReportGenerationUpdated = Schedule_InventoryProductChangesReportGenerationUpdated & "," & GenerateInventoryProductChangesReportFriday
	Schedule_InventoryProductChangesReportGenerationUpdated = Schedule_InventoryProductChangesReportGenerationUpdated & "," & GenerateInventoryProductChangesReportSaturday
	
	Schedule_InventoryProductChangesReportGenerationUpdated = Schedule_InventoryProductChangesReportGenerationUpdated & "," & GenerateInventoryProductChangesReportSundayTime
	Schedule_InventoryProductChangesReportGenerationUpdated = Schedule_InventoryProductChangesReportGenerationUpdated & "," & GenerateInventoryProductChangesReportMondayTime
	Schedule_InventoryProductChangesReportGenerationUpdated = Schedule_InventoryProductChangesReportGenerationUpdated & "," & GenerateInventoryProductChangesReportTuesdayTime
	Schedule_InventoryProductChangesReportGenerationUpdated = Schedule_InventoryProductChangesReportGenerationUpdated & "," & GenerateInventoryProductChangesReportWednesdayTime
	Schedule_InventoryProductChangesReportGenerationUpdated = Schedule_InventoryProductChangesReportGenerationUpdated & "," & GenerateInventoryProductChangesReportThursdayTime
	Schedule_InventoryProductChangesReportGenerationUpdated = Schedule_InventoryProductChangesReportGenerationUpdated & "," & GenerateInventoryProductChangesReportFridayTime
	Schedule_InventoryProductChangesReportGenerationUpdated = Schedule_InventoryProductChangesReportGenerationUpdated & "," & GenerateInventoryProductChangesReportSaturdayTime

	
	Schedule_InventoryProductChangesReportGenerationUpdated = Schedule_InventoryProductChangesReportGenerationUpdated & "," & RunInventoryProductChangesReportIfClosed
	Schedule_InventoryProductChangesReportGenerationUpdated = Schedule_InventoryProductChangesReportGenerationUpdated & "," & RunInventoryProductChangesReportIfClosingEarly
	
	Response.Write("<br><br><br>Schedule_InventoryProductChangesReportGenerationUpdated: " & Schedule_InventoryProductChangesReportGenerationUpdated)

	'*********************************************************************
	'Update SQL with Array/String of Schedule Data
	'*********************************************************************
		
	SQL = "UPDATE Settings_InventoryControl SET Schedule_InventoryProductChangesReportGeneration = '" & cStr(Schedule_InventoryProductChangesReportGenerationUpdated) & "' "
	
	Response.Write("<br><br><br>SQL: " & SQL)
	
	Set cnn8 = Server.CreateObject("ADODB.Connection")
	cnn8.open (Session("ClientCnnString"))
	Set rs = Server.CreateObject("ADODB.Recordset")
	rs.CursorLocation = 3 
	Set rs = cnn8.Execute(SQL)
	
	set rs = Nothing
	cnn8.close
	set cnn8 = Nothing


	 Response.Redirect("inventory.asp")
	
%><!--#include file="../../../inc/footer-main.asp"-->