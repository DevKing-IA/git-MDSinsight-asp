<!--#include file="../../../../inc/header.asp"-->

<%

	'***********************************************************
	'Get Values Of All Form Fields Posted 
	'***********************************************************
	
	GenerateInventoryNeedToKnowReportSunday = Request.Form("chkNoInventoryNeedToKnowReportSunday")
	GenerateInventoryNeedToKnowReportMonday = Request.Form("chkNoInventoryNeedToKnowReportMonday")
	GenerateInventoryNeedToKnowReportTuesday = Request.Form("chkNoInventoryNeedToKnowReportTuesday")
	GenerateInventoryNeedToKnowReportWednesday = Request.Form("chkNoInventoryNeedToKnowReportWednesday")
	GenerateInventoryNeedToKnowReportThursday = Request.Form("chkNoInventoryNeedToKnowReportThursday")
	GenerateInventoryNeedToKnowReportFriday = Request.Form("chkNoInventoryNeedToKnowReportFriday")
	GenerateInventoryNeedToKnowReportSaturday = Request.Form("chkNoInventoryNeedToKnowReportSaturday")
	
	GenerateInventoryNeedToKnowReportSundayTime = Request.Form("txtInventoryNeedToKnowReportSchedulerSundayTime")
	GenerateInventoryNeedToKnowReportMondayTime = Request.Form("txtInventoryNeedToKnowReportSchedulerMondayTime")
	GenerateInventoryNeedToKnowReportTuesdayTime = Request.Form("txtInventoryNeedToKnowReportSchedulerTuesdayTime")
	GenerateInventoryNeedToKnowReportWednesdayTime = Request.Form("txtInventoryNeedToKnowReportSchedulerWednesdayTime")
	GenerateInventoryNeedToKnowReportThursdayTime = Request.Form("txtInventoryNeedToKnowReportSchedulerThursdayTime")
	GenerateInventoryNeedToKnowReportFridayTime = Request.Form("txtInventoryNeedToKnowReportSchedulerFridayTime")
	GenerateInventoryNeedToKnowReportSaturdayTime = Request.Form("txtInventoryNeedToKnowReportSchedulerSaturdayTime")
	
	RunInventoryNeedToKnowReportIfClosed = Request.Form("chkNoInventoryNeedToKnowReportIfClosed")
	RunInventoryNeedToKnowReportIfClosingEarly = Request.Form("chkNoInventoryNeedToKnowReportIfClosingEarly")


	If Request.Form("chkNoInventoryNeedToKnowReportSunday") = "on" Then
		GenerateInventoryNeedToKnowReportSunday = 0
		GenerateInventoryNeedToKnowReportSundayTime = ""
	Else 
		GenerateInventoryNeedToKnowReportSunday = 1
	End If

	If Request.Form("chkNoInventoryNeedToKnowReportMonday") = "on" Then
		GenerateInventoryNeedToKnowReportMonday = 0
		GenerateInventoryNeedToKnowReportMondayTime = ""
	Else 
		GenerateInventoryNeedToKnowReportMonday = 1
	End If

	If Request.Form("chkNoInventoryNeedToKnowReportTuesday") = "on" Then
		GenerateInventoryNeedToKnowReportTuesday = 0
		GenerateInventoryNeedToKnowReportTuesdayTime = ""
	Else 
		GenerateInventoryNeedToKnowReportTuesday = 1
	End If

	If Request.Form("chkNoInventoryNeedToKnowReportWednesday") = "on" Then
		GenerateInventoryNeedToKnowReportWednesday = 0
		GenerateInventoryNeedToKnowReportWednesdayTime = ""
	Else 
		GenerateInventoryNeedToKnowReportWednesday = 1
	End If

	If Request.Form("chkNoInventoryNeedToKnowReportThursday") = "on" Then
		GenerateInventoryNeedToKnowReportThursday = 0
		GenerateInventoryNeedToKnowReportThursdayTime = ""
	Else 
		GenerateInventoryNeedToKnowReportThursday = 1
	End If

	If Request.Form("chkNoInventoryNeedToKnowReportFriday") = "on" Then
		GenerateInventoryNeedToKnowReportFriday = 0
		GenerateInventoryNeedToKnowReportFridayTime = ""
	Else 
		GenerateInventoryNeedToKnowReportFriday = 1
	End If

	If Request.Form("chkNoInventoryNeedToKnowReportSaturday") = "on" Then
		GenerateInventoryNeedToKnowReportSaturday = 0
		GenerateInventoryNeedToKnowReportSaturdayTime = ""
	Else 
		GenerateInventoryNeedToKnowReportSaturday = 1
	End If

	If Request.Form("chkNoInventoryNeedToKnowReportIfClosed") = "on" Then RunInventoryNeedToKnowReportIfClosed = 0 Else RunInventoryNeedToKnowReportIfClosed = 1
	If Request.Form("chkNoInventoryNeedToKnowReportIfClosingEarly") = "on" Then RunInventoryNeedToKnowReportIfClosingEarly = 0 Else RunInventoryNeedToKnowReportIfClosingEarly = 1
	
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
	
		Schedule_InventoryNeedToKnowReportGeneration = rsPropsectingSettings("Schedule_InventoryNeedToKnowReportGeneration")
		
		Schedule_InventoryNeedToKnowReportGenerationSettings = Split(Schedule_InventoryNeedToKnowReportGeneration,",")

		GenerateInventoryNeedToKnowReportSunday_ORIG = cInt(Schedule_InventoryNeedToKnowReportGenerationSettings(0))
		GenerateInventoryNeedToKnowReportMonday_ORIG = cInt(Schedule_InventoryNeedToKnowReportGenerationSettings(1))
		GenerateInventoryNeedToKnowReportTuesday_ORIG = cInt(Schedule_InventoryNeedToKnowReportGenerationSettings(2))
		GenerateInventoryNeedToKnowReportWednesday_ORIG = cInt(Schedule_InventoryNeedToKnowReportGenerationSettings(3))
		GenerateInventoryNeedToKnowReportThursday_ORIG = cInt(Schedule_InventoryNeedToKnowReportGenerationSettings(4))
		GenerateInventoryNeedToKnowReportFriday_ORIG = cInt(Schedule_InventoryNeedToKnowReportGenerationSettings(5))
		GenerateInventoryNeedToKnowReportSaturday_ORIG = cInt(Schedule_InventoryNeedToKnowReportGenerationSettings(6))
		GenerateInventoryNeedToKnowReportSundayTime_ORIG = Schedule_InventoryNeedToKnowReportGenerationSettings(7)
		GenerateInventoryNeedToKnowReportMondayTime_ORIG = Schedule_InventoryNeedToKnowReportGenerationSettings(8)
		GenerateInventoryNeedToKnowReportTuesdayTime_ORIG = Schedule_InventoryNeedToKnowReportGenerationSettings(9)
		GenerateInventoryNeedToKnowReportWednesdayTime_ORIG = Schedule_InventoryNeedToKnowReportGenerationSettings(10)
		GenerateInventoryNeedToKnowReportThursdayTime_ORIG = Schedule_InventoryNeedToKnowReportGenerationSettings(11)
		GenerateInventoryNeedToKnowReportFridayTime_ORIG = Schedule_InventoryNeedToKnowReportGenerationSettings(12)
		GenerateInventoryNeedToKnowReportSaturdayTime_ORIG = Schedule_InventoryNeedToKnowReportGenerationSettings(13)
		RunInventoryNeedToKnowReportIfClosed_ORIG = cInt(Schedule_InventoryNeedToKnowReportGenerationSettings(14))
		RunInventoryNeedToKnowReportIfClosingEarly_ORIG = cInt(Schedule_InventoryNeedToKnowReportGenerationSettings(15))
	
	End If
	
	set rsPropsectingSettings = Nothing
	cnnPropsectingSettings.close
	set cnnPropsectingSettings = Nothing
	

	'***********************************************************
	'Perform Audit Trail Entry Inserts
	'***********************************************************

	
	If Request.Form("chkNoInventoryNeedToKnowReportSunday") = "on" then GenerateInventoryNeedToKnowReportSundayMsg = "On" Else GenerateInventoryNeedToKnowReportSundayMsg = "Off"
	If GenerateInventoryNeedToKnowReportSunday_ORIG = 1 then GenerateInventoryNeedToKnowReportSundayMsgOrig = "On" Else GenerateInventoryNeedToKnowReportSundayMsgOrig = "Off"
	
	If GenerateInventoryNeedToKnowReportSunday <> GenerateInventoryNeedToKnowReportSunday_ORIG Then
		CreateAuditLogEntry "Global Settings Change", "Global Settings Change", "Major", 1, "Inventory Need To Know Report generation schedule for SUNDAY changed from " & GenerateInventoryNeedToKnowReportSundayMsgOrig & " to " & GenerateInventoryNeedToKnowReportSundayMsg
	End If
	
	If GenerateInventoryNeedToKnowReportSundayTime <> GenerateInventoryNeedToKnowReportSundayTime_ORIG Then
		If GenerateInventoryNeedToKnowReportSunday_ORIG = 0 AND GenerateInventoryNeedToKnowReportSunday = 1 Then
			CreateAuditLogEntry "Global Settings Change", "Global Settings Change", "Major", 1, "Inventory Need To Know Report generation schedule for SUNDAY turned on and set to run at " & GenerateInventoryNeedToKnowReportSundayTime
		Else
			CreateAuditLogEntry "Global Settings Change", "Global Settings Change", "Major", 1, "Inventory Need To Know Report generation scheduled run time for SUNDAY changed from " & GenerateInventoryNeedToKnowReportSundayTime_ORIG & " to " & GenerateInventoryNeedToKnowReportSundayTime
		End If
	End If
	
	

	If Request.Form("chkNoInventoryNeedToKnowReportMonday") = "on" then GenerateInventoryNeedToKnowReportMondayMsg = "On" Else GenerateInventoryNeedToKnowReportMondayMsg = "Off"
	If GenerateInventoryNeedToKnowReportMonday_ORIG = 1 then GenerateInventoryNeedToKnowReportMondayMsgOrig = "On" Else GenerateInventoryNeedToKnowReportMondayMsgOrig = "Off"
	
	If GenerateInventoryNeedToKnowReportMonday <> GenerateInventoryNeedToKnowReportMonday_ORIG Then
		CreateAuditLogEntry "Global Settings Change", "Global Settings Change", "Major", 1, "Inventory Need To Know Report generation schedule for Monday changed from " & GenerateInventoryNeedToKnowReportMondayMsgOrig & " to " & GenerateInventoryNeedToKnowReportMondayMsg
	End If
	
	If GenerateInventoryNeedToKnowReportMondayTime <> GenerateInventoryNeedToKnowReportMondayTime_ORIG Then
		If GenerateInventoryNeedToKnowReportMonday_ORIG = 0 AND GenerateInventoryNeedToKnowReportMonday = 1 Then
			CreateAuditLogEntry "Global Settings Change", "Global Settings Change", "Major", 1, "Inventory Need To Know Report generation schedule for Monday turned on and set to run at " & GenerateInventoryNeedToKnowReportMondayTime
		Else
			CreateAuditLogEntry "Global Settings Change", "Global Settings Change", "Major", 1, "Inventory Need To Know Report generation scheduled run time for Monday changed from " & GenerateInventoryNeedToKnowReportMondayTime_ORIG & " to " & GenerateInventoryNeedToKnowReportMondayTime
		End If
	End If
	


	If Request.Form("chkNoInventoryNeedToKnowReportTuesday") = "on" then GenerateInventoryNeedToKnowReportTuesdayMsg = "On" Else GenerateInventoryNeedToKnowReportTuesdayMsg = "Off"
	If GenerateInventoryNeedToKnowReportTuesday_ORIG = 1 then GenerateInventoryNeedToKnowReportTuesdayMsgOrig = "On" Else GenerateInventoryNeedToKnowReportTuesdayMsgOrig = "Off"
	
	If GenerateInventoryNeedToKnowReportTuesday <> GenerateInventoryNeedToKnowReportTuesday_ORIG Then
		CreateAuditLogEntry "Global Settings Change", "Global Settings Change", "Major", 1, "Inventory Need To Know Report generation schedule for Tuesday changed from " & GenerateInventoryNeedToKnowReportTuesdayMsgOrig & " to " & GenerateInventoryNeedToKnowReportTuesdayMsg
	End If
	
	If GenerateInventoryNeedToKnowReportTuesdayTime <> GenerateInventoryNeedToKnowReportTuesdayTime_ORIG Then
		If GenerateInventoryNeedToKnowReportTuesday_ORIG = 0 AND GenerateInventoryNeedToKnowReportTuesday = 1 Then
			CreateAuditLogEntry "Global Settings Change", "Global Settings Change", "Major", 1, "Inventory Need To Know Report generation schedule for Tuesday turned on and set to run at " & GenerateInventoryNeedToKnowReportTuesdayTime
		Else
			CreateAuditLogEntry "Global Settings Change", "Global Settings Change", "Major", 1, "Inventory Need To Know Report generation scheduled run time for Tuesday changed from " & GenerateInventoryNeedToKnowReportTuesdayTime_ORIG & " to " & GenerateInventoryNeedToKnowReportTuesdayTime
		End If
	End If



	If Request.Form("chkNoInventoryNeedToKnowReportWednesday") = "on" then GenerateInventoryNeedToKnowReportWednesdayMsg = "On" Else GenerateInventoryNeedToKnowReportWednesdayMsg = "Off"
	If GenerateInventoryNeedToKnowReportWednesday_ORIG = 1 then GenerateInventoryNeedToKnowReportWednesdayMsgOrig = "On" Else GenerateInventoryNeedToKnowReportWednesdayMsgOrig = "Off"
	
	If GenerateInventoryNeedToKnowReportWednesday <> GenerateInventoryNeedToKnowReportWednesday_ORIG Then
		CreateAuditLogEntry "Global Settings Change", "Global Settings Change", "Major", 1, "Inventory Need To Know Report generation schedule for Wednesday changed from " & GenerateInventoryNeedToKnowReportWednesdayMsgOrig & " to " & GenerateInventoryNeedToKnowReportWednesdayMsg
	End If
	
	If GenerateInventoryNeedToKnowReportWednesdayTime <> GenerateInventoryNeedToKnowReportWednesdayTime_ORIG Then
		If GenerateInventoryNeedToKnowReportWednesday_ORIG = 0 AND GenerateInventoryNeedToKnowReportWednesday = 1 Then
			CreateAuditLogEntry "Global Settings Change", "Global Settings Change", "Major", 1, "Inventory Need To Know Report generation schedule for Wednesday turned on and set to run at " & GenerateInventoryNeedToKnowReportWednesdayTime
		Else
			CreateAuditLogEntry "Global Settings Change", "Global Settings Change", "Major", 1, "Inventory Need To Know Report generation scheduled run time for Wednesday changed from " & GenerateInventoryNeedToKnowReportWednesdayTime_ORIG & " to " & GenerateInventoryNeedToKnowReportWednesdayTime
		End If
	End If



	If Request.Form("chkNoInventoryNeedToKnowReportThursday") = "on" then GenerateInventoryNeedToKnowReportThursdayMsg = "On" Else GenerateInventoryNeedToKnowReportThursdayMsg = "Off"
	If GenerateInventoryNeedToKnowReportThursday_ORIG = 1 then GenerateInventoryNeedToKnowReportThursdayMsgOrig = "On" Else GenerateInventoryNeedToKnowReportThursdayMsgOrig = "Off"
	
	If GenerateInventoryNeedToKnowReportThursday <> GenerateInventoryNeedToKnowReportThursday_ORIG Then
		CreateAuditLogEntry "Global Settings Change", "Global Settings Change", "Major", 1, "Inventory Need To Know Report generation schedule for Thursday changed from " & GenerateInventoryNeedToKnowReportThursdayMsgOrig & " to " & GenerateInventoryNeedToKnowReportThursdayMsg
	End If
	
	If GenerateInventoryNeedToKnowReportThursdayTime <> GenerateInventoryNeedToKnowReportThursdayTime_ORIG Then
		If GenerateInventoryNeedToKnowReportThursday_ORIG = 0 AND GenerateInventoryNeedToKnowReportThursday = 1 Then
			CreateAuditLogEntry "Global Settings Change", "Global Settings Change", "Major", 1, "Inventory Need To Know Report generation schedule for Thursday turned on and set to run at " & GenerateInventoryNeedToKnowReportThursdayTime
		Else
			CreateAuditLogEntry "Global Settings Change", "Global Settings Change", "Major", 1, "Inventory Need To Know Report generation scheduled run time for Thursday changed from " & GenerateInventoryNeedToKnowReportThursdayTime_ORIG & " to " & GenerateInventoryNeedToKnowReportThursdayTime
		End If
	End If



	If Request.Form("chkNoInventoryNeedToKnowReportFriday") = "on" then GenerateInventoryNeedToKnowReportFridayMsg = "On" Else GenerateInventoryNeedToKnowReportFridayMsg = "Off"
	If GenerateInventoryNeedToKnowReportFriday_ORIG = 1 then GenerateInventoryNeedToKnowReportFridayMsgOrig = "On" Else GenerateInventoryNeedToKnowReportFridayMsgOrig = "Off"
	
	If GenerateInventoryNeedToKnowReportFriday <> GenerateInventoryNeedToKnowReportFriday_ORIG Then
		CreateAuditLogEntry "Global Settings Change", "Global Settings Change", "Major", 1, "Inventory Need To Know Report generation schedule for Friday changed from " & GenerateInventoryNeedToKnowReportFridayMsgOrig & " to " & GenerateInventoryNeedToKnowReportFridayMsg
	End If
	
	If GenerateInventoryNeedToKnowReportFridayTime <> GenerateInventoryNeedToKnowReportFridayTime_ORIG Then
		If GenerateInventoryNeedToKnowReportFriday_ORIG = 0 AND GenerateInventoryNeedToKnowReportFriday = 1 Then
			CreateAuditLogEntry "Global Settings Change", "Global Settings Change", "Major", 1, "Inventory Need To Know Report generation schedule for Friday turned on and set to run at " & GenerateInventoryNeedToKnowReportFridayTime
		Else
			CreateAuditLogEntry "Global Settings Change", "Global Settings Change", "Major", 1, "Inventory Need To Know Report generation scheduled run time for Friday changed from " & GenerateInventoryNeedToKnowReportFridayTime_ORIG & " to " & GenerateInventoryNeedToKnowReportFridayTime
		End If
	End If



	If Request.Form("chkNoInventoryNeedToKnowReportSaturday") = "on" then GenerateInventoryNeedToKnowReportSaturdayMsg = "On" Else GenerateInventoryNeedToKnowReportSaturdayMsg = "Off"
	If GenerateInventoryNeedToKnowReportSaturday_ORIG = 1 then GenerateInventoryNeedToKnowReportSaturdayMsgOrig = "On" Else GenerateInventoryNeedToKnowReportSaturdayMsgOrig = "Off"
	
	If GenerateInventoryNeedToKnowReportSaturday <> GenerateInventoryNeedToKnowReportSaturday_ORIG Then
		CreateAuditLogEntry "Global Settings Change", "Global Settings Change", "Major", 1, "Inventory Need To Know Report generation schedule for Saturday changed from " & GenerateInventoryNeedToKnowReportSaturdayMsgOrig & " to " & GenerateInventoryNeedToKnowReportSaturdayMsg
	End If
	
	If GenerateInventoryNeedToKnowReportSaturdayTime <> GenerateInventoryNeedToKnowReportSaturdayTime_ORIG Then
		If GenerateInventoryNeedToKnowReportSaturday_ORIG = 0 AND GenerateInventoryNeedToKnowReportSaturday = 1 Then
			CreateAuditLogEntry "Global Settings Change", "Global Settings Change", "Major", 1, "Inventory Need To Know Report generation schedule for Saturday turned on and set to run at " & GenerateInventoryNeedToKnowReportSaturdayTime
		Else
			CreateAuditLogEntry "Global Settings Change", "Global Settings Change", "Major", 1, "Inventory Need To Know Report generation scheduled run time for Saturday changed from " & GenerateInventoryNeedToKnowReportSaturdayTime_ORIG & " to " & GenerateInventoryNeedToKnowReportSaturdayTime
		End If
	End If


	If Request.Form("chkNoInventoryNeedToKnowReportIfClosed") = "on" then RunInventoryNeedToKnowReportIfClosedMsg = "On" Else RunInventoryNeedToKnowReportIfClosedMsg = "Off"
	If RunInventoryNeedToKnowReportIfClosed_ORIG = 1 then RunInventoryNeedToKnowReportIfClosedMsgOrig = "On" Else RunInventoryNeedToKnowReportIfClosedMsgOrig = "Off"
	
	If RunInventoryNeedToKnowReportIfClosed <> RunInventoryNeedToKnowReportIfClosed_ORIG Then
		CreateAuditLogEntry "Global Settings Change", "Global Settings Change", "Major", 1, "Inventory Need To Know Report generation schedule - Don''t Run If Closed (Monday-Friday Only) -  changed from " & RunInventoryNeedToKnowReportIfClosedMsgOrig & " to " & RunInventoryNeedToKnowReportIfClosedMsg
	End If


	If Request.Form("chkNoInventoryNeedToKnowReportIfClosingEarly") = "on" then RunInventoryNeedToKnowReportIfClosingEarlyMsg = "On" Else RunInventoryNeedToKnowReportIfClosingEarlyMsg = "Off"
	If RunInventoryNeedToKnowReportIfClosingEarly_ORIG = 1 then RunInventoryNeedToKnowReportIfClosingEarlyMsgOrig = "On" Else RunInventoryNeedToKnowReportIfClosingEarlyMsgOrig = "Off"
	
	If RunInventoryNeedToKnowReportIfClosingEarly <> RunInventoryNeedToKnowReportIfClosingEarly_ORIG Then
		CreateAuditLogEntry "Global Settings Change", "Global Settings Change", "Major", 1, "Inventory Need To Know Report generation schedule - Don''t Run If Closed (Monday-Friday Only) -  changed from " & RunInventoryNeedToKnowReportIfClosingEarlyMsgOrig & " to " & RunInventoryNeedToKnowReportIfClosingEarlyMsg
	End If


	'*********************************************************************
	'Build Array/String of Schedule Data From Request Form Field Data
	'*********************************************************************

	Schedule_InventoryNeedToKnowReportGenerationUpdated = ""
	
	Schedule_InventoryNeedToKnowReportGenerationUpdated = GenerateInventoryNeedToKnowReportSunday
	Schedule_InventoryNeedToKnowReportGenerationUpdated = Schedule_InventoryNeedToKnowReportGenerationUpdated & "," & GenerateInventoryNeedToKnowReportMonday
	Schedule_InventoryNeedToKnowReportGenerationUpdated = Schedule_InventoryNeedToKnowReportGenerationUpdated & "," & GenerateInventoryNeedToKnowReportTuesday
	Schedule_InventoryNeedToKnowReportGenerationUpdated = Schedule_InventoryNeedToKnowReportGenerationUpdated & "," & GenerateInventoryNeedToKnowReportWednesday
	Schedule_InventoryNeedToKnowReportGenerationUpdated = Schedule_InventoryNeedToKnowReportGenerationUpdated & "," & GenerateInventoryNeedToKnowReportThursday
	Schedule_InventoryNeedToKnowReportGenerationUpdated = Schedule_InventoryNeedToKnowReportGenerationUpdated & "," & GenerateInventoryNeedToKnowReportFriday
	Schedule_InventoryNeedToKnowReportGenerationUpdated = Schedule_InventoryNeedToKnowReportGenerationUpdated & "," & GenerateInventoryNeedToKnowReportSaturday
	
	Schedule_InventoryNeedToKnowReportGenerationUpdated = Schedule_InventoryNeedToKnowReportGenerationUpdated & "," & GenerateInventoryNeedToKnowReportSundayTime
	Schedule_InventoryNeedToKnowReportGenerationUpdated = Schedule_InventoryNeedToKnowReportGenerationUpdated & "," & GenerateInventoryNeedToKnowReportMondayTime
	Schedule_InventoryNeedToKnowReportGenerationUpdated = Schedule_InventoryNeedToKnowReportGenerationUpdated & "," & GenerateInventoryNeedToKnowReportTuesdayTime
	Schedule_InventoryNeedToKnowReportGenerationUpdated = Schedule_InventoryNeedToKnowReportGenerationUpdated & "," & GenerateInventoryNeedToKnowReportWednesdayTime
	Schedule_InventoryNeedToKnowReportGenerationUpdated = Schedule_InventoryNeedToKnowReportGenerationUpdated & "," & GenerateInventoryNeedToKnowReportThursdayTime
	Schedule_InventoryNeedToKnowReportGenerationUpdated = Schedule_InventoryNeedToKnowReportGenerationUpdated & "," & GenerateInventoryNeedToKnowReportFridayTime
	Schedule_InventoryNeedToKnowReportGenerationUpdated = Schedule_InventoryNeedToKnowReportGenerationUpdated & "," & GenerateInventoryNeedToKnowReportSaturdayTime

	
	Schedule_InventoryNeedToKnowReportGenerationUpdated = Schedule_InventoryNeedToKnowReportGenerationUpdated & "," & RunInventoryNeedToKnowReportIfClosed
	Schedule_InventoryNeedToKnowReportGenerationUpdated = Schedule_InventoryNeedToKnowReportGenerationUpdated & "," & RunInventoryNeedToKnowReportIfClosingEarly
	
	Response.Write("<br><br><br>Schedule_InventoryNeedToKnowReportGenerationUpdated: " & Schedule_InventoryNeedToKnowReportGenerationUpdated)

	'*********************************************************************
	'Update SQL with Array/String of Schedule Data
	'*********************************************************************
		
	SQL = "UPDATE Settings_NeedToKnow SET Schedule_InventoryNeedToKnowReportGeneration = '" & cStr(Schedule_InventoryNeedToKnowReportGenerationUpdated) & "' "
	
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
	
%><!--#include file="../../../../inc/footer-main.asp"-->