<!--#include file="../../../../inc/header.asp"-->

<%

	'***********************************************************
	'Get Values Of All Form Fields Posted 
	'***********************************************************
	
	GenerateOrderAPINeedToKnowReportSunday = Request.Form("chkNoOrderAPINeedToKnowReportSunday")
	GenerateOrderAPINeedToKnowReportMonday = Request.Form("chkNoOrderAPINeedToKnowReportMonday")
	GenerateOrderAPINeedToKnowReportTuesday = Request.Form("chkNoOrderAPINeedToKnowReportTuesday")
	GenerateOrderAPINeedToKnowReportWednesday = Request.Form("chkNoOrderAPINeedToKnowReportWednesday")
	GenerateOrderAPINeedToKnowReportThursday = Request.Form("chkNoOrderAPINeedToKnowReportThursday")
	GenerateOrderAPINeedToKnowReportFriday = Request.Form("chkNoOrderAPINeedToKnowReportFriday")
	GenerateOrderAPINeedToKnowReportSaturday = Request.Form("chkNoOrderAPINeedToKnowReportSaturday")
	
	GenerateOrderAPINeedToKnowReportSundayTime = Request.Form("txtOrderAPINeedToKnowReportSchedulerSundayTime")
	GenerateOrderAPINeedToKnowReportMondayTime = Request.Form("txtOrderAPINeedToKnowReportSchedulerMondayTime")
	GenerateOrderAPINeedToKnowReportTuesdayTime = Request.Form("txtOrderAPINeedToKnowReportSchedulerTuesdayTime")
	GenerateOrderAPINeedToKnowReportWednesdayTime = Request.Form("txtOrderAPINeedToKnowReportSchedulerWednesdayTime")
	GenerateOrderAPINeedToKnowReportThursdayTime = Request.Form("txtOrderAPINeedToKnowReportSchedulerThursdayTime")
	GenerateOrderAPINeedToKnowReportFridayTime = Request.Form("txtOrderAPINeedToKnowReportSchedulerFridayTime")
	GenerateOrderAPINeedToKnowReportSaturdayTime = Request.Form("txtOrderAPINeedToKnowReportSchedulerSaturdayTime")
	
	RunOrderAPINeedToKnowReportIfClosed = Request.Form("chkNoOrderAPINeedToKnowReportIfClosed")
	RunOrderAPINeedToKnowReportIfClosingEarly = Request.Form("chkNoOrderAPINeedToKnowReportIfClosingEarly")


	If Request.Form("chkNoOrderAPINeedToKnowReportSunday") = "on" Then
		GenerateOrderAPINeedToKnowReportSunday = 0
		GenerateOrderAPINeedToKnowReportSundayTime = ""
	Else 
		GenerateOrderAPINeedToKnowReportSunday = 1
	End If

	If Request.Form("chkNoOrderAPINeedToKnowReportMonday") = "on" Then
		GenerateOrderAPINeedToKnowReportMonday = 0
		GenerateOrderAPINeedToKnowReportMondayTime = ""
	Else 
		GenerateOrderAPINeedToKnowReportMonday = 1
	End If

	If Request.Form("chkNoOrderAPINeedToKnowReportTuesday") = "on" Then
		GenerateOrderAPINeedToKnowReportTuesday = 0
		GenerateOrderAPINeedToKnowReportTuesdayTime = ""
	Else 
		GenerateOrderAPINeedToKnowReportTuesday = 1
	End If

	If Request.Form("chkNoOrderAPINeedToKnowReportWednesday") = "on" Then
		GenerateOrderAPINeedToKnowReportWednesday = 0
		GenerateOrderAPINeedToKnowReportWednesdayTime = ""
	Else 
		GenerateOrderAPINeedToKnowReportWednesday = 1
	End If

	If Request.Form("chkNoOrderAPINeedToKnowReportThursday") = "on" Then
		GenerateOrderAPINeedToKnowReportThursday = 0
		GenerateOrderAPINeedToKnowReportThursdayTime = ""
	Else 
		GenerateOrderAPINeedToKnowReportThursday = 1
	End If

	If Request.Form("chkNoOrderAPINeedToKnowReportFriday") = "on" Then
		GenerateOrderAPINeedToKnowReportFriday = 0
		GenerateOrderAPINeedToKnowReportFridayTime = ""
	Else 
		GenerateOrderAPINeedToKnowReportFriday = 1
	End If

	If Request.Form("chkNoOrderAPINeedToKnowReportSaturday") = "on" Then
		GenerateOrderAPINeedToKnowReportSaturday = 0
		GenerateOrderAPINeedToKnowReportSaturdayTime = ""
	Else 
		GenerateOrderAPINeedToKnowReportSaturday = 1
	End If

	If Request.Form("chkNoOrderAPINeedToKnowReportIfClosed") = "on" Then RunOrderAPINeedToKnowReportIfClosed = 0 Else RunOrderAPINeedToKnowReportIfClosed = 1
	If Request.Form("chkNoOrderAPINeedToKnowReportIfClosingEarly") = "on" Then RunOrderAPINeedToKnowReportIfClosingEarly = 0 Else RunOrderAPINeedToKnowReportIfClosingEarly = 1
	
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
	
		Schedule_OrderAPINeedToKnowReportGeneration = rsPropsectingSettings("Schedule_APINeedToKnowReportGeneration")
		
		Schedule_OrderAPINeedToKnowReportGenerationSettings = Split(Schedule_OrderAPINeedToKnowReportGeneration,",")

		GenerateOrderAPINeedToKnowReportSunday_ORIG = cInt(Schedule_OrderAPINeedToKnowReportGenerationSettings(0))
		GenerateOrderAPINeedToKnowReportMonday_ORIG = cInt(Schedule_OrderAPINeedToKnowReportGenerationSettings(1))
		GenerateOrderAPINeedToKnowReportTuesday_ORIG = cInt(Schedule_OrderAPINeedToKnowReportGenerationSettings(2))
		GenerateOrderAPINeedToKnowReportWednesday_ORIG = cInt(Schedule_OrderAPINeedToKnowReportGenerationSettings(3))
		GenerateOrderAPINeedToKnowReportThursday_ORIG = cInt(Schedule_OrderAPINeedToKnowReportGenerationSettings(4))
		GenerateOrderAPINeedToKnowReportFriday_ORIG = cInt(Schedule_OrderAPINeedToKnowReportGenerationSettings(5))
		GenerateOrderAPINeedToKnowReportSaturday_ORIG = cInt(Schedule_OrderAPINeedToKnowReportGenerationSettings(6))
		GenerateOrderAPINeedToKnowReportSundayTime_ORIG = Schedule_OrderAPINeedToKnowReportGenerationSettings(7)
		GenerateOrderAPINeedToKnowReportMondayTime_ORIG = Schedule_OrderAPINeedToKnowReportGenerationSettings(8)
		GenerateOrderAPINeedToKnowReportTuesdayTime_ORIG = Schedule_OrderAPINeedToKnowReportGenerationSettings(9)
		GenerateOrderAPINeedToKnowReportWednesdayTime_ORIG = Schedule_OrderAPINeedToKnowReportGenerationSettings(10)
		GenerateOrderAPINeedToKnowReportThursdayTime_ORIG = Schedule_OrderAPINeedToKnowReportGenerationSettings(11)
		GenerateOrderAPINeedToKnowReportFridayTime_ORIG = Schedule_OrderAPINeedToKnowReportGenerationSettings(12)
		GenerateOrderAPINeedToKnowReportSaturdayTime_ORIG = Schedule_OrderAPINeedToKnowReportGenerationSettings(13)
		RunOrderAPINeedToKnowReportIfClosed_ORIG = cInt(Schedule_OrderAPINeedToKnowReportGenerationSettings(14))
		RunOrderAPINeedToKnowReportIfClosingEarly_ORIG = cInt(Schedule_OrderAPINeedToKnowReportGenerationSettings(15))
	
	End If
	
	set rsPropsectingSettings = Nothing
	cnnPropsectingSettings.close
	set cnnPropsectingSettings = Nothing
	

	'***********************************************************
	'Perform Audit Trail Entry Inserts
	'***********************************************************

	
	If Request.Form("chkNoOrderAPINeedToKnowReportSunday") = "on" then GenerateOrderAPINeedToKnowReportSundayMsg = "On" Else GenerateOrderAPINeedToKnowReportSundayMsg = "Off"
	If GenerateOrderAPINeedToKnowReportSunday_ORIG = 1 then GenerateOrderAPINeedToKnowReportSundayMsgOrig = "On" Else GenerateOrderAPINeedToKnowReportSundayMsgOrig = "Off"
	
	If GenerateOrderAPINeedToKnowReportSunday <> GenerateOrderAPINeedToKnowReportSunday_ORIG Then
		CreateAuditLogEntry "Global Settings Change", "Global Settings Change", "Major", 1, "Order API Need To Know Report generation schedule for SUNDAY changed from " & GenerateOrderAPINeedToKnowReportSundayMsgOrig & " to " & GenerateOrderAPINeedToKnowReportSundayMsg
	End If
	
	If GenerateOrderAPINeedToKnowReportSundayTime <> GenerateOrderAPINeedToKnowReportSundayTime_ORIG Then
		If GenerateOrderAPINeedToKnowReportSunday_ORIG = 0 AND GenerateOrderAPINeedToKnowReportSunday = 1 Then
			CreateAuditLogEntry "Global Settings Change", "Global Settings Change", "Major", 1, "Order API Need To Know Report generation schedule for SUNDAY turned on and set to run at " & GenerateOrderAPINeedToKnowReportSundayTime
		Else
			CreateAuditLogEntry "Global Settings Change", "Global Settings Change", "Major", 1, "Order API Need To Know Report generation scheduled run time for SUNDAY changed from " & GenerateOrderAPINeedToKnowReportSundayTime_ORIG & " to " & GenerateOrderAPINeedToKnowReportSundayTime
		End If
	End If
	
	

	If Request.Form("chkNoOrderAPINeedToKnowReportMonday") = "on" then GenerateOrderAPINeedToKnowReportMondayMsg = "On" Else GenerateOrderAPINeedToKnowReportMondayMsg = "Off"
	If GenerateOrderAPINeedToKnowReportMonday_ORIG = 1 then GenerateOrderAPINeedToKnowReportMondayMsgOrig = "On" Else GenerateOrderAPINeedToKnowReportMondayMsgOrig = "Off"
	
	If GenerateOrderAPINeedToKnowReportMonday <> GenerateOrderAPINeedToKnowReportMonday_ORIG Then
		CreateAuditLogEntry "Global Settings Change", "Global Settings Change", "Major", 1, "Order API Need To Know Report generation schedule for Monday changed from " & GenerateOrderAPINeedToKnowReportMondayMsgOrig & " to " & GenerateOrderAPINeedToKnowReportMondayMsg
	End If
	
	If GenerateOrderAPINeedToKnowReportMondayTime <> GenerateOrderAPINeedToKnowReportMondayTime_ORIG Then
		If GenerateOrderAPINeedToKnowReportMonday_ORIG = 0 AND GenerateOrderAPINeedToKnowReportMonday = 1 Then
			CreateAuditLogEntry "Global Settings Change", "Global Settings Change", "Major", 1, "Order API Need To Know Report generation schedule for Monday turned on and set to run at " & GenerateOrderAPINeedToKnowReportMondayTime
		Else
			CreateAuditLogEntry "Global Settings Change", "Global Settings Change", "Major", 1, "Order API Need To Know Report generation scheduled run time for Monday changed from " & GenerateOrderAPINeedToKnowReportMondayTime_ORIG & " to " & GenerateOrderAPINeedToKnowReportMondayTime
		End If
	End If
	


	If Request.Form("chkNoOrderAPINeedToKnowReportTuesday") = "on" then GenerateOrderAPINeedToKnowReportTuesdayMsg = "On" Else GenerateOrderAPINeedToKnowReportTuesdayMsg = "Off"
	If GenerateOrderAPINeedToKnowReportTuesday_ORIG = 1 then GenerateOrderAPINeedToKnowReportTuesdayMsgOrig = "On" Else GenerateOrderAPINeedToKnowReportTuesdayMsgOrig = "Off"
	
	If GenerateOrderAPINeedToKnowReportTuesday <> GenerateOrderAPINeedToKnowReportTuesday_ORIG Then
		CreateAuditLogEntry "Global Settings Change", "Global Settings Change", "Major", 1, "Order API Need To Know Report generation schedule for Tuesday changed from " & GenerateOrderAPINeedToKnowReportTuesdayMsgOrig & " to " & GenerateOrderAPINeedToKnowReportTuesdayMsg
	End If
	
	If GenerateOrderAPINeedToKnowReportTuesdayTime <> GenerateOrderAPINeedToKnowReportTuesdayTime_ORIG Then
		If GenerateOrderAPINeedToKnowReportTuesday_ORIG = 0 AND GenerateOrderAPINeedToKnowReportTuesday = 1 Then
			CreateAuditLogEntry "Global Settings Change", "Global Settings Change", "Major", 1, "Order API Need To Know Report generation schedule for Tuesday turned on and set to run at " & GenerateOrderAPINeedToKnowReportTuesdayTime
		Else
			CreateAuditLogEntry "Global Settings Change", "Global Settings Change", "Major", 1, "Order API Need To Know Report generation scheduled run time for Tuesday changed from " & GenerateOrderAPINeedToKnowReportTuesdayTime_ORIG & " to " & GenerateOrderAPINeedToKnowReportTuesdayTime
		End If
	End If



	If Request.Form("chkNoOrderAPINeedToKnowReportWednesday") = "on" then GenerateOrderAPINeedToKnowReportWednesdayMsg = "On" Else GenerateOrderAPINeedToKnowReportWednesdayMsg = "Off"
	If GenerateOrderAPINeedToKnowReportWednesday_ORIG = 1 then GenerateOrderAPINeedToKnowReportWednesdayMsgOrig = "On" Else GenerateOrderAPINeedToKnowReportWednesdayMsgOrig = "Off"
	
	If GenerateOrderAPINeedToKnowReportWednesday <> GenerateOrderAPINeedToKnowReportWednesday_ORIG Then
		CreateAuditLogEntry "Global Settings Change", "Global Settings Change", "Major", 1, "Order API Need To Know Report generation schedule for Wednesday changed from " & GenerateOrderAPINeedToKnowReportWednesdayMsgOrig & " to " & GenerateOrderAPINeedToKnowReportWednesdayMsg
	End If
	
	If GenerateOrderAPINeedToKnowReportWednesdayTime <> GenerateOrderAPINeedToKnowReportWednesdayTime_ORIG Then
		If GenerateOrderAPINeedToKnowReportWednesday_ORIG = 0 AND GenerateOrderAPINeedToKnowReportWednesday = 1 Then
			CreateAuditLogEntry "Global Settings Change", "Global Settings Change", "Major", 1, "Order API Need To Know Report generation schedule for Wednesday turned on and set to run at " & GenerateOrderAPINeedToKnowReportWednesdayTime
		Else
			CreateAuditLogEntry "Global Settings Change", "Global Settings Change", "Major", 1, "Order API Need To Know Report generation scheduled run time for Wednesday changed from " & GenerateOrderAPINeedToKnowReportWednesdayTime_ORIG & " to " & GenerateOrderAPINeedToKnowReportWednesdayTime
		End If
	End If



	If Request.Form("chkNoOrderAPINeedToKnowReportThursday") = "on" then GenerateOrderAPINeedToKnowReportThursdayMsg = "On" Else GenerateOrderAPINeedToKnowReportThursdayMsg = "Off"
	If GenerateOrderAPINeedToKnowReportThursday_ORIG = 1 then GenerateOrderAPINeedToKnowReportThursdayMsgOrig = "On" Else GenerateOrderAPINeedToKnowReportThursdayMsgOrig = "Off"
	
	If GenerateOrderAPINeedToKnowReportThursday <> GenerateOrderAPINeedToKnowReportThursday_ORIG Then
		CreateAuditLogEntry "Global Settings Change", "Global Settings Change", "Major", 1, "Order API Need To Know Report generation schedule for Thursday changed from " & GenerateOrderAPINeedToKnowReportThursdayMsgOrig & " to " & GenerateOrderAPINeedToKnowReportThursdayMsg
	End If
	
	If GenerateOrderAPINeedToKnowReportThursdayTime <> GenerateOrderAPINeedToKnowReportThursdayTime_ORIG Then
		If GenerateOrderAPINeedToKnowReportThursday_ORIG = 0 AND GenerateOrderAPINeedToKnowReportThursday = 1 Then
			CreateAuditLogEntry "Global Settings Change", "Global Settings Change", "Major", 1, "Order API Need To Know Report generation schedule for Thursday turned on and set to run at " & GenerateOrderAPINeedToKnowReportThursdayTime
		Else
			CreateAuditLogEntry "Global Settings Change", "Global Settings Change", "Major", 1, "Order API Need To Know Report generation scheduled run time for Thursday changed from " & GenerateOrderAPINeedToKnowReportThursdayTime_ORIG & " to " & GenerateOrderAPINeedToKnowReportThursdayTime
		End If
	End If



	If Request.Form("chkNoOrderAPINeedToKnowReportFriday") = "on" then GenerateOrderAPINeedToKnowReportFridayMsg = "On" Else GenerateOrderAPINeedToKnowReportFridayMsg = "Off"
	If GenerateOrderAPINeedToKnowReportFriday_ORIG = 1 then GenerateOrderAPINeedToKnowReportFridayMsgOrig = "On" Else GenerateOrderAPINeedToKnowReportFridayMsgOrig = "Off"
	
	If GenerateOrderAPINeedToKnowReportFriday <> GenerateOrderAPINeedToKnowReportFriday_ORIG Then
		CreateAuditLogEntry "Global Settings Change", "Global Settings Change", "Major", 1, "Order API Need To Know Report generation schedule for Friday changed from " & GenerateOrderAPINeedToKnowReportFridayMsgOrig & " to " & GenerateOrderAPINeedToKnowReportFridayMsg
	End If
	
	If GenerateOrderAPINeedToKnowReportFridayTime <> GenerateOrderAPINeedToKnowReportFridayTime_ORIG Then
		If GenerateOrderAPINeedToKnowReportFriday_ORIG = 0 AND GenerateOrderAPINeedToKnowReportFriday = 1 Then
			CreateAuditLogEntry "Global Settings Change", "Global Settings Change", "Major", 1, "Order API Need To Know Report generation schedule for Friday turned on and set to run at " & GenerateOrderAPINeedToKnowReportFridayTime
		Else
			CreateAuditLogEntry "Global Settings Change", "Global Settings Change", "Major", 1, "Order API Need To Know Report generation scheduled run time for Friday changed from " & GenerateOrderAPINeedToKnowReportFridayTime_ORIG & " to " & GenerateOrderAPINeedToKnowReportFridayTime
		End If
	End If



	If Request.Form("chkNoOrderAPINeedToKnowReportSaturday") = "on" then GenerateOrderAPINeedToKnowReportSaturdayMsg = "On" Else GenerateOrderAPINeedToKnowReportSaturdayMsg = "Off"
	If GenerateOrderAPINeedToKnowReportSaturday_ORIG = 1 then GenerateOrderAPINeedToKnowReportSaturdayMsgOrig = "On" Else GenerateOrderAPINeedToKnowReportSaturdayMsgOrig = "Off"
	
	If GenerateOrderAPINeedToKnowReportSaturday <> GenerateOrderAPINeedToKnowReportSaturday_ORIG Then
		CreateAuditLogEntry "Global Settings Change", "Global Settings Change", "Major", 1, "Order API Need To Know Report generation schedule for Saturday changed from " & GenerateOrderAPINeedToKnowReportSaturdayMsgOrig & " to " & GenerateOrderAPINeedToKnowReportSaturdayMsg
	End If
	
	If GenerateOrderAPINeedToKnowReportSaturdayTime <> GenerateOrderAPINeedToKnowReportSaturdayTime_ORIG Then
		If GenerateOrderAPINeedToKnowReportSaturday_ORIG = 0 AND GenerateOrderAPINeedToKnowReportSaturday = 1 Then
			CreateAuditLogEntry "Global Settings Change", "Global Settings Change", "Major", 1, "Order API Need To Know Report generation schedule for Saturday turned on and set to run at " & GenerateOrderAPINeedToKnowReportSaturdayTime
		Else
			CreateAuditLogEntry "Global Settings Change", "Global Settings Change", "Major", 1, "Order API Need To Know Report generation scheduled run time for Saturday changed from " & GenerateOrderAPINeedToKnowReportSaturdayTime_ORIG & " to " & GenerateOrderAPINeedToKnowReportSaturdayTime
		End If
	End If


	If Request.Form("chkNoOrderAPINeedToKnowReportIfClosed") = "on" then RunOrderAPINeedToKnowReportIfClosedMsg = "On" Else RunOrderAPINeedToKnowReportIfClosedMsg = "Off"
	If RunOrderAPINeedToKnowReportIfClosed_ORIG = 1 then RunOrderAPINeedToKnowReportIfClosedMsgOrig = "On" Else RunOrderAPINeedToKnowReportIfClosedMsgOrig = "Off"
	
	If RunOrderAPINeedToKnowReportIfClosed <> RunOrderAPINeedToKnowReportIfClosed_ORIG Then
		CreateAuditLogEntry "Global Settings Change", "Global Settings Change", "Major", 1, "Order API Need To Know Report generation schedule - Don''t Run If Closed (Monday-Friday Only) -  changed from " & RunOrderAPINeedToKnowReportIfClosedMsgOrig & " to " & RunOrderAPINeedToKnowReportIfClosedMsg
	End If


	If Request.Form("chkNoOrderAPINeedToKnowReportIfClosingEarly") = "on" then RunOrderAPINeedToKnowReportIfClosingEarlyMsg = "On" Else RunOrderAPINeedToKnowReportIfClosingEarlyMsg = "Off"
	If RunOrderAPINeedToKnowReportIfClosingEarly_ORIG = 1 then RunOrderAPINeedToKnowReportIfClosingEarlyMsgOrig = "On" Else RunOrderAPINeedToKnowReportIfClosingEarlyMsgOrig = "Off"
	
	If RunOrderAPINeedToKnowReportIfClosingEarly <> RunOrderAPINeedToKnowReportIfClosingEarly_ORIG Then
		CreateAuditLogEntry "Global Settings Change", "Global Settings Change", "Major", 1, "Order API Need To Know Report generation schedule - Don''t Run If Closed (Monday-Friday Only) -  changed from " & RunOrderAPINeedToKnowReportIfClosingEarlyMsgOrig & " to " & RunOrderAPINeedToKnowReportIfClosingEarlyMsg
	End If


	'*********************************************************************
	'Build Array/String of Schedule Data From Request Form Field Data
	'*********************************************************************

	Schedule_OrderAPINeedToKnowReportGenerationUpdated = ""
	
	Schedule_OrderAPINeedToKnowReportGenerationUpdated = GenerateOrderAPINeedToKnowReportSunday
	Schedule_OrderAPINeedToKnowReportGenerationUpdated = Schedule_OrderAPINeedToKnowReportGenerationUpdated & "," & GenerateOrderAPINeedToKnowReportMonday
	Schedule_OrderAPINeedToKnowReportGenerationUpdated = Schedule_OrderAPINeedToKnowReportGenerationUpdated & "," & GenerateOrderAPINeedToKnowReportTuesday
	Schedule_OrderAPINeedToKnowReportGenerationUpdated = Schedule_OrderAPINeedToKnowReportGenerationUpdated & "," & GenerateOrderAPINeedToKnowReportWednesday
	Schedule_OrderAPINeedToKnowReportGenerationUpdated = Schedule_OrderAPINeedToKnowReportGenerationUpdated & "," & GenerateOrderAPINeedToKnowReportThursday
	Schedule_OrderAPINeedToKnowReportGenerationUpdated = Schedule_OrderAPINeedToKnowReportGenerationUpdated & "," & GenerateOrderAPINeedToKnowReportFriday
	Schedule_OrderAPINeedToKnowReportGenerationUpdated = Schedule_OrderAPINeedToKnowReportGenerationUpdated & "," & GenerateOrderAPINeedToKnowReportSaturday
	
	Schedule_OrderAPINeedToKnowReportGenerationUpdated = Schedule_OrderAPINeedToKnowReportGenerationUpdated & "," & GenerateOrderAPINeedToKnowReportSundayTime
	Schedule_OrderAPINeedToKnowReportGenerationUpdated = Schedule_OrderAPINeedToKnowReportGenerationUpdated & "," & GenerateOrderAPINeedToKnowReportMondayTime
	Schedule_OrderAPINeedToKnowReportGenerationUpdated = Schedule_OrderAPINeedToKnowReportGenerationUpdated & "," & GenerateOrderAPINeedToKnowReportTuesdayTime
	Schedule_OrderAPINeedToKnowReportGenerationUpdated = Schedule_OrderAPINeedToKnowReportGenerationUpdated & "," & GenerateOrderAPINeedToKnowReportWednesdayTime
	Schedule_OrderAPINeedToKnowReportGenerationUpdated = Schedule_OrderAPINeedToKnowReportGenerationUpdated & "," & GenerateOrderAPINeedToKnowReportThursdayTime
	Schedule_OrderAPINeedToKnowReportGenerationUpdated = Schedule_OrderAPINeedToKnowReportGenerationUpdated & "," & GenerateOrderAPINeedToKnowReportFridayTime
	Schedule_OrderAPINeedToKnowReportGenerationUpdated = Schedule_OrderAPINeedToKnowReportGenerationUpdated & "," & GenerateOrderAPINeedToKnowReportSaturdayTime

	
	Schedule_OrderAPINeedToKnowReportGenerationUpdated = Schedule_OrderAPINeedToKnowReportGenerationUpdated & "," & RunOrderAPINeedToKnowReportIfClosed
	Schedule_OrderAPINeedToKnowReportGenerationUpdated = Schedule_OrderAPINeedToKnowReportGenerationUpdated & "," & RunOrderAPINeedToKnowReportIfClosingEarly
	
	Response.Write("<br><br><br>Schedule_OrderAPINeedToKnowReportGenerationUpdated: " & Schedule_OrderAPINeedToKnowReportGenerationUpdated)

	'*********************************************************************
	'Update SQL with Array/String of Schedule Data
	'*********************************************************************
		
	SQL = "UPDATE Settings_NeedToKnow SET Schedule_APINeedToKnowReportGeneration = '" & cStr(Schedule_OrderAPINeedToKnowReportGenerationUpdated) & "' "
	
	Response.Write("<br><br><br>SQL: " & SQL)
	
	Set cnn8 = Server.CreateObject("ADODB.Connection")
	cnn8.open (Session("ClientCnnString"))
	Set rs = Server.CreateObject("ADODB.Recordset")
	rs.CursorLocation = 3 
	Set rs = cnn8.Execute(SQL)
	
	set rs = Nothing
	cnn8.close
	set cnn8 = Nothing


	 Response.Redirect("order-api.asp")
	
%><!--#include file="../../../../inc/footer-main.asp"-->