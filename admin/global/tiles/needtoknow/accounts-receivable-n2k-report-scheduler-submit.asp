<!--#include file="../../../../inc/header.asp"-->

<%

	'***********************************************************
	'Get Values Of All Form Fields Posted 
	'***********************************************************
	
	GenerateFinanceNeedToKnowReportSunday = Request.Form("chkNoFinanceNeedToKnowReportSunday")
	GenerateFinanceNeedToKnowReportMonday = Request.Form("chkNoFinanceNeedToKnowReportMonday")
	GenerateFinanceNeedToKnowReportTuesday = Request.Form("chkNoFinanceNeedToKnowReportTuesday")
	GenerateFinanceNeedToKnowReportWednesday = Request.Form("chkNoFinanceNeedToKnowReportWednesday")
	GenerateFinanceNeedToKnowReportThursday = Request.Form("chkNoFinanceNeedToKnowReportThursday")
	GenerateFinanceNeedToKnowReportFriday = Request.Form("chkNoFinanceNeedToKnowReportFriday")
	GenerateFinanceNeedToKnowReportSaturday = Request.Form("chkNoFinanceNeedToKnowReportSaturday")
	
	GenerateFinanceNeedToKnowReportSundayTime = Request.Form("txtFinanceNeedToKnowReportSchedulerSundayTime")
	GenerateFinanceNeedToKnowReportMondayTime = Request.Form("txtFinanceNeedToKnowReportSchedulerMondayTime")
	GenerateFinanceNeedToKnowReportTuesdayTime = Request.Form("txtFinanceNeedToKnowReportSchedulerTuesdayTime")
	GenerateFinanceNeedToKnowReportWednesdayTime = Request.Form("txtFinanceNeedToKnowReportSchedulerWednesdayTime")
	GenerateFinanceNeedToKnowReportThursdayTime = Request.Form("txtFinanceNeedToKnowReportSchedulerThursdayTime")
	GenerateFinanceNeedToKnowReportFridayTime = Request.Form("txtFinanceNeedToKnowReportSchedulerFridayTime")
	GenerateFinanceNeedToKnowReportSaturdayTime = Request.Form("txtFinanceNeedToKnowReportSchedulerSaturdayTime")
	
	RunFinanceNeedToKnowReportIfClosed = Request.Form("chkNoFinanceNeedToKnowReportIfClosed")
	RunFinanceNeedToKnowReportIfClosingEarly = Request.Form("chkNoFinanceNeedToKnowReportIfClosingEarly")


	If Request.Form("chkNoFinanceNeedToKnowReportSunday") = "on" Then
		GenerateFinanceNeedToKnowReportSunday = 0
		GenerateFinanceNeedToKnowReportSundayTime = ""
	Else 
		GenerateFinanceNeedToKnowReportSunday = 1
	End If

	If Request.Form("chkNoFinanceNeedToKnowReportMonday") = "on" Then
		GenerateFinanceNeedToKnowReportMonday = 0
		GenerateFinanceNeedToKnowReportMondayTime = ""
	Else 
		GenerateFinanceNeedToKnowReportMonday = 1
	End If

	If Request.Form("chkNoFinanceNeedToKnowReportTuesday") = "on" Then
		GenerateFinanceNeedToKnowReportTuesday = 0
		GenerateFinanceNeedToKnowReportTuesdayTime = ""
	Else 
		GenerateFinanceNeedToKnowReportTuesday = 1
	End If

	If Request.Form("chkNoFinanceNeedToKnowReportWednesday") = "on" Then
		GenerateFinanceNeedToKnowReportWednesday = 0
		GenerateFinanceNeedToKnowReportWednesdayTime = ""
	Else 
		GenerateFinanceNeedToKnowReportWednesday = 1
	End If

	If Request.Form("chkNoFinanceNeedToKnowReportThursday") = "on" Then
		GenerateFinanceNeedToKnowReportThursday = 0
		GenerateFinanceNeedToKnowReportThursdayTime = ""
	Else 
		GenerateFinanceNeedToKnowReportThursday = 1
	End If

	If Request.Form("chkNoFinanceNeedToKnowReportFriday") = "on" Then
		GenerateFinanceNeedToKnowReportFriday = 0
		GenerateFinanceNeedToKnowReportFridayTime = ""
	Else 
		GenerateFinanceNeedToKnowReportFriday = 1
	End If

	If Request.Form("chkNoFinanceNeedToKnowReportSaturday") = "on" Then
		GenerateFinanceNeedToKnowReportSaturday = 0
		GenerateFinanceNeedToKnowReportSaturdayTime = ""
	Else 
		GenerateFinanceNeedToKnowReportSaturday = 1
	End If

	If Request.Form("chkNoFinanceNeedToKnowReportIfClosed") = "on" Then RunFinanceNeedToKnowReportIfClosed = 0 Else RunFinanceNeedToKnowReportIfClosed = 1
	If Request.Form("chkNoFinanceNeedToKnowReportIfClosingEarly") = "on" Then RunFinanceNeedToKnowReportIfClosingEarly = 0 Else RunFinanceNeedToKnowReportIfClosingEarly = 1
	
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
	
		Schedule_FinanceNeedToKnowReportGeneration = rsPropsectingSettings("Schedule_FinanceNeedToKnowReportGeneration")
		
		Schedule_FinanceNeedToKnowReportGenerationSettings = Split(Schedule_FinanceNeedToKnowReportGeneration,",")

		GenerateFinanceNeedToKnowReportSunday_ORIG = cInt(Schedule_FinanceNeedToKnowReportGenerationSettings(0))
		GenerateFinanceNeedToKnowReportMonday_ORIG = cInt(Schedule_FinanceNeedToKnowReportGenerationSettings(1))
		GenerateFinanceNeedToKnowReportTuesday_ORIG = cInt(Schedule_FinanceNeedToKnowReportGenerationSettings(2))
		GenerateFinanceNeedToKnowReportWednesday_ORIG = cInt(Schedule_FinanceNeedToKnowReportGenerationSettings(3))
		GenerateFinanceNeedToKnowReportThursday_ORIG = cInt(Schedule_FinanceNeedToKnowReportGenerationSettings(4))
		GenerateFinanceNeedToKnowReportFriday_ORIG = cInt(Schedule_FinanceNeedToKnowReportGenerationSettings(5))
		GenerateFinanceNeedToKnowReportSaturday_ORIG = cInt(Schedule_FinanceNeedToKnowReportGenerationSettings(6))
		GenerateFinanceNeedToKnowReportSundayTime_ORIG = Schedule_FinanceNeedToKnowReportGenerationSettings(7)
		GenerateFinanceNeedToKnowReportMondayTime_ORIG = Schedule_FinanceNeedToKnowReportGenerationSettings(8)
		GenerateFinanceNeedToKnowReportTuesdayTime_ORIG = Schedule_FinanceNeedToKnowReportGenerationSettings(9)
		GenerateFinanceNeedToKnowReportWednesdayTime_ORIG = Schedule_FinanceNeedToKnowReportGenerationSettings(10)
		GenerateFinanceNeedToKnowReportThursdayTime_ORIG = Schedule_FinanceNeedToKnowReportGenerationSettings(11)
		GenerateFinanceNeedToKnowReportFridayTime_ORIG = Schedule_FinanceNeedToKnowReportGenerationSettings(12)
		GenerateFinanceNeedToKnowReportSaturdayTime_ORIG = Schedule_FinanceNeedToKnowReportGenerationSettings(13)
		RunFinanceNeedToKnowReportIfClosed_ORIG = cInt(Schedule_FinanceNeedToKnowReportGenerationSettings(14))
		RunFinanceNeedToKnowReportIfClosingEarly_ORIG = cInt(Schedule_FinanceNeedToKnowReportGenerationSettings(15))
	
	End If
	
	set rsPropsectingSettings = Nothing
	cnnPropsectingSettings.close
	set cnnPropsectingSettings = Nothing
	

	'***********************************************************
	'Perform Audit Trail Entry Inserts
	'***********************************************************

	
	If Request.Form("chkNoFinanceNeedToKnowReportSunday") = "on" then GenerateFinanceNeedToKnowReportSundayMsg = "On" Else GenerateFinanceNeedToKnowReportSundayMsg = "Off"
	If GenerateFinanceNeedToKnowReportSunday_ORIG = 1 then GenerateFinanceNeedToKnowReportSundayMsgOrig = "On" Else GenerateFinanceNeedToKnowReportSundayMsgOrig = "Off"
	
	If GenerateFinanceNeedToKnowReportSunday <> GenerateFinanceNeedToKnowReportSunday_ORIG Then
		CreateAuditLogEntry "Global Settings Change", "Global Settings Change", "Major", 1, GetTerm("Accounts Receivable") & " Need To Know Report generation schedule for SUNDAY changed from " & GenerateFinanceNeedToKnowReportSundayMsgOrig & " to " & GenerateFinanceNeedToKnowReportSundayMsg
	End If
	
	If GenerateFinanceNeedToKnowReportSundayTime <> GenerateFinanceNeedToKnowReportSundayTime_ORIG Then
		If GenerateFinanceNeedToKnowReportSunday_ORIG = 0 AND GenerateFinanceNeedToKnowReportSunday = 1 Then
			CreateAuditLogEntry "Global Settings Change", "Global Settings Change", "Major", 1, GetTerm("Accounts Receivable") & " Need To Know Report generation schedule for SUNDAY turned on and set to run at " & GenerateFinanceNeedToKnowReportSundayTime
		Else
			CreateAuditLogEntry "Global Settings Change", "Global Settings Change", "Major", 1, GetTerm("Accounts Receivable") & " Need To Know Report generation scheduled run time for SUNDAY changed from " & GenerateFinanceNeedToKnowReportSundayTime_ORIG & " to " & GenerateFinanceNeedToKnowReportSundayTime
		End If
	End If
	
	

	If Request.Form("chkNoFinanceNeedToKnowReportMonday") = "on" then GenerateFinanceNeedToKnowReportMondayMsg = "On" Else GenerateFinanceNeedToKnowReportMondayMsg = "Off"
	If GenerateFinanceNeedToKnowReportMonday_ORIG = 1 then GenerateFinanceNeedToKnowReportMondayMsgOrig = "On" Else GenerateFinanceNeedToKnowReportMondayMsgOrig = "Off"
	
	If GenerateFinanceNeedToKnowReportMonday <> GenerateFinanceNeedToKnowReportMonday_ORIG Then
		CreateAuditLogEntry "Global Settings Change", "Global Settings Change", "Major", 1, GetTerm("Accounts Receivable") & " Need To Know Report generation schedule for Monday changed from " & GenerateFinanceNeedToKnowReportMondayMsgOrig & " to " & GenerateFinanceNeedToKnowReportMondayMsg
	End If
	
	If GenerateFinanceNeedToKnowReportMondayTime <> GenerateFinanceNeedToKnowReportMondayTime_ORIG Then
		If GenerateFinanceNeedToKnowReportMonday_ORIG = 0 AND GenerateFinanceNeedToKnowReportMonday = 1 Then
			CreateAuditLogEntry "Global Settings Change", "Global Settings Change", "Major", 1, GetTerm("Accounts Receivable") & " Need To Know Report generation schedule for Monday turned on and set to run at " & GenerateFinanceNeedToKnowReportMondayTime
		Else
			CreateAuditLogEntry "Global Settings Change", "Global Settings Change", "Major", 1, GetTerm("Accounts Receivable") & " Need To Know Report generation scheduled run time for Monday changed from " & GenerateFinanceNeedToKnowReportMondayTime_ORIG & " to " & GenerateFinanceNeedToKnowReportMondayTime
		End If
	End If
	


	If Request.Form("chkNoFinanceNeedToKnowReportTuesday") = "on" then GenerateFinanceNeedToKnowReportTuesdayMsg = "On" Else GenerateFinanceNeedToKnowReportTuesdayMsg = "Off"
	If GenerateFinanceNeedToKnowReportTuesday_ORIG = 1 then GenerateFinanceNeedToKnowReportTuesdayMsgOrig = "On" Else GenerateFinanceNeedToKnowReportTuesdayMsgOrig = "Off"
	
	If GenerateFinanceNeedToKnowReportTuesday <> GenerateFinanceNeedToKnowReportTuesday_ORIG Then
		CreateAuditLogEntry "Global Settings Change", "Global Settings Change", "Major", 1, GetTerm("Accounts Receivable") & " Need To Know Report generation schedule for Tuesday changed from " & GenerateFinanceNeedToKnowReportTuesdayMsgOrig & " to " & GenerateFinanceNeedToKnowReportTuesdayMsg
	End If
	
	If GenerateFinanceNeedToKnowReportTuesdayTime <> GenerateFinanceNeedToKnowReportTuesdayTime_ORIG Then
		If GenerateFinanceNeedToKnowReportTuesday_ORIG = 0 AND GenerateFinanceNeedToKnowReportTuesday = 1 Then
			CreateAuditLogEntry "Global Settings Change", "Global Settings Change", "Major", 1, GetTerm("Accounts Receivable") & " Need To Know Report generation schedule for Tuesday turned on and set to run at " & GenerateFinanceNeedToKnowReportTuesdayTime
		Else
			CreateAuditLogEntry "Global Settings Change", "Global Settings Change", "Major", 1, GetTerm("Accounts Receivable") & " Need To Know Report generation scheduled run time for Tuesday changed from " & GenerateFinanceNeedToKnowReportTuesdayTime_ORIG & " to " & GenerateFinanceNeedToKnowReportTuesdayTime
		End If
	End If



	If Request.Form("chkNoFinanceNeedToKnowReportWednesday") = "on" then GenerateFinanceNeedToKnowReportWednesdayMsg = "On" Else GenerateFinanceNeedToKnowReportWednesdayMsg = "Off"
	If GenerateFinanceNeedToKnowReportWednesday_ORIG = 1 then GenerateFinanceNeedToKnowReportWednesdayMsgOrig = "On" Else GenerateFinanceNeedToKnowReportWednesdayMsgOrig = "Off"
	
	If GenerateFinanceNeedToKnowReportWednesday <> GenerateFinanceNeedToKnowReportWednesday_ORIG Then
		CreateAuditLogEntry "Global Settings Change", "Global Settings Change", "Major", 1, GetTerm("Accounts Receivable") & " Need To Know Report generation schedule for Wednesday changed from " & GenerateFinanceNeedToKnowReportWednesdayMsgOrig & " to " & GenerateFinanceNeedToKnowReportWednesdayMsg
	End If
	
	If GenerateFinanceNeedToKnowReportWednesdayTime <> GenerateFinanceNeedToKnowReportWednesdayTime_ORIG Then
		If GenerateFinanceNeedToKnowReportWednesday_ORIG = 0 AND GenerateFinanceNeedToKnowReportWednesday = 1 Then
			CreateAuditLogEntry "Global Settings Change", "Global Settings Change", "Major", 1, GetTerm("Accounts Receivable") & " Need To Know Report generation schedule for Wednesday turned on and set to run at " & GenerateFinanceNeedToKnowReportWednesdayTime
		Else
			CreateAuditLogEntry "Global Settings Change", "Global Settings Change", "Major", 1, GetTerm("Accounts Receivable") & " Need To Know Report generation scheduled run time for Wednesday changed from " & GenerateFinanceNeedToKnowReportWednesdayTime_ORIG & " to " & GenerateFinanceNeedToKnowReportWednesdayTime
		End If
	End If



	If Request.Form("chkNoFinanceNeedToKnowReportThursday") = "on" then GenerateFinanceNeedToKnowReportThursdayMsg = "On" Else GenerateFinanceNeedToKnowReportThursdayMsg = "Off"
	If GenerateFinanceNeedToKnowReportThursday_ORIG = 1 then GenerateFinanceNeedToKnowReportThursdayMsgOrig = "On" Else GenerateFinanceNeedToKnowReportThursdayMsgOrig = "Off"
	
	If GenerateFinanceNeedToKnowReportThursday <> GenerateFinanceNeedToKnowReportThursday_ORIG Then
		CreateAuditLogEntry "Global Settings Change", "Global Settings Change", "Major", 1, GetTerm("Accounts Receivable") & " Need To Know Report generation schedule for Thursday changed from " & GenerateFinanceNeedToKnowReportThursdayMsgOrig & " to " & GenerateFinanceNeedToKnowReportThursdayMsg
	End If
	
	If GenerateFinanceNeedToKnowReportThursdayTime <> GenerateFinanceNeedToKnowReportThursdayTime_ORIG Then
		If GenerateFinanceNeedToKnowReportThursday_ORIG = 0 AND GenerateFinanceNeedToKnowReportThursday = 1 Then
			CreateAuditLogEntry "Global Settings Change", "Global Settings Change", "Major", 1, GetTerm("Accounts Receivable") & " Need To Know Report generation schedule for Thursday turned on and set to run at " & GenerateFinanceNeedToKnowReportThursdayTime
		Else
			CreateAuditLogEntry "Global Settings Change", "Global Settings Change", "Major", 1, GetTerm("Accounts Receivable") & " Need To Know Report generation scheduled run time for Thursday changed from " & GenerateFinanceNeedToKnowReportThursdayTime_ORIG & " to " & GenerateFinanceNeedToKnowReportThursdayTime
		End If
	End If



	If Request.Form("chkNoFinanceNeedToKnowReportFriday") = "on" then GenerateFinanceNeedToKnowReportFridayMsg = "On" Else GenerateFinanceNeedToKnowReportFridayMsg = "Off"
	If GenerateFinanceNeedToKnowReportFriday_ORIG = 1 then GenerateFinanceNeedToKnowReportFridayMsgOrig = "On" Else GenerateFinanceNeedToKnowReportFridayMsgOrig = "Off"
	
	If GenerateFinanceNeedToKnowReportFriday <> GenerateFinanceNeedToKnowReportFriday_ORIG Then
		CreateAuditLogEntry "Global Settings Change", "Global Settings Change", "Major", 1, GetTerm("Accounts Receivable") & " Need To Know Report generation schedule for Friday changed from " & GenerateFinanceNeedToKnowReportFridayMsgOrig & " to " & GenerateFinanceNeedToKnowReportFridayMsg
	End If
	
	If GenerateFinanceNeedToKnowReportFridayTime <> GenerateFinanceNeedToKnowReportFridayTime_ORIG Then
		If GenerateFinanceNeedToKnowReportFriday_ORIG = 0 AND GenerateFinanceNeedToKnowReportFriday = 1 Then
			CreateAuditLogEntry "Global Settings Change", "Global Settings Change", "Major", 1, GetTerm("Accounts Receivable") & " Need To Know Report generation schedule for Friday turned on and set to run at " & GenerateFinanceNeedToKnowReportFridayTime
		Else
			CreateAuditLogEntry "Global Settings Change", "Global Settings Change", "Major", 1, GetTerm("Accounts Receivable") & " Need To Know Report generation scheduled run time for Friday changed from " & GenerateFinanceNeedToKnowReportFridayTime_ORIG & " to " & GenerateFinanceNeedToKnowReportFridayTime
		End If
	End If



	If Request.Form("chkNoFinanceNeedToKnowReportSaturday") = "on" then GenerateFinanceNeedToKnowReportSaturdayMsg = "On" Else GenerateFinanceNeedToKnowReportSaturdayMsg = "Off"
	If GenerateFinanceNeedToKnowReportSaturday_ORIG = 1 then GenerateFinanceNeedToKnowReportSaturdayMsgOrig = "On" Else GenerateFinanceNeedToKnowReportSaturdayMsgOrig = "Off"
	
	If GenerateFinanceNeedToKnowReportSaturday <> GenerateFinanceNeedToKnowReportSaturday_ORIG Then
		CreateAuditLogEntry "Global Settings Change", "Global Settings Change", "Major", 1, GetTerm("Accounts Receivable") & " Need To Know Report generation schedule for Saturday changed from " & GenerateFinanceNeedToKnowReportSaturdayMsgOrig & " to " & GenerateFinanceNeedToKnowReportSaturdayMsg
	End If
	
	If GenerateFinanceNeedToKnowReportSaturdayTime <> GenerateFinanceNeedToKnowReportSaturdayTime_ORIG Then
		If GenerateFinanceNeedToKnowReportSaturday_ORIG = 0 AND GenerateFinanceNeedToKnowReportSaturday = 1 Then
			CreateAuditLogEntry "Global Settings Change", "Global Settings Change", "Major", 1, GetTerm("Accounts Receivable") & " Need To Know Report generation schedule for Saturday turned on and set to run at " & GenerateFinanceNeedToKnowReportSaturdayTime
		Else
			CreateAuditLogEntry "Global Settings Change", "Global Settings Change", "Major", 1, GetTerm("Accounts Receivable") & " Need To Know Report generation scheduled run time for Saturday changed from " & GenerateFinanceNeedToKnowReportSaturdayTime_ORIG & " to " & GenerateFinanceNeedToKnowReportSaturdayTime
		End If
	End If


	If Request.Form("chkNoFinanceNeedToKnowReportIfClosed") = "on" then RunFinanceNeedToKnowReportIfClosedMsg = "On" Else RunFinanceNeedToKnowReportIfClosedMsg = "Off"
	If RunFinanceNeedToKnowReportIfClosed_ORIG = 1 then RunFinanceNeedToKnowReportIfClosedMsgOrig = "On" Else RunFinanceNeedToKnowReportIfClosedMsgOrig = "Off"
	
	If RunFinanceNeedToKnowReportIfClosed <> RunFinanceNeedToKnowReportIfClosed_ORIG Then
		CreateAuditLogEntry "Global Settings Change", "Global Settings Change", "Major", 1, GetTerm("Accounts Receivable") & " Need To Know Report generation schedule - Don''t Run If Closed (Monday-Friday Only) -  changed from " & RunFinanceNeedToKnowReportIfClosedMsgOrig & " to " & RunFinanceNeedToKnowReportIfClosedMsg
	End If


	If Request.Form("chkNoFinanceNeedToKnowReportIfClosingEarly") = "on" then RunFinanceNeedToKnowReportIfClosingEarlyMsg = "On" Else RunFinanceNeedToKnowReportIfClosingEarlyMsg = "Off"
	If RunFinanceNeedToKnowReportIfClosingEarly_ORIG = 1 then RunFinanceNeedToKnowReportIfClosingEarlyMsgOrig = "On" Else RunFinanceNeedToKnowReportIfClosingEarlyMsgOrig = "Off"
	
	If RunFinanceNeedToKnowReportIfClosingEarly <> RunFinanceNeedToKnowReportIfClosingEarly_ORIG Then
		CreateAuditLogEntry "Global Settings Change", "Global Settings Change", "Major", 1, GetTerm("Accounts Receivable") & " Need To Know Report generation schedule - Don''t Run If Closed (Monday-Friday Only) -  changed from " & RunFinanceNeedToKnowReportIfClosingEarlyMsgOrig & " to " & RunFinanceNeedToKnowReportIfClosingEarlyMsg
	End If


	'*********************************************************************
	'Build Array/String of Schedule Data From Request Form Field Data
	'*********************************************************************

	Schedule_FinanceNeedToKnowReportGenerationUpdated = ""
	
	Schedule_FinanceNeedToKnowReportGenerationUpdated = GenerateFinanceNeedToKnowReportSunday
	Schedule_FinanceNeedToKnowReportGenerationUpdated = Schedule_FinanceNeedToKnowReportGenerationUpdated & "," & GenerateFinanceNeedToKnowReportMonday
	Schedule_FinanceNeedToKnowReportGenerationUpdated = Schedule_FinanceNeedToKnowReportGenerationUpdated & "," & GenerateFinanceNeedToKnowReportTuesday
	Schedule_FinanceNeedToKnowReportGenerationUpdated = Schedule_FinanceNeedToKnowReportGenerationUpdated & "," & GenerateFinanceNeedToKnowReportWednesday
	Schedule_FinanceNeedToKnowReportGenerationUpdated = Schedule_FinanceNeedToKnowReportGenerationUpdated & "," & GenerateFinanceNeedToKnowReportThursday
	Schedule_FinanceNeedToKnowReportGenerationUpdated = Schedule_FinanceNeedToKnowReportGenerationUpdated & "," & GenerateFinanceNeedToKnowReportFriday
	Schedule_FinanceNeedToKnowReportGenerationUpdated = Schedule_FinanceNeedToKnowReportGenerationUpdated & "," & GenerateFinanceNeedToKnowReportSaturday
	
	Schedule_FinanceNeedToKnowReportGenerationUpdated = Schedule_FinanceNeedToKnowReportGenerationUpdated & "," & GenerateFinanceNeedToKnowReportSundayTime
	Schedule_FinanceNeedToKnowReportGenerationUpdated = Schedule_FinanceNeedToKnowReportGenerationUpdated & "," & GenerateFinanceNeedToKnowReportMondayTime
	Schedule_FinanceNeedToKnowReportGenerationUpdated = Schedule_FinanceNeedToKnowReportGenerationUpdated & "," & GenerateFinanceNeedToKnowReportTuesdayTime
	Schedule_FinanceNeedToKnowReportGenerationUpdated = Schedule_FinanceNeedToKnowReportGenerationUpdated & "," & GenerateFinanceNeedToKnowReportWednesdayTime
	Schedule_FinanceNeedToKnowReportGenerationUpdated = Schedule_FinanceNeedToKnowReportGenerationUpdated & "," & GenerateFinanceNeedToKnowReportThursdayTime
	Schedule_FinanceNeedToKnowReportGenerationUpdated = Schedule_FinanceNeedToKnowReportGenerationUpdated & "," & GenerateFinanceNeedToKnowReportFridayTime
	Schedule_FinanceNeedToKnowReportGenerationUpdated = Schedule_FinanceNeedToKnowReportGenerationUpdated & "," & GenerateFinanceNeedToKnowReportSaturdayTime

	
	Schedule_FinanceNeedToKnowReportGenerationUpdated = Schedule_FinanceNeedToKnowReportGenerationUpdated & "," & RunFinanceNeedToKnowReportIfClosed
	Schedule_FinanceNeedToKnowReportGenerationUpdated = Schedule_FinanceNeedToKnowReportGenerationUpdated & "," & RunFinanceNeedToKnowReportIfClosingEarly
	
	Response.Write("<br><br><br>Schedule_FinanceNeedToKnowReportGenerationUpdated: " & Schedule_FinanceNeedToKnowReportGenerationUpdated)

	'*********************************************************************
	'Update SQL with Array/String of Schedule Data
	'*********************************************************************
		
	SQL = "UPDATE Settings_NeedToKnow SET Schedule_FinanceNeedToKnowReportGeneration = '" & cStr(Schedule_FinanceNeedToKnowReportGenerationUpdated) & "' "
	
	Response.Write("<br><br><br>SQL: " & SQL)
	
	Set cnn8 = Server.CreateObject("ADODB.Connection")
	cnn8.open (Session("ClientCnnString"))
	Set rs = Server.CreateObject("ADODB.Recordset")
	rs.CursorLocation = 3 
	Set rs = cnn8.Execute(SQL)
	
	set rs = Nothing
	cnn8.close
	set cnn8 = Nothing


	 Response.Redirect("accounts-receivable.asp")
	
%><!--#include file="../../../../inc/footer-main.asp"-->