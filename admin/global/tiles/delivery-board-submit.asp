<!--#include file="../../../inc/header.asp"-->
<!--#include file="../../../inc/Insightfuncs_Routing.asp"-->
<%

	'***********************************************************
	'Get Values Of All Form Fields Posted
	'***********************************************************
		
	DelBoardScheduledColor = Request.Form("txtScheduledColor")
	DelBoardCompletedColor = Request.Form("txtCompletedDeliveries")
	DelBoardInProgressColor = Request.Form("txtInProgress")
	DelBoardSkippedColor = Request.Form("txtSkippedDeliveries")
	DelBoardNextStopColor = Request.Form("txtNextStop")
	DelBoardAMColor = Request.Form("txtDelBoardAMColor")
	DelBoardPriorityColor = Request.Form("txtDelBoardPriorityColor")
	DelBoardPieTimerColor = Request.Form("txtDelBoardPieTimerColor")
	DelBoardTitleGradientColor = Request.Form("txtDelBoardTitleGradientColor")
	DelBoardTitleText = Request.Form("txtDelBoardTitleText")
	DelBoardTitleText = Replace(DelBoardTitleText,"'","")
	DelBoardTitleTextFontColor = Request.Form("txtDelBoardTitleTextFontColor")
	DelBoardProfitDollars = Request.Form("selProfitDollars")
	DelBoardAtOrAboveProfitColor = Request.Form("txtAboveProfit")
	DelBoardBelowProfitColor = Request.Form("txtBelowProfit")
	DelBoardUserAlertColor = Request.Form("txtUserDefinedAlert")
	DelBoardRoutesToIgnore = Replace(Request.Form("txtDelBoardRoutesToIgnore")," ","")
	DelBoardUPSRoutes = Replace(Request.Form("txtDelBoardUPSRoutes")," ","")
	
	If MUV_Read("routingModuleOn") = "Enabled" Then	
		If Request.Form("chkNextStopNagMessageONOFF") = "on" then NextStopNagMessageONOFF = 1 Else NextStopNagMessageONOFF = 0
		NextStopNagMinutes = Request.Form("selNextStopNagMinutes")
		NextStopNagIntervalMinutes = Request.Form("selNextStopNagIntervalMinutes")
		NextStopNagMessageMaxToSendPerStop = Request.Form("selNextStopNagMessageMaxToSendPerStop")
		NextStopNagMessageMaxToSendPerDriverPerDay = Request.Form("selNextStopNagMessageMaxToSendPerDriverPerDay")
		NextStopNagMessageSendMethod = Request.Form("selNextStopNagMessageSendMethod")
		
		If Request.Form("chkNoActivityNagMessageONOFF") = "on" then NoActivityNagMessageONOFF = 1 Else NoActivityNagMessageONOFF = 0
		NoActivityNagMinutes = Request.Form("selNoActivityNagMinutes")
		NoActivityNagIntervalMinutes = Request.Form("selNoActivityNagIntervalMinutes")
		NoActivityNagMessageMaxToSendPerStop = Request.Form("selNoActivityNagMessageMaxToSendPerStop")
		NoActivityNagMessageMaxToSendPerDriverPerDay = Request.Form("selNoActivityNagMessageMaxToSendPerDriverPerDay")
		NoActivityNagMessageSendMethod = Request.Form("selNoActivityNagMessageSendMethod")
		NoActivityNagTimeOfDay = Request.Form("selNoActivityNagTimeOfDay")
	End If
	

	If Request.Form("chkDoNotShowDeliveryLineItems") = "on" then DoNotShowDeliveryLineItems = 1 Else DoNotShowDeliveryLineItems = 0
	If Request.Form("chkAutoPromptNextStop") = "on" then AutoPromptNextStop = 1 Else AutoPromptNextStop = 0
	If Request.Form("chkAutoForceSelectNextStop") = "on" then AutoForceSelectNextStop = 1 Else AutoForceSelectNextStop = 0
	If Request.Form("chkDelBoardDontUseStopSequence") = "on" then DelBoardDontUseStopSequence = 1 Else DelBoardDontUseStopSequence = 0
	



	'***********************************************************
	'Get Original Values For Audit Trail Entries
	'***********************************************************
	
	SQL = "SELECT * FROM Settings_Global"
	
	Set cnn8 = Server.CreateObject("ADODB.Connection")
	cnn8.open (Session("ClientCnnString"))
	Set rs = Server.CreateObject("ADODB.Recordset")
	rs.CursorLocation = 3 
	Set rs = cnn8.Execute(SQL)
	
	If not rs.EOF Then
	
		DelBoardScheduledColor_ORIG = rs("DelBoardScheduledColor")	
		DelBoardCompletedColor_ORIG = rs("DelBoardCompletedColor")				
		DelBoardInProgressColor_ORIG = rs("DelBoardInProgressColor")						
		DelBoardSkippedColor_ORIG = rs("DelBoardSkippedColor")		
		DelBoardNextStopColor_ORIG = rs("DelBoardNextStopColor")	
		DelBoardAMColor_ORIG = rs("DelBoardAMColor")
		DelBoardPriorityColor_ORIG = rs("DelBoardPriorityColor")	
		DelBoardPieTimerColor_ORIG = rs("DelBoardPieTimerColor")	
		DelBoardTitleGradientColor_ORIG = rs("DelBoardTitleGradientColor")	
		DelBoardTitleText_ORIG = rs("DelBoardTitleText")	
		DelBoardTitleTextFontColor_ORIG = rs("DelBoardTitleTextFontColor")					
		DelBoardProfitDollars_ORIG = rs("DelBoardProfitDollars")			
		DelBoardAtOrAboveProfitColor_ORIG = rs("DelBoardAtOrAboveProfitColor")			
		DelBoardBelowProfitColor_ORIG = rs("DelBoardBelowProfitColor")			
		DelBoardUserAlertColor_ORIG = rs("DelBoardUserAlertColor")	
		AutoPromptNextStop_ORIG = rs("AutoPromptNextStop")
		AutoForceSelectNextStop_ORIG = rs("AutoForceSelectNextStop")
		DoNotShowDeliveryLineItems_ORIG = rs("DoNotShowDeliveryLineItems")
		DelBoardDontUseStopSequence_ORIG = rs("DelBoardDontUseStopSequencing")
		DelBoardRoutesToIgnore_ORIG = rs("DelBoardRoutesToIgnore")
		DelBoardUPSRoutes_ORIG = rs("DelBoardUPSRoutes")		
			
		If MUV_Read("routingModuleOn") = "Enabled" Then	
		
			NextStopNagMessageONOFF_ORIG = rs("NextStopNagMessageONOFF")
			NextStopNagMinutes_ORIG = rs("NextStopNagMinutes")
			NextStopNagIntervalMinutes_ORIG = rs("NextStopNagIntervalMinutes")
			NextStopNagMessageMaxToSendPerStop_ORIG = rs("NextStopNagMessageMaxToSendPerStop")
			NextStopNagMessageMaxToSendPerDriverPerDay_ORIG = rs("NextStopNagMessageMaxToSendPerDriverPerDay")
			NextStopNagMessageSendMethod_ORIG = rs("NextStopNagMessageSendMethod")
			NoActivityNagMessageONOFF_ORIG = rs("NoActivityNagMessageONOFF")
			NoActivityNagMinutes_ORIG = rs("NoActivityNagMinutes")
			NoActivityNagIntervalMinutes_ORIG = rs("NoActivityNagIntervalMinutes")
			NoActivityNagMessageMaxToSendPerStop_ORIG = rs("NoActivityNagMessageMaxToSendPerStop")
			NoActivityNagMessageMaxToSendPerDriverPerDay_ORIG = rs("NoActivityNagMessageMaxToSendPerDriverPerDay")
			NoActivityNagMessageSendMethod_ORIG = rs("NoActivityNagMessageSendMethod")
			NoActivityNagTimeOfDay_ORIG = rs("NoActivityNagTimeOfDay")
		End If   		
		
	End If


	'***********************************************************
	'Perform Audit Trail Entry Inserts
	'***********************************************************		
			
	If DelBoardScheduledColor <> DelBoardScheduledColor_ORIG Then
		CreateAuditLogEntry "Global Settings Change", "Global Settings Change", "Major", 1, "Delivery Board - highlight color for scheduled deliveries changed from  " & DelBoardScheduledColor_ORIG & " to " & DelBoardScheduledColor
	End If
	If DelBoardCompletedColor <> DelBoardCompletedColor_ORIG Then
		CreateAuditLogEntry "Global Settings Change", "Global Settings Change", "Major", 1, "Delivery Board - highlight color for completed deliveries changed from  " & DelBoardCompletedColor_ORIG & " to " & DelBoardCompletedColor
	End If
	If DelBoardInProgressColor <> DelBoardInProgressColor_ORIG Then
		CreateAuditLogEntry "Global Settings Change", "Global Settings Change", "Major", 1, "Delivery Board - highlight color for In Progress deliveries changed from  " & DelBoardInProgressColor_ORIG & " to " & DelBoardInProgressColor
	End If
	If DelBoardSkippedColor <> DelBoardSkippedColor_ORIG Then
		CreateAuditLogEntry "Global Settings Change", "Global Settings Change", "Major", 1, "Delivery Board - highlight color for skipped deliveries changed from  " & DelBoardSkippedColor_ORIG & " to " & DelBoardSkippedColor
	End If
	If DelBoardNextStopColor <> DelBoardNextStopColor_ORIG Then
		CreateAuditLogEntry "Global Settings Change", "Global Settings Change", "Major", 1, "Delivery Board - highlight color for next stop changed from  " & DelBoardNextStopColor_ORIG & " to " & DelBoardNextStopColor
	End If
	If DelBoardAMColor <> DelBoardAMColor_ORIG Then
		CreateAuditLogEntry "Global Settings Change", "Global Settings Change", "Major", 1, "Delivery Board - highlight color for AM deliveries changed from  " & DelBoardAMColor_ORIG & " to " & DelBoardAMColor
	End If
	If DelBoardPriorityColor <> DelBoardPriorityColor_ORIG Then
		CreateAuditLogEntry "Global Settings Change", "Global Settings Change", "Major", 1, "Delivery Board - highlight color for Priority deliveries changed from  " & DelBoardPriorityColor_ORIG & " to " & DelBoardPriorityColor
	End If
	If DelBoardPieTimerColor <> DelBoardPieTimerColor_ORIG Then
		CreateAuditLogEntry "Global Settings Change", "Global Settings Change", "Major", 1, "Delivery Board - Pie timer color changed from  " & DelBoardPieTimerColor_ORIG & " to " & DelBoardPieTimerColor
	End If
	If DelBoardTitleGradientColor <> DelBoardTitleGradientColor_ORIG Then
		CreateAuditLogEntry "Global Settings Change", "Global Settings Change", "Major", 1, "Delivery Board - title bar gradient and border color changed from " & DelBoardTitleGradientColor_ORIG & " to " & DelBoardTitleGradientColor
	End If
	If DelBoardTitleText <> DelBoardTitleText_ORIG Then
		CreateAuditLogEntry "Global Settings Change", "Global Settings Change", "Major", 1, "Delivery Board - title text changed from " & DelBoardTitleText_ORIG & " to " & DelBoardTitleText
	End If
	If DelBoardTitleTextFontColor <> DelBoardTitleTextFontColor_ORIG Then
		CreateAuditLogEntry "Global Settings Change", "Global Settings Change", "Major", 1, "Delivery Board - title text font color changed from " & DelBoardTitleTextFontColor_ORIG & " to " & DelBoardTitleTextFontColor
	End If
	If DelBoardAtOrAboveProfitColor <> DelBoardAtOrAboveProfitColor_ORIG Then
		CreateAuditLogEntry "Global Settings Change", "Global Settings Change", "Major", 1, "Delivery Board -  highlight color for routes at or above profit target changed from " & DelBoardTitleGradientColor_ORIG & " to " & DelBoardTitleGradientColor
	End If
	If DelBoardBelowProfitColor <> DelBoardBelowProfitColor_ORIG Then
		CreateAuditLogEntry "Global Settings Change", "Global Settings Change", "Major", 1, "Delivery Board - highlight color for routes below profit target changed from  " & DelBoardBelowProfitColor_ORIG & " to " & DelBoardBelowProfitColor
	End If
	If DelBoardUserAlertColor <> DelBoardUserAlertColor_ORIG Then
		CreateAuditLogEntry "Global Settings Change", "Global Settings Change", "Major", 1, "Delivery Board - highlight color for deliveries with user defined alerts changed from  " & DelBoardUserAlertColor_ORIG & " to " & DelBoardUserAlertColor
	End If
	If cdbl(DelBoardProfitDollars) <> cdbl(DelBoardProfitDollars_ORIG) Then
		CreateAuditLogEntry "Global Settings Change", "Global Settings Change", "Major", 1, "Delivery Board - route profitability $ target changed from  " & DelBoardProfitDollars_ORIG & " to " & DelBoardProfitDollars
	End If	
	If DelBoardRoutesToIgnore <> DelBoardRoutesToIgnore_ORIG Then
		CreateAuditLogEntry "Global Settings Change", "Global Settings Change", "Major", 1, "Delivery Board routes to ignore changed from  " & DelBoardRoutesToIgnore_ORIG & " to " & DelBoardRoutesToIgnore
	End If
	If DelBoardUPSRoutes <> DelBoardUPSRoutes_ORIG Then
		CreateAuditLogEntry "Global Settings Change", "Global Settings Change", "Major", 1, "Delivery Board UPS routes changed from  " & DelBoardUPSRoutes_ORIG & " to " & DelBoardUPSRoutes
	End If
	
	If Request.Form("chkDoNotShowDeliveryLineItems")="on" then DoNotShowDeliveryLineItemsMsg = "On" Else DoNotShowDeliveryLineItemsMsg = "Off"

	IF DoNotShowDeliveryLineItems <> DoNotShowDeliveryLineItems_ORIG Then
		CreateAuditLogEntry "Global Settings Change", "Global Settings Change", "Major", 1, "Do Not Show Order Line Items For Deliveries (mobile webapp) changed from " & DoNotShowDeliveryLineItems_ORIG & " to " & DoNotShowDeliveryLineItemsMsg 
	End If
	
	If Request.Form("chkAutoPromptNextStop")="on" then AutoPromptNextStopMsg = "On" Else AutoPromptNextStopMsg = "Off"
	
	IF AutoPromptNextStop <> AutoPromptNextStop_ORIG Then
		CreateAuditLogEntry "Global Settings Change", "Global Settings Change", "Major", 1, "Automatically prompt for next stop (mobile webapp) changed from " & AutoPromptNextStop_ORIG & " to " & AutoPromptNextStopMsg 
	End If
	
	If Request.Form("chkAutoForceSelectNextStop")="on" then AutoForceSelectNextStopMsg = "On" Else AutoForceSelectNextStopMsg = "Off"
	
	IF AutoForceSelectNextStop <> AutoForceSelectNextStop_ORIG Then
		CreateAuditLogEntry "Global Settings Change", "Global Settings Change", "Major", 1, "Automatically force drive to select next stop (mobile webapp) changed from " & AutoForceSelectNextStop_ORIG & " to " & AutoForceSelectNextStopMsg 
	End If
	
	If Request.Form("chkDelBoardDontUseStopSequence")="on" then DelBoardDontUseStopSequenceMsg = "On" Else DelBoardDontUseStopSequenceMsg = "Off"
	
	IF DelBoardDontUseStopSequence <> DelBoardDontUseStopSequencing_ORIG Then
		CreateAuditLogEntry "Global Settings Change", "Global Settings Change", "Major", 1, "Do not use stop sequencing changed from " & DelBoardDontUseStopSequencing_ORIG & " to " & DelBoardDontUseStopSequenceMsg 
	End If	
	
	If cInt(NoActivityNagMessageONOFF) = 1 then NoActivityNagMessageONOFFMsg = "On" Else NoActivityNagMessageONOFFMsg = "Off"
			
	IF cInt(NoActivityNagMessageONOFF) <> cInt(NoActivityNagMessageONOFF_ORIG) Then
		CreateAuditLogEntry "Global Settings Change", "Global Settings Change", "Major", 1, GetTerm("Routing") & " - Send nag messages when there has been a period of No Activity changed from " & NoActivityNagMessageONOFF_ORIG & " to " & NoActivityNagMessageONOFFMsg
	End If

	If NoActivityNagMessageSendMethod <> NoActivityNagMessageSendMethod_ORIG Then
		CreateAuditLogEntry "Global Settings Change", "Global Settings Change", "Major", 1, GetTerm("Routing") & " - nag send method changed from  " & NoActivityNagMessageSendMethod_ORIG & " to " & NoActivityNagMessageSendMethod
	End If

	If cint(NoActivityNagMinutes) <> cint(NoActivityNagMinutes_ORIG) Then
		CreateAuditLogEntry "Global Settings Change", "Global Settings Change", "Major", 1, GetTerm("Routing") & " - Send when there has been No Activity for X minutes changed from " & NoActivityNagMinutes_ORIG & " to " & NoActivityNagMinutes
	End If
	
	If cint(NoActivityNagMessageMaxToSendPerDriverPerDay) <> cint(NoActivityNagMessageMaxToSendPerDriverPerDay_ORIG) Then
		CreateAuditLogEntry "Global Settings Change", "Global Settings Change", "Major", 1, GetTerm("Routing") & " - Send a maximum of X messages to any driver on a given day changed from " & NoActivityNagMessageMaxToSendPerDriverPerDay_ORIG & " to " & NoActivityNagMessageMaxToSendPerDriverPerDay
	End If
	
	If cint(NoActivityNagMessageMaxToSendPerStop) <> cint(NoActivityNagMessageMaxToSendPerStop_ORIG) Then
		CreateAuditLogEntry "Global Settings Change", "Global Settings Change", "Major", 1, GetTerm("Routing") & " - Send a maximum of X messages each time a No Activity event occurs changed from " & NoActivityNagMessageMaxToSendPerStop_ORIG & " to " & NoActivityNagMessageMaxToSendPerStop
	End If

	If cint(NoActivityNagIntervalMinutes) <> cint(NoActivityNagIntervalMinutes_ORIG) Then
		CreateAuditLogEntry "Global Settings Change", "Global Settings Change", "Major", 1, GetTerm("Routing") & " - Continue to send nag alerts every X minutes changed from " & NoActivityNagIntervalMinutes_ORIG & " to " & NoActivityNagIntervalMinutes
	End If

	If NoActivityNagTimeOfDay <> NoActivityNagTimeOfDay_ORIG Then
		CreateAuditLogEntry "Global Settings Change", "Global Settings Change", "Major", 1, GetTerm("Routing") & " - Send nag when there has been No Activity by X time changed from " & NoActivityNagTimeOfDay_ORIG & " to " & NoActivityNagTimeOfDay
	End If
	
	If cInt(NextStopNagMessageONOFF) = 1  then NextStopNagMessageONOFFMsg = "On" Else NextStopNagMessageONOFFMsg = "Off"
	
	IF cInt(NextStopNagMessageONOFF) <> cInt(NextStopNagMessageONOFF_ORIG) Then
		CreateAuditLogEntry "Global Settings Change", "Global Settings Change", "Major", 1, GetTerm("Routing") & " - Send nag messages when a driver does not select the Next Stop changed from " & NextStopNagMessageONOFF_ORIG & " to " & NextStopNagMessageONOFFMsg
	End If

	If NextStopNagMessageSendMethod <> NextStopNagMessageSendMethod_ORIG Then
		CreateAuditLogEntry "Global Settings Change", "Global Settings Change", "Major", 1, GetTerm("Routing") & " - nag Next Stop send method changed from  " & NextStopNagMessageSendMethod_ORIG & " to " & NextStopNagMessageSendMethod
	End If

	If cint(NextStopNagMinutes) <> cint(NextStopNagMinutes_ORIG) Then
		CreateAuditLogEntry "Global Settings Change", "Global Settings Change", "Major", 1, GetTerm("Routing") & " - Send when the Next Stop has not been set for X minutes change from " & NextStopNagMinutes_ORIG & " to " & NextStopNagMinutes
	End If
	
	If cint(NextStopNagMessageMaxToSendPerDriverPerDay) <> cint(NextStopNagMessageMaxToSendPerDriverPerDay_ORIG) Then
		CreateAuditLogEntry "Global Settings Change", "Global Settings Change", "Major", 1, GetTerm("Routing") & " - Send a maximum of X Next Stop messages to any driver on a given day changed from " & NextStopNagMessageMaxToSendPerDriverPerDay_ORIG & " to " & NextStopNagMessageMaxToSendPerDriverPerDay
	End If
	
	If cint(NextStopNagMessageMaxToSendPerStop) <> cint(NextStopNagMessageMaxToSendPerStop_ORIG) Then
		CreateAuditLogEntry "Global Settings Change", "Global Settings Change", "Major", 1, GetTerm("Routing") & " - Send a maximum of X Next Stop messages each time a No Next Stop event occurs changed from " & NextStopNagMessageMaxToSendPerStop_ORIG & " to " & NextStopNagMessageMaxToSendPerStop
	End If

	If cint(NextStopNagIntervalMinutes) <> cint(NextStopNagIntervalMinutes_ORIG) Then
		CreateAuditLogEntry "Global Settings Change", "Global Settings Change", "Major", 1, GetTerm("Routing") & " - Continue to send Next Stop nag alerts every X minutes changed from " & NextStopNagIntervalMinutes_ORIG & " to " & NextStopNagIntervalMinutes
	End If
			

	'***********************************************************
	'Update SQL with Request Form Field Data
	'***********************************************************	

	SQL = "UPDATE  " & MUV_Read("SQL_Owner") & ".Settings_Global SET  "	
	SQL = SQL & "DelBoardScheduledColor = '" & DelBoardScheduledColor & "',"
	SQL = SQL & "DelBoardCompletedColor = '" & DelBoardCompletedColor & "',"
	SQL = SQL & "DelBoardInProgressColor = '" & DelBoardInProgressColor & "',"
	SQL = SQL & "DelBoardSkippedColor = '" & DelBoardSkippedColor & "',"
	SQL = SQL & "DelBoardNextStopColor= '" & DelBoardNextStopColor & "',"
	SQL = SQL & "DelBoardAMColor = '" & DelBoardAMColor & "',"
	SQL = SQL & "DelBoardPriorityColor = '" & DelBoardPriorityColor & "',"
	SQL = SQL & "DelBoardPieTimerColor = '" & DelBoardPieTimerColor & "',"
	SQL = SQL & "DelBoardTitleGradientColor = '" & DelBoardTitleGradientColor & "',"
	SQL = SQL & "DelBoardTitleText = '" & DelBoardTitleText & "',"
	SQL = SQL & "DelBoardTitleTextFontColor = '" & DelBoardTitleTextFontColor & "',"		
	SQL = SQL & "DelBoardProfitDollars = '" & DelBoardProfitDollars & "',"
	SQL = SQL & "DelBoardAtOrAboveProfitColor = '" & DelBoardAtOrAboveProfitColor & "',"
	SQL = SQL & "DelBoardBelowProfitColor = '" & DelBoardBelowProfitColor & "',"
	SQL = SQL & "DelBoardUserAlertColor = '" & DelBoardUserAlertColor & "',"
	SQL = SQL & "DoNotShowDeliveryLineItems = " & DoNotShowDeliveryLineItems & ","	
	SQL = SQL & "AutoPromptNextStop = " & AutoPromptNextStop & ","
	SQL = SQL & "AutoForceSelectNextStop = " & AutoForceSelectNextStop & ","
	SQL = SQL & "DelBoardDontUseStopSequencing = " & DelBoardDontUseStopSequence& ","
	SQL = SQL & "DelBoardRoutesToIgnore = '" & DelBoardRoutesToIgnore & "',"
	SQL = SQL & "DelBoardUPSRoutes = '" & DelBoardUPSRoutes & "',"

	If MUV_Read("routingModuleOn") = "Enabled" Then	
	
		SQL = SQL & "NextStopNagMessageONOFF = " & NextStopNagMessageONOFF & ","
		SQL = SQL & "NextStopNagMinutes = " & NextStopNagMinutes & ","
		SQL = SQL & "NextStopNagIntervalMinutes = " & NextStopNagIntervalMinutes & ","
		SQL = SQL & "NextStopNagMessageMaxToSendPerStop = " & NextStopNagMessageMaxToSendPerStop & ","
		SQL = SQL & "NextStopNagMessageMaxToSendPerDriverPerDay = " & NextStopNagMessageMaxToSendPerDriverPerDay & ","
		SQL = SQL & "NextStopNagMessageSendMethod = '" & NextStopNagMessageSendMethod & "',"
		SQL = SQL & "NoActivityNagMessageONOFF = " & NoActivityNagMessageONOFF & ","
		SQL = SQL & "NoActivityNagMinutes = " & NoActivityNagMinutes & ","
		SQL = SQL & "NoActivityNagIntervalMinutes = " & NoActivityNagIntervalMinutes & ","
		SQL = SQL & "NoActivityNagMessageMaxToSendPerStop = " & NoActivityNagMessageMaxToSendPerStop & ","
		SQL = SQL & "NoActivityNagMessageMaxToSendPerDriverPerDay = " & NoActivityNagMessageMaxToSendPerDriverPerDay & ","
		SQL = SQL & "NoActivityNagMessageSendMethod = '" & NoActivityNagMessageSendMethod & "',"
		SQL = SQL & "NoActivityNagTimeOfDay = '" & NoActivityNagTimeOfDay & "',"
	End If
		

	If Right(SQL,1) = "," Then SQL = Left(SQL,Len(SQL)-1) ' Strip trailing comma	
									

	Set cnn8 = Server.CreateObject("ADODB.Connection")
	cnn8.open (Session("ClientCnnString"))
	Set rs = Server.CreateObject("ADODB.Recordset")
	rs.CursorLocation = 3 
	Set rs = cnn8.Execute(SQL)
	
	set rs = Nothing
	cnn8.close
	set cnn8 = Nothing



	Response.Redirect("delivery-board.asp")
%>
<!--#include file="../../../inc/footer-main.asp"-->