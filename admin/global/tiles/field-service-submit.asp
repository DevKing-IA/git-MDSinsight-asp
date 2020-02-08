<!--#include file="../../../inc/header.asp"-->

<%

	'***********************************************************
	'Get Values Of All Form Fields Posted
	'***********************************************************
	
	ServiceDayStartTime = Request.Form("txtServiceDayStartTime")
	ServiceDayEndTime = Request.Form("txtServiceDayEndTime")
	ServiceDayElapsedTimeCalculationMethod = Request.Form("optCalcElapsedTime")
	
	If Request.Form("chkServiceTicketCarryoverReportOnOff") = "on" then ServiceTicketCarryoverReportOnOff = 1 Else ServiceTicketCarryoverReportOnOff = 0
	If Request.Form("chkServiceTicketCarryoverReportToPrimarySalesman") = "on" then ServiceTicketCarryoverReportToPrimarySalesman = 1 Else ServiceTicketCarryoverReportToPrimarySalesman = 0
	If Request.Form("chkServiceTicketCarryoverReportToSecondarySalesman") = "on" then ServiceTicketCarryoverReportToSecondarySalesman = 1 Else ServiceTicketCarryoverReportToSecondarySalesman = 0
	If Request.Form("chkCarryoverReportInclCustType") = "on" then CarryoverReportInclCustType = 1 Else CarryoverReportInclCustType = 0
	If Request.Form("chkCarryoverReportInclTicketNum") = "on" then CarryoverReportInclTicketNum = 1 Else CarryoverReportInclTicketNum = 0
	If Request.Form("chkCarryoverReportShowRedoBreakdown") = "on" then CarryoverReportShowRedoBreakdown = 1 Else CarryoverReportShowRedoBreakdown = 0	
 	If Request.Form("chkCarryoverReportIncludeRegions") = "on" then ServiceTicketCarryoverReportIncludeRegions = 1 Else ServiceTicketCarryoverReportIncludeRegions = 0

	ServiceTicketCarryoverReportEmailSubject = Request.Form("txtServiceTicketCarryoverReportEmailSubject")
 	
 	If Request.Form("chkServiceTicketCarryoverReportTextSummaryOnOff") = "on" then ServiceTicketCarryoverReportTextSummaryOnOff = 1 Else ServiceTicketCarryoverReportTextSummaryOnOff = 0
 	
 	If Request.Form("chkServiceTicketthresholdReportONOFF") = "on" then ServiceTicketthresholdReportONOFF = 1 Else ServiceTicketthresholdReportONOFF = 0
 	If Request.Form("chkServiceTicketthresholdReportOnlyUndispatched") = "on" then ServiceTicketthresholdReportOnlyUndispatched = 1 Else ServiceTicketthresholdReportOnlyUndispatched = 0
 	If Request.Form("chkServiceTicketthresholdReportOnlySkipFilterChanges") = "on" then ServiceTicketthresholdReportOnlySkipFilterChanges = 1 Else ServiceTicketthresholdReportOnlySkipFilterChanges = 0
 	ServiceTicketthresholdReportthresholdHours= Request.Form("selServiceTicketthresholdReportthresholdHours")

	DLinkInEmail = Request.Form("chkDLinkInEmail")
	If DLinkInEmail = "on" Then DLinkInEmail = 1 Else DLinkInEmail = 0
	DLinkInText = Request.Form("chkDLinkInText")
	If DLinkInText = "on" Then DLinkInText = 1 Else DLinkInText = 0
	FS_SignatureOptional= Request.Form("chkFSSignatureOptional")
	If FS_SignatureOptional = "on" Then FS_SignatureOptional = 1 Else FS_SignatureOptional = 0
	FS_TechCanDecline = Request.Form("chkFSTechCanDecline")
	FS_ShowPartsButton = Request.Form("chkFSShowPartsButton")
	ServiceTicketScreenShowHoldTab = Request.Form("chkServiceTicketScreenShowHoldTab")
	If FS_TechCanDecline = "on" Then FS_TechCanDecline = 1 Else FS_TechCanDecline = 0
	If FS_ShowPartsButton = "on" Then FS_ShowPartsButton = 1 Else FS_ShowPartsButton = 0
	If ServiceTicketScreenShowHoldTab = "on" Then ServiceTicketScreenShowHoldTab  = 1 Else ServiceTicketScreenShowHoldTab  = 0
	AutoDispatchUsersOnOff = Request.Form("chkAutoDispatchUsersOnOff")
	If AutoDispatchUsersOnOff = "on" Then AutoDispatchUsersOnOff  = 1 Else AutoDispatchUsersOnOff  = 0	
	AutoDispatchUserNos = Request.Form("selAutoDispatchUserNos")
	
	FSDefaultNotificationMethod = Request.Form("selFSDefaultNotificationMethod") 
	
	ServiceColorsOn = Request.Form("chkServiceColorsOn")
	If ServiceColorsOn ="on" Then ServiceColorsOn = 1 Else ServiceColorsOn = 0
	ServiceNormalAlertColor = Request.Form("txtNormalAlert")
	ServicePriorityColor = Request.Form("txtPriorityAccount")
	ServicePriorityAlertColor = Request.Form("txtPriorityAccountAlert")
	
	If Request.Form("chkFSBoardKioskGlobalUseRegions") = "on" then FSBoardKioskGlobalUseRegions = 1 Else FSBoardKioskGlobalUseRegions = 0
	
	FSBoardKioskGlobalTitleGradientColor = Request.Form("txtFSBoardKioskGlobalTitleGradientColor")
	FSBoardKioskGlobalTitleText = Request.Form("txtFSBoardKioskGlobalTitleText")
	FSBoardKioskGlobalTitleTextFontColor = Request.Form("txtFSBoardKioskGlobalTitleTextFontColor")
	FSBoardKioskGlobalColorPieTimer = Request.Form("txtFSBoardKioskGlobalColorPieTimer")
	FSBoardKioskGlobalColorUrgent = Request.Form("txtFSBoardKioskGlobalColorUrgent")
	FSBoardKioskGlobalColorAwaitingDispatch = Request.Form("txtFSBoardKioskGlobalColorAwaitingDispatch")
	FSBoardKioskGlobalColorAwaitingAcknowledgement = Request.Form("txtFSBoardKioskGlobalColorAwaitingAcknowledgement")
	FSBoardKioskGlobalColorDispatchAcknowledged = Request.Form("txtFSBoardKioskGlobalColorDispatchAcknowledged")
	FSBoardKioskGlobalColorDispatchDeclined = Request.Form("txtFSBoardKioskGlobalColorDispatchDeclined")
	FSBoardKioskGlobalColorEnRoute = Request.Form("txtFSBoardKioskGlobalColorEnRoute")
	FSBoardKioskGlobalColorOnSite = Request.Form("txtFSBoardKioskGlobalColorOnSite")
	FSBoardKioskGlobalColorClosed = Request.Form("txtFSBoardKioskGlobalColorClosed")
	FSBoardKioskGlobalColorRedoSwap = Request.Form("txtFSBoardKioskGlobalColorRedoSwap")
	FSBoardKioskGlobalColorRedoWaitForParts = Request.Form("txtFSBoardKioskGlobalColorRedoWaitForParts")
	FSBoardKioskGlobalColorRedoFollowUp = Request.Form("txtFSBoardKioskGlobalColorRedoFollowUp")
	FSBoardKioskGlobalColorRedoUnableToWork = Request.Form("txtFSBoardKioskGlobalColorRedoUnableToWork")

	If Request.Form("chkFieldServiceNotesReportOnOff") = "on" then FieldServiceNotesReportOnOff = 1 Else FieldServiceNotesReportOnOff = 0
	FieldServiceNotesReportEmailSubject = Request.Form("txtFieldServiceNotesReportEmailSubject")
		
	If Request.Form("chkNoActivityNagMessageONOFF_FS") = "on" then NoActivityNagMessageONOFF_FS = 1 Else NoActivityNagMessageONOFF_FS = 0
	NoActivityNagMinutes_FS = Request.Form("selNoActivityNagMinutes_FS")
	NoActivityNagIntervalMinutes_FS = Request.Form("selNoActivityNagIntervalMinutes_FS")
	NoActivityNagMessageMaxToSendPerStop_FS = Request.Form("selNoActivityNagMessageMaxToSendPerStop_FS")
	NoActivityNagMessageMaxToSendPerDriverPerDay_FS = Request.Form("selNoActivityNagMessageMaxToSendPerDriverPerDay_FS")
	NoActivityNagMessageSendMethod_FS = Request.Form("selNoActivityNagMessageSendMethod_FS")
	NoActivityNagTimeOfDay_FS = Request.Form("selNoActivityNagTimeOfDay_FS")

 	FieldServiceNotesReportUserNos = Request.Form("lstSelectedFieldServiceNotesReportUserIDs")
	ServiceTicketCarryoverReportUserNos = Request.Form("lstSelectedServiceTicketCarryoverReportUserIDs")
	ServiceTicketCarryoverReportTextSummaryUserNos = Request.Form("lstSelectedServiceTicketCarryoverReportTextSummmaryUserIDs")
    ServiceTicketthresholdReportUserNos = Request.Form("lstSelectedServiceTicketthresholdReportUserNos")
	ServiceTicketCarryoverReportTeamIntRecIDs = Request.Form("lstSelectedServiceTicketCarryoverReportTeamIntRecIDs")

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
		
		ServiceColorsOn_ORIG = rs("ServiceColorsOn")
		ServiceNormalAlertColor_ORIG = rs("ServiceNormalAlertColor")
		ServicePriorityColor_ORIG = rs("ServicePriorityColor")
		ServicePriorityAlertColor_ORIG = rs("ServicePriorityAlertColor")			
		NoActivityNagMessageONOFF_FS_ORIG = rs("NoActivityNagMessageONOFF_FS")
		NoActivityNagMinutes_FS_ORIG = rs("NoActivityNagMinutes_FS")
		NoActivityNagIntervalMinutes_FS_ORIG = rs("NoActivityNagIntervalMinutes_FS")
		NoActivityNagMessageMaxToSendPerStop_FS_ORIG = rs("NoActivityNagMessageMaxToSendPerStop_FS")
		NoActivityNagMessageMaxToSendPerDriverPerDay_FS_ORIG = rs("NoActivityNagMessageMaxToSendPerDriverPerDay_FS")
		NoActivityNagMessageSendMethod_FS_ORIG = rs("NoActivityNagMessageSendMethod_FS")
		NoActivityNagTimeOfDay_FS_ORIG = rs("NoActivityNagTimeOfDay_FS")  		
		FS_SignatureOptional_ORIG = rs("FS_SignatureOptional")
		FS_TechCanDecline_ORIG = rs("FS_TechCanDecline")
		FSDefaultNotificationMethod_ORIG = rs("FSDefaultNotificationMethod")
		
	End If
	
	set rs = Nothing
	cnn8.close
	set cnn8 = Nothing

	'************************************************
	'The newest settings are in Settings_FieldService
	'************************************************
	SQL = "SELECT * FROM Settings_FieldService "
	
	Set cnn8 = Server.CreateObject("ADODB.Connection")
	cnn8.open (Session("ClientCnnString"))
	Set rs = Server.CreateObject("ADODB.Recordset")
	rs.CursorLocation = 3 
	Set rs = cnn8.Execute(SQL)
	
	If not rs.EOF Then
		ServiceTicketCarryoverReportOnOff_ORIG = rs("ServiceTicketCarryoverReportOnOff")
		ServiceTicketCarryoverReportToPrimarySalesman_ORIG = rs("ServiceTicketCarryoverReportToPrimarySalesman")
		ServiceTicketCarryoverReportToSecondarySalesman_ORIG = rs("ServiceTicketCarryoverReportToSecondarySalesman")
		ServiceTicketCarryoverReportEmailSubject_ORIG = rs("ServiceTicketCarryoverReportEmailSubject")
		ServiceTicketCarryoverReportUserNos_ORIG = rs("ServiceTicketCarryoverReportUserNos")		
		ServiceTicketCarryoverReportAdditionalEmails_ORIG = rs("ServiceTicketCarryoverReportAdditionalEmails")	
		ServiceTicketCarryoverReportTextSummaryOnOff_ORIG = rs("ServiceTicketCarryoverReportTextSummaryOnOff")
		ServiceTicketCarryoverReportTextSummaryUserNos_ORIG = rs("ServiceTicketCarryoverReportTextSummaryUserNos")
		ServiceTicketCarryoverReportTeamIntRecIDs_ORIG = rs("ServiceTicketCarryoverReportTeamIntRecIDs")
		ServiceTicketCarryoverReportIncludeRegions_ORIG = rs("ServiceTicketCarryoverReportIncludeRegions")
		ServiceDayStartTime_ORIG = rs("ServiceDayStartTime")
		ServiceDayEndTime_ORIG = rs("ServiceDayEndTime")
		ServiceDayElapsedTimeCalculationMethod_ORIG = rs("ServiceDayElapsedTimeCalculationMethod")
		FieldServiceNotesReportOnOff_ORIG = rs("FieldServiceNotesReportOnOff")
		FieldServiceNotesReportUserNos_ORIG = rs("FieldServiceNotesReportUserNos")
		FieldServiceNotesReportAdditionalEmails_ORIG = rs("FieldServiceNotesReportAdditionalEmails")
		FieldServiceNotesReportEmailSubject_ORIG = rs("FieldServiceNotesReportEmailSubject")	
		AutoDispatchUsersOnOff_ORIG  = rs("AutoDispatchUsersOnOff")	
		AutoDispatchUserNos_ORIG  = rs("AutoDispatchUserNos")	
		CarryoverReportInclCustType_ORIG  = rs("CarryoverReportInclCustType")	
		FS_ShowPartsButton_ORIG = rs("ShowPartsButton")
		ServiceTicketScreenShowHoldTab_ORIG = rs("ServiceTicketScreenShowHoldTab")
		CarryoverReportInclTicketNum_ORIG  = rs("CarryoverReportInclTicketNum")	
		CarryoverReportShowRedoBreakdown_ORIG  = rs("CarryoverReportShowRedoBreakdown")		
		ServiceTicketthresholdReportONOFF_Orig = rs("ServiceTicketthresholdReportONOFF")
		ServiceTicketthresholdReportOnlyUndispatched_Orig = rs("ServiceTicketthresholdReportOnlyUndispatched")
		ServiceTicketthresholdReportOnlySkipFilterChanges_Orig = rs("ServiceTicketthresholdReportOnlySkipFilterChanges")
		ServiceTicketthresholdReportthresholdHours_Orig = rs("ServiceTicketthresholdReportthresholdHours")
		ServiceTicketthresholdReportUserNos_Orig = rs("ServiceTicketthresholdReportUserNos")
		ServiceTicketthresholdReportAdditionalEmails_Orig = rs("ServiceTicketthresholdReportAdditionalEmails")
		FSBoardKioskGlobalUseRegions_Orig = rs("FSBoardKioskGlobalUseRegions")
		FSBoardKioskGlobalTitleGradientColor_Orig = rs("FSBoardKioskGlobalTitleGradientColor")
		FSBoardKioskGlobalTitleText_Orig = rs("FSBoardKioskGlobalTitleText")
		FSBoardKioskGlobalTitleTextFontColor_Orig = rs("FSBoardKioskGlobalTitleTextFontColor")
		FSBoardKioskGlobalColorPieTimer_Orig = rs("FSBoardKioskGlobalColorPieTimer")
		FSBoardKioskGlobalColorUrgent_Orig = rs("FSBoardKioskGlobalColorUrgent")
		FSBoardKioskGlobalColorAwaitingDispatch_Orig = rs("FSBoardKioskGlobalColorAwaitingDispatch")
		FSBoardKioskGlobalColorAwaitingAcknowledgement_Orig = rs("FSBoardKioskGlobalColorAwaitingAcknowledgement")
		FSBoardKioskGlobalColorDispatchAcknowledged_Orig = rs("FSBoardKioskGlobalColorDispatchAcknowledged")
		FSBoardKioskGlobalColorDispatchDeclined_Orig = rs("FSBoardKioskGlobalColorDispatchDeclined")
		FSBoardKioskGlobalColorEnRoute_Orig = rs("FSBoardKioskGlobalColorEnRoute")
		FSBoardKioskGlobalColorOnSite_Orig = rs("FSBoardKioskGlobalColorOnSite")
		FSBoardKioskGlobalColorClosed_Orig = rs("FSBoardKioskGlobalColorClosed")
		FSBoardKioskGlobalColorRedoSwap_Orig = rs("FSBoardKioskGlobalColorRedoSwap")
		FSBoardKioskGlobalColorRedoWaitForParts_Orig = rs("FSBoardKioskGlobalColorRedoWaitForParts")
		FSBoardKioskGlobalColorRedoFollowUp_Orig = rs("FSBoardKioskGlobalColorRedoFollowUp")
		FSBoardKioskGlobalColorRedoUnableToWork_Orig = rs("FSBoardKioskGlobalColorRedoUnableToWork")
		
	End If
		
	set rs = Nothing
	cnn8.close
	set cnn8 = Nothing
	'***********************************************************
	'Perform Audit Trail Entry Inserts
	'***********************************************************

    If FieldServiceNotesReportUserNos <> FieldServiceNotesReportUserNos_ORIG Then
		CreateAuditLogEntry "Field Service Settings Change", "Field Service Settings Change", "Major", 1, "FieldServiceNotesReportUserNos changed from " & FieldServiceNotesReportUserNos_ORIG & " to " & FieldServiceNotesReportUserNos
	End If

    If ServiceTicketCarryoverReportUserNos <> ServiceTicketCarryoverReportUserNos_ORIG Then
		CreateAuditLogEntry "Field Service Settings Change", "Field Service Settings Change", "Major", 1, "ServiceTicketCarryoverReportUserNos changed from " & ServiceTicketCarryoverReportUserNos_ORIG & " to " & ServiceTicketCarryoverReportUserNos
	End If

    If ServiceTicketthresholdReportUserNos <> ServiceTicketthresholdReportUserNos_ORIG Then
		CreateAuditLogEntry "Field Service Settings Change", "Field Service Settings Change", "Major", 1, "ServiceTicketthresholdReportUserNos changed from " & ServiceTicketthresholdReportUserNos_ORIG & " to " & ServiceTicketthresholdReportUserNos
	End If
	
    If ServiceTicketCarryoverReportTextSummaryUserNos <> ServiceTicketCarryoverReportTextSummaryUserNos_ORIG Then
		CreateAuditLogEntry "Field Service Settings Change", "Field Service Settings Change", "Major", 1, "ServiceTicketCarryoverReportTextSummaryUserNos changed from " & ServiceTicketCarryoverReportTextSummaryUserNos_ORIG & " to " & ServiceTicketCarryoverReportTextSummaryUserNos
	End If

    If ServiceTicketCarryoverReportTeamIntRecIDs <> ServiceTicketCarryoverReportTeamIntRecIDs_ORIG Then
		CreateAuditLogEntry "Field Service Settings Change", "Field Service Settings Change", "Major", 1, "ServiceTicketCarryoverReportTeamIntRecIDs changed from " & ServiceTicketCarryoverReportTeamIntRecIDs_ORIG & " to " & ServiceTicketCarryoverReportTeamIntRecIDs
	End If

	If ServiceDayStartTime <> ServiceDayStartTime_ORIG Then 
		AuditMessage = GetTerm("Field Service") & " service day start time changed from " & ServiceDayStartTime_ORIG & " to " & ServiceDayStartTime			
		CreateAuditLogEntry GetTerm("Field Service") & " service day start time change", GetTerm("Field Service") & " service day start time change", "Major", 1, AuditMessage 	
	End If
	
	If ServiceDayEndTime <> ServiceDayEndTime_ORIG Then 
		AuditMessage = GetTerm("Field Service") & " service day end time changed from " & ServiceDayEndTime_ORIG & " to " & ServiceDayEndTime			
		CreateAuditLogEntry GetTerm("Field Service") & " service day end time change", GetTerm("Field Service") & " service day end time change", "Major", 1, AuditMessage 	
	End If

	If ServiceDayElapsedTimeCalculationMethod <> ServiceDayElapsedTimeCalculationMethod_ORIG Then 
		AuditMessage = GetTerm("Field Service") & " Elapsed time calculation method changed from " & ServiceDayElapsedTimeCalculationMethod_ORIG & " to " & ServiceDayElapsedTimeCalculationMethod			
		CreateAuditLogEntry GetTerm("Field Service") & " elapsed time calculation method change", GetTerm("Field Service") & " elapsed time calculation method change", "Major", 1, AuditMessage 	
	End If

	If ServiceColorsOn <> ServiceColorsOn_ORIG Then 
		AuditMessage = GetTerm("Field Service") & " screen highlight colors on changed from " & ServiceColorsOn_ORIG & " to " & ServiceColorsOn 			
		CreateAuditLogEntry GetTerm("Field Service") & " screen highlight color change", GetTerm("Field Service") & " screen highlight color change", "Major", 1, AuditMessage 	
	End If
	
	If ServiceNormalAlertColor <> ServiceNormalAlertColor_ORIG Then 
		AuditMessage = GetTerm("Field Service") & " screen highlight color for Normal customer - Alert Sent changed from " & ServiceNormalAlertColor_ORIG & " to " & ServiceNormalAlertColor 			
		CreateAuditLogEntry GetTerm("Field Service") & " screen highlight color change", GetTerm("Field Service") & " screen highlight color change", "Major", 1, AuditMessage 	
	End If
	
	If ServicePriorityColor <> ServicePriorityColor_ORIG Then 
		AuditMessage = GetTerm("Field Service") & " screen highlight color for Priority customer changed from " & ServicePriorityColor_ORIG & " to " & ServicePriorityColor			
		CreateAuditLogEntry GetTerm("Field Service") & " screen highlight color change", GetTerm("Field Service") & " screen highlight color change", "Major", 1, AuditMessage 	
	End If
	
	If ServicePriorityAlertColor <> ServicePriorityAlertColor_ORIG Then 
		AuditMessage = GetTerm("Field Service") & "screen highlight color for Priority customer - Alert Sent changed from " & ServicePriorityAlertColor_ORIG & " to " & ServicePriorityAlertColor			
		CreateAuditLogEntry GetTerm("Field Service") & " screen highlight color change", GetTerm("Field Service") & " screen highlight color change", "Major", 1, AuditMessage 	
	End If

	If cInt(FSBoardKioskGlobalUseRegions) = 1 then FSBoardKioskGlobalUseRegions_Msg = "On" Else FSBoardKioskGlobalUseRegions_Msg = "Off"
	
	If IsNumeric(FSBoardKioskGlobalUseRegions_ORIG) Then 
		If cInt(FSBoardKioskGlobalUseRegions) <> cInt(FSBoardKioskGlobalUseRegions_ORIG) Then
			CreateAuditLogEntry "Global Settings Change", "Global Settings Change", "Major", 1, GetTerm("Field Service") & " - kiosk use regions changed from " & FSBoardKioskGlobalUseRegions_ORIG & " to " & FSBoardKioskGlobalUseRegions_Msg
		End If
	End If

	If FSBoardKioskGlobalTitleGradientColor <> FSBoardKioskGlobalTitleGradientColor_ORIG Then 
		AuditMessage = GetTerm("Field Service") & " kiosk screen title gradient color changed from " & FSBoardKioskGlobalTitleGradientColor_ORIG & " to " & FSBoardKioskGlobalTitleGradientColor			
		CreateAuditLogEntry GetTerm("Field Service") & "  kiosk screen title gradient color change", GetTerm("Field Service") & " kiosk screen title gradient color change", "Major", 1, AuditMessage 	
	End If
	
	If FSBoardKioskGlobalTitleText <> FSBoardKioskGlobalTitleText_ORIG Then 
		AuditMessage = GetTerm("Field Service") & " kiosk screen title text changed from " & FSBoardKioskGlobalTitleText_ORIG & " to " & FSBoardKioskGlobalTitleText			
		CreateAuditLogEntry GetTerm("Field Service") & " kiosk screen title text change", GetTerm("Field Service") & " kiosk screen title text change", "Major", 1, AuditMessage 	
	End If
	
	If FSBoardKioskGlobalTitleTextFontColor <> FSBoardKioskGlobalTitleTextFontColor_ORIG Then 
		AuditMessage = GetTerm("Field Service") & " kiosk screen title text font color changed from " & FSBoardKioskGlobalTitleTextFontColor_ORIG & " to " & FSBoardKioskGlobalTitleTextFontColor			
		CreateAuditLogEntry GetTerm("Field Service") & " kiosk screen title text font color change", GetTerm("Field Service") & " kiosk screen title text font color change", "Major", 1, AuditMessage 	
	End If
	
	If FSBoardKioskGlobalColorPieTimer <> FSBoardKioskGlobalColorPieTimer_ORIG Then 
		AuditMessage = GetTerm("Field Service") & " screen pie timer color changed from " & FSBoardKioskGlobalColorPieTimer_ORIG & " to " & FSBoardKioskGlobalColorPieTimer			
		CreateAuditLogEntry GetTerm("Field Service") & " kiosk screen highlight color change", GetTerm("Field Service") & " kiosk screen highlight color change", "Major", 1, AuditMessage 	
	End If

	If FSBoardKioskGlobalColorUrgent <> FSBoardKioskGlobalColorUrgent_ORIG Then 
		AuditMessage = GetTerm("Field Service") & " kiosk screen highlight color for Urgent Service Ticket changed from " & FSBoardKioskGlobalColorUrgent_ORIG & " to " & FSBoardKioskGlobalColorUrgent			
		CreateAuditLogEntry GetTerm("Field Service") & " kiosk screen highlight color change", GetTerm("Field Service") & " kiosk screen highlight color change", "Major", 1, AuditMessage 	
	End If

	If FSBoardKioskGlobalColorAwaitingDispatch <> FSBoardKioskGlobalColorAwaitingDispatch_ORIG Then 
		AuditMessage = GetTerm("Field Service") & " kiosk screen highlight color for Awaiting Dispatch Service Ticket changed from " & FSBoardKioskGlobalColorAwaitingDispatch_ORIG & " to " & FSBoardKioskGlobalColorAwaitingDispatch			
		CreateAuditLogEntry GetTerm("Field Service") & " kiosk screen highlight color change", GetTerm("Field Service") & " kiosk screen highlight color change", "Major", 1, AuditMessage 	
	End If

	If FSBoardKioskGlobalColorAwaitingAcknowledgement <> FSBoardKioskGlobalColorAwaitingAcknowledgement_ORIG Then 
		AuditMessage = GetTerm("Field Service") & " Ticket Screen Status Label Color for Awaiting Acknowledgement changed from " & FSBoardKioskGlobalColorAwaitingAcknowledgement_ORIG & " to " & FSBoardKioskGlobalColorAwaitingAcknowledgement			
		CreateAuditLogEntry GetTerm("Field Service") & " Ticket Screen Status Label Color Change", GetTerm("Field Service") & " Ticket Screen Status Label Color Change for Awaiting Acknowledgement", "Major", 1, AuditMessage 	
	End If
	
	If FSBoardKioskGlobalColorDispatchAcknowledged <> FSBoardKioskGlobalColorDispatchAcknowledged_ORIG Then 
		AuditMessage = GetTerm("Field Service") & " kiosk screen highlight color for Acknowledged Service Ticket changed from " & FSBoardKioskGlobalColorDispatchAcknowledged_ORIG & " to " & FSBoardKioskGlobalColorDispatchAcknowledged			
		CreateAuditLogEntry GetTerm("Field Service") & " kiosk screen highlight color change", GetTerm("Field Service") & " kiosk screen highlight color change", "Major", 1, AuditMessage 	
	End If
	
	If FSBoardKioskGlobalColorDispatchDeclined <> FSBoardKioskGlobalColorDispatchDeclined_ORIG Then 
		AuditMessage = GetTerm("Field Service") & " kiosk screen highlight color for Declined Dispatches changed from " & FSBoardKioskGlobalColorDispatchDeclined_ORIG & " to " & FSBoardKioskGlobalColorDispatchDeclined			
		CreateAuditLogEntry GetTerm("Field Service") & " kiosk screen highlight color change", GetTerm("Field Service") & " kiosk screen declined highlight color change", "Major", 1, AuditMessage 	
	End If
	
	If FSBoardKioskGlobalColorEnRoute <> FSBoardKioskGlobalColorEnRoute_ORIG Then 
		AuditMessage = GetTerm("Field Service") & " kiosk screen highlight color for En Route Service Ticket changed from " & FSBoardKioskGlobalColorEnRoute_ORIG & " to " & FSBoardKioskGlobalColorEnRoute			
		CreateAuditLogEntry GetTerm("Field Service") & " kiosk screen highlight color change", GetTerm("Field Service") & " kiosk screen highlight color change", "Major", 1, AuditMessage 	
	End If
	
	If FSBoardKioskGlobalColorOnSite <> FSBoardKioskGlobalColorOnSite_ORIG Then 
		AuditMessage = GetTerm("Field Service") & " kiosk screen highlight color for On Site Service Ticket changed from " & FSBoardKioskGlobalColorOnSite_ORIG & " to " & FSBoardKioskGlobalColorOnSite			
		CreateAuditLogEntry GetTerm("Field Service") & " kiosk screen highlight color change", GetTerm("Field Service") & " kiosk screen highlight color change", "Major", 1, AuditMessage 	
	End If

	If FSBoardKioskGlobalColorRedoSwap <> FSBoardKioskGlobalColorRedoSwap_ORIG Then 
		AuditMessage = GetTerm("Field Service") & " Ticket Screen Status Label Color for Swap changed from " & FSBoardKioskGlobalColorRedoSwap_ORIG & " to " & FSBoardKioskGlobalColorRedoSwap			
		CreateAuditLogEntry GetTerm("Field Service") & " Ticket Screen Status Label Color Change", GetTerm("Field Service") & " Ticket Screen Status Label Color Change for Swap (Redo)", "Major", 1, AuditMessage 	
	End If
	
	If FSBoardKioskGlobalColorRedoWaitForParts <> FSBoardKioskGlobalColorRedoWaitForParts_ORIG Then 
		AuditMessage = GetTerm("Field Service") & " Ticket Screen Status Label Color for Wait For Parts changed from " & FSBoardKioskGlobalColorRedoWaitForParts_ORIG & " to " & FSBoardKioskGlobalColorRedoWaitForParts			
		CreateAuditLogEntry GetTerm("Field Service") & " Ticket Screen Status Label Color Change", GetTerm("Field Service") & " Ticket Screen Status Label Color Change for Wait For Parts (Redo)", "Major", 1, AuditMessage 	
	End If
	
	If FSBoardKioskGlobalColorRedoFollowUp <> FSBoardKioskGlobalColorRedoFollowUp_ORIG Then 
		AuditMessage = GetTerm("Field Service") & " Ticket Screen Status Label Color for Follow Up changed from " & FSBoardKioskGlobalColorRedoFollowUp_ORIG & " to " & FSBoardKioskGlobalColorRedoFollowUp			
		CreateAuditLogEntry GetTerm("Field Service") & " Ticket Screen Status Label Color Change", GetTerm("Field Service") & " Ticket Screen Status Label Color Change for Follow Up (Redo)", "Major", 1, AuditMessage 	
	End If
	
	If FSBoardKioskGlobalColorRedoUnableToWork <> FSBoardKioskGlobalColorRedoUnableToWork_ORIG Then 
		AuditMessage = GetTerm("Field Service") & " Ticket Screen Status Label Color for Unable To Work changed from " & FSBoardKioskGlobalColorRedoUnableToWork_ORIG & " to " & FSBoardKioskGlobalColorRedoUnableToWork			
		CreateAuditLogEntry GetTerm("Field Service") & " Ticket Screen Status Label Color Change", GetTerm("Field Service") & " Ticket Screen Status Label Color Change for Unable To Work (Redo)", "Major", 1, AuditMessage 	
	End If

	
	If FSBoardKioskGlobalColorClosed <> FSBoardKioskGlobalColorClosed_ORIG Then 
		AuditMessage = GetTerm("Field Service") & " kiosk screen highlight color for Closed Service Ticket changed from " & FSBoardKioskGlobalColorClosed_ORIG & " to " & FSBoardKioskGlobalColorClosed			
		CreateAuditLogEntry GetTerm("Field Service") & " kiosk screen highlight color change", GetTerm("Field Service") & " kiosk screen highlight color change", "Major", 1, AuditMessage 	
	End If

	



	If cInt(NoActivityNagMessageONOFF_FS) = 1 then NoActivityNagMessageONOFF_FSMsg = "On" Else NoActivityNagMessageONOFF_FSMsg = "Off"
	
	If IsNumeric(NoActivityNagMessageONOFF_FS_ORIG) Then 
		If cInt(NoActivityNagMessageONOFF_FS) <> cInt(NoActivityNagMessageONOFF_FS_ORIG) Then
			CreateAuditLogEntry "Global Settings Change", "Global Settings Change", "Major", 1, GetTerm("Field Service") & " - send nag messages when there has been a period of No Activity changed from " & NoActivityNagMessageONOFF_FS_ORIG & " to " & NoActivityNagMessageONOFF_FSMsg
		End If
	End If
	
	If NoActivityNagMessageSendMethod_FS <> NoActivityNagMessageSendMethod_FS_ORIG Then
		CreateAuditLogEntry "Global Settings Change", "Global Settings Change", "Major", 1, GetTerm("Field Service") & " - nag send method changed from  " & NoActivityNagMessageSendMethod_FS_ORIG & " to " & NoActivityNagMessageSendMethod_FS
	End If

	If cint(NoActivityNagMinutes_FS) <> cint(NoActivityNagMinutes_FS_ORIG) Then
		CreateAuditLogEntry "Global Settings Change", "Global Settings Change", "Major", 1, GetTerm("Field Service") & " - Send when there has been No Activity for X minutes changed from " & NoActivityNagMinutes_FS_ORIG & " to " & NoActivityNagMinutes_FS
	End If
	
	If cint(NoActivityNagMessageMaxToSendPerDriverPerDay_FS) <> cint(NoActivityNagMessageMaxToSendPerDriverPerDay_FS_ORIG) Then
		CreateAuditLogEntry "Global Settings Change", "Global Settings Change", "Major", 1, GetTerm("Field Service") & " - Send a maximum of X messages to any technician on a given day changed from " & NoActivityNagMessageMaxToSendPerDriverPerDay_FS_ORIG & " to " & NoActivityNagMessageMaxToSendPerDriverPerDay_FS
	End If
	
	If cint(NoActivityNagMessageMaxToSendPerStop_FS) <> cint(NoActivityNagMessageMaxToSendPerStop_FS_ORIG) Then
		CreateAuditLogEntry "Global Settings Change", "Global Settings Change", "Major", 1, GetTerm("Field Service") & " - Send a maximum of X messages each time a No Activity event occurs changed from " & NoActivityNagMessageMaxToSendPerStop_FS_ORIG & " to " & NoActivityNagMessageMaxToSendPerStop_FS
	End If

	If cint(NoActivityNagIntervalMinutes_FS) <> cint(NoActivityNagIntervalMinutes_FS_ORIG) Then
		CreateAuditLogEntry "Global Settings Change", "Global Settings Change", "Major", 1, GetTerm("Field Service") & " - Continue to send nag alerts every X minutes changed from " & NoActivityNagIntervalMinutes_FS_ORIG & " to " & NoActivityNagIntervalMinutes_FS
	End If

	If NoActivityNagTimeOfDay_FS <> NoActivityNagTimeOfDay_FS_ORIG Then
		CreateAuditLogEntry "Global Settings Change", "Global Settings Change", "Major", 1, GetTerm("Field Service") & " - Send nag when there has been No Activity by X time changed from " & NoActivityNagTimeOfDay_FS_ORIG & " to " & NoActivityNagTimeOfDay_FS
	End If
	
	If FSDefaultNotificationMethod <> FSDefaultNotificationMethod_ORIG Then
		CreateAuditLogEntry "Global Settings Change", "Global Settings Change", "Major", 1, GetTerm("Field Service") & " - default notification method changed from " & FSDefaultNotificationMethod_ORIG & " to " & FSDefaultNotificationMethod
	End If
	

	'*********************************************************************************
	' This code to write to audit trail file for SETTINGS_EMAILSERVICE table fields
	'*********************************************************************************
	SQL = "SELECT * FROM Settings_EmailService"
	Set cnn8 = Server.CreateObject("ADODB.Connection")
	cnn8.open (Session("ClientCnnString"))
	Set rs = Server.CreateObject("ADODB.Recordset")
	rs.CursorLocation = 3 
	Set rs = cnn8.Execute(SQL)
	If not rs.EOF Then
		DLinkInEmail4Compare  =  DLinkInEmail
		If DLinkInEmail4Compare <> rs("IncludeACKinDispatchEmail") Then
			' Just make it say On/Off instead of True/False
			VerbiageForReport = rs("IncludeACKinDispatchEmail")
			If DLinkInEmail = 1 then DLinkInEmail4Compare = "True" else DLinkInEmail4Compare = "False" 
			CreateAuditLogEntry "Dispatch Email Settings Change", "Dispatch Email Settings Change", "Major", 1, MUV_Read("DisplayName") & " changed: Include acknowledgement link in dispatch email from " & VerbiageForReport & " to " & DLinkInEmail
		End If
		DLinkInText4Compare  =  DLinkInText
		If DLinkInTextCompare <> rs("IncludeACKinDispatchText") Then
			' Just make it say On/Off instead of True/False
			VerbiageForReport = rs("IncludeACKinDispatchText")
			If DLinkInText = 1 then DLinkInText4Compare = "True" else DLinkInText4Compare = "False" 
			CreateAuditLogEntry "Dispatch Text Settings Change", "Dispatch Text Settings Change", "Major", 1, MUV_Read("DisplayName") & " changed: Include acknowledgement link in dispatch text message from " & VerbiageForReport & " to " & DLinkInText
		End If
		If FS_SignatureOptional_ORIG  <> FS_SignatureOptional Then
			If FS_SignatureOptional_ORIG = 0 Then Verbiage1 = "False" Else Verbiage1 = "True"
			If FS_SignatureOptional = 0 Then Verbiage2 = "False" Else Verbiage2 = "True"
				CreateAuditLogEntry "Dispatch Text Settings Change", "Dispatch Text Settings Change", "Major", 1, MUV_Read("DisplayName") & " changed: FS Signatures optional from " & Verbiage1 & " to " & Verbiage2		
			End If
		End If
		
		If FS_TechCanDecline_ORIG  <> FS_TechCanDecline Then
			If FS_TechCanDecline_ORIG = 0 Then Verbiage1 = "False" Else Verbiage1 = "True"
			If FS_TechCanDecline = 0 Then Verbiage2 = "False" Else Verbiage2 = "True"
			CreateAuditLogEntry "Dispatch Text Settings Change", "Dispatch Text Settings Change", "Major", 1, MUV_Read("DisplayName") & " changed: FS Signatures optional from " & Verbiage1 & " to " & Verbiage2		
		End If

		If FS_ShowPartsButton_ORIG  <> FS_ShowPartsButton Then
			If FS_ShowPartsButton_ORIG = 0 Then Verbiage1 = "False" Else Verbiage1 = "True"
			If FS_ShowPartsButton = 0 Then Verbiage2 = "False" Else Verbiage2 = "True"
			CreateAuditLogEntry "Field Servie Settings Change", "FieldS Service Settings Change", "Major", 1, MUV_Read("DisplayName") & " changed: FS Show Parts Button from " & Verbiage1 & " to " & Verbiage2		
		End If


		If ServiceTicketScreenShowHoldTab_ORIG  <> ServiceTicketScreenShowHoldTab Then
			If ServiceTicketScreenShowHoldTab_ORIG = 0 Then Verbiage1 = "False" Else Verbiage1 = "True"
			If ServiceTicketScreenShowHoldTab = 0 Then Verbiage2 = "False" Else Verbiage2 = "True"
			CreateAuditLogEntry "Field Servie Settings Change", "FieldS Service Settings Change", "Major", 1, MUV_Read("DisplayName") & " changed: service ticket screen show hold tab from " & Verbiage1 & " to " & Verbiage2		
		End If


		If AutoDispatchUsersOnOff_ORIG  <> AutoDispatchUsersOnOff Then
			If AutoDispatchUsersOnOff_ORIG = 0 Then Verbiage1 = "False" Else Verbiage1 = "True"
			If AutoDispatchUsersOnOff = 0 Then Verbiage2 = "False" Else Verbiage2 = "True"
			CreateAuditLogEntry "Field Service Settings Change", "Field Service Settings Change", "Major", 1, MUV_Read("DisplayName") & " changed: Auto dispatch serice tickets from " & Verbiage1 & " to " & Verbiage2		
		End If

		'Fix it - show the actual user names
		If AutoDispatchUserNos_ORIG  <> AutoDispatchUserNos  Then
			CreateAuditLogEntry "Field Service Settings Change", "Field Service Settings Change", "Major", 1, MUV_Read("DisplayName") & " changed: user numbers to auto dispatch optional from " & AutoDispatchUserNos_ORIG & " to " & AutoDispatchUserNos  
		End If



	'***********************************************************
	'Update SQL with Request Form Field Data
	'***********************************************************

	
	SQL = "UPDATE  " & MUV_Read("SQL_Owner") & ".Settings_Global SET "
	SQL = SQL & "ServiceColorsOn = " & ServiceColorsOn & ", "
	SQL = SQL & "ServiceNormalAlertColor = '" & ServiceNormalAlertColor & "', "
	SQL = SQL & "ServicePriorityColor = '" & ServicePriorityColor & "', "
	SQL = SQL & "ServicePriorityAlertColor = '" & ServicePriorityAlertColor & "', "
	SQL = SQL & "NoActivityNagMessageONOFF_FS = " & NoActivityNagMessageONOFF_FS & ","
	SQL = SQL & "NoActivityNagMinutes_FS = " & NoActivityNagMinutes_FS & ","
	SQL = SQL & "NoActivityNagIntervalMinutes_FS = " & NoActivityNagIntervalMinutes_FS & ","
	SQL = SQL & "NoActivityNagMessageMaxToSendPerStop_FS = " & NoActivityNagMessageMaxToSendPerStop_FS & ","
	SQL = SQL & "NoActivityNagMessageMaxToSendPerDriverPerDay_FS = " & NoActivityNagMessageMaxToSendPerDriverPerDay_FS & ","
	SQL = SQL & "NoActivityNagMessageSendMethod_FS = '" & NoActivityNagMessageSendMethod_FS & "',"
	SQL = SQL & "NoActivityNagTimeOfDay_FS = '" & NoActivityNagTimeOfDay_FS & "',"
	SQL = SQL & "FS_SignatureOptional = " & FS_SignatureOptional & ","
	SQL = SQL & "FS_TechCanDecline = " & FS_TechCanDecline & ","
	SQL = SQL & "FSDefaultNotificationMethod = '" & FSDefaultNotificationMethod & "'" 
	
	'Response.Write("<br><br><br><br>" & SQL)

	Set cnn8 = Server.CreateObject("ADODB.Connection")
	cnn8.open (Session("ClientCnnString"))
	Set rs = Server.CreateObject("ADODB.Recordset")
	rs.CursorLocation = 3 
	Set rs = cnn8.Execute(SQL)
	
	set rs = Nothing
	cnn8.close
	set cnn8 = Nothing


	'***********************************************************
	'Update SQL with Request Form Field Data
	'***********************************************************

	Set cnn8 = Server.CreateObject("ADODB.Connection")
	cnn8.open (Session("ClientCnnString"))
	Set rs = Server.CreateObject("ADODB.Recordset")
	rs.CursorLocation = 3 

	SQL = "UPDATE  " & MUV_Read("SQL_Owner") & ".Settings_EmailService SET "
	SQL = SQL & "IncludeACKinDispatchEmail = " & DLinkInEmail & ","	
	SQL = SQL & "IncludeACKinDispatchText  = " & DLinkInText
	
	Response.write("<br><br>SQL SETTINGS_EMAILSERVICE: " & SQL & "<br>")

	Set cnn8 = Server.CreateObject("ADODB.Connection")
	cnn8.open (Session("ClientCnnString"))
	Set rs = Server.CreateObject("ADODB.Recordset")
	rs.CursorLocation = 3 
	Set rs = cnn8.Execute(SQL)
		

	
	'***********************************************************
	'Update SQL with Request Form Field Data
	'***********************************************************
	
	Set cnn8 = Server.CreateObject("ADODB.Connection")
	cnn8.open (Session("ClientCnnString"))
	Set rs = Server.CreateObject("ADODB.Recordset")
	rs.CursorLocation = 3 

	SQL = "UPDATE  " & MUV_Read("SQL_Owner") & ".Settings_FieldService SET "
	
	SQL = SQL & "ServiceTicketCarryoverReportOnOff = " & ServiceTicketCarryoverReportOnOff 	
	SQL = SQL & ", ServiceTicketCarryoverReportToPrimarySalesman = " & ServiceTicketCarryoverReportToPrimarySalesman 
	SQL = SQL & ", ServiceTicketCarryoverReportToSecondarySalesman = " & ServiceTicketCarryoverReportToSecondarySalesman 
	SQL = SQL & ", ServiceTicketCarryoverReportEmailSubject = '" & ServiceTicketCarryoverReportEmailSubject & "'"
	SQL = SQL & ", ServiceTicketCarryoverReportTextSummaryOnOff = " & ServiceTicketCarryoverReportTextSummaryOnOff
	SQL = SQL & ", ServiceTicketCarryoverReportUserNos = '" & ServiceTicketCarryoverReportUserNos & "'"
	SQL = SQL & ", ServiceTicketCarryoverReportTextSummaryUserNos = '" & ServiceTicketCarryoverReportTextSummaryUserNos & "'"
	SQL = SQL & ", ServiceTicketCarryoverReportTeamIntRecIDs = '" & ServiceTicketCarryoverReportTeamIntRecIDs & "'"
	SQL = SQL & ", ServiceDayStartTime = '" & ServiceDayStartTime & "'"
	SQL = SQL & ", ServiceDayEndTime = '" & ServiceDayEndTime & "'"
	SQL = SQL & ", ServiceDayElapsedTimeCalculationMethod = '" & ServiceDayElapsedTimeCalculationMethod & "'"
	SQL = SQL & ", FieldServiceNotesReportOnOff = " & FieldServiceNotesReportOnOff
	SQL = SQL & ", FieldServiceNotesReportEmailSubject = '" & FieldServiceNotesReportEmailSubject & "'"
	SQL = SQL & ", FieldServiceNotesReportUserNos = '" & FieldServiceNotesReportUserNos & "'"
	SQL = SQL & ", AutoDispatchUsersOnOff = " & AutoDispatchUsersOnOff
	SQL = SQL & ", AutoDispatchUserNos = '" & AutoDispatchUserNos & "' "
	SQL = SQL & ", CarryoverReportInclCustType = " & CarryoverReportInclCustType 
	SQL = SQL & ", ShowPartsButton = " & FS_ShowPartsButton 
	SQL = SQL & ", ServiceTicketScreenShowHoldTab = " & ServiceTicketScreenShowHoldTab
	SQL = SQL & ", CarryoverReportInclTicketNum = " & CarryoverReportInclTicketNum
	SQL = SQL & ", CarryoverReportShowRedoBreakdown = " & CarryoverReportShowRedoBreakdown
	SQL = SQL & ", ServiceTicketCarryoverReportIncludeRegions = " & ServiceTicketCarryoverReportIncludeRegions
	SQL = SQL & ", ServiceTicketthresholdReportONOFF = " & ServiceTicketthresholdReportONOFF
	SQL = SQL & ", ServiceTicketthresholdReportOnlyUndispatched = " & ServiceTicketthresholdReportOnlyUndispatched
	SQL = SQL & ", ServiceTicketthresholdReportOnlySkipFilterChanges = " & ServiceTicketthresholdReportOnlySkipFilterChanges
	SQL = SQL & ", ServiceTicketthresholdReportthresholdHours = " & ServiceTicketthresholdReportthresholdHours
	SQL = SQL & ", FSBoardKioskGlobalUseRegions = " & FSBoardKioskGlobalUseRegions
	SQL = SQL & ", FSBoardKioskGlobalTitleText = '" & FSBoardKioskGlobalTitleText & "'"
	SQL = SQL & ", FSBoardKioskGlobalTitleTextFontColor = '" & FSBoardKioskGlobalTitleTextFontColor & "'"
	SQL = SQL & ", FSBoardKioskGlobalTitleGradientColor = '" & FSBoardKioskGlobalTitleGradientColor & "'"
	SQL = SQL & ", FSBoardKioskGlobalColorAwaitingDispatch = '" & FSBoardKioskGlobalColorAwaitingDispatch & "'"
	SQL = SQL & ", FSBoardKioskGlobalColorAwaitingAcknowledgement = '" & FSBoardKioskGlobalColorAwaitingAcknowledgement & "'"
	SQL = SQL & ", FSBoardKioskGlobalColorDispatchAcknowledged = '" & FSBoardKioskGlobalColorDispatchAcknowledged & "'"
	SQL = SQL & ", FSBoardKioskGlobalColorDispatchDeclined = '" & FSBoardKioskGlobalColorDispatchDeclined & "'"
	SQL = SQL & ", FSBoardKioskGlobalColorEnRoute = '" & FSBoardKioskGlobalColorEnRoute & "'"
	SQL = SQL & ", FSBoardKioskGlobalColorOnSite = '" & FSBoardKioskGlobalColorOnSite & "'"
	SQL = SQL & ", FSBoardKioskGlobalColorRedoSwap = '" & FSBoardKioskGlobalColorRedoSwap & "'"
	SQL = SQL & ", FSBoardKioskGlobalColorRedoWaitForParts = '" & FSBoardKioskGlobalColorRedoWaitForParts & "'"
	SQL = SQL & ", FSBoardKioskGlobalColorRedoFollowUp = '" & FSBoardKioskGlobalColorRedoFollowUp & "'"
	SQL = SQL & ", FSBoardKioskGlobalColorRedoUnableToWork = '" & FSBoardKioskGlobalColorRedoUnableToWork & "'"
	SQL = SQL & ", FSBoardKioskGlobalColorClosed = '" & FSBoardKioskGlobalColorClosed & "'"
	SQL = SQL & ", FSBoardKioskGlobalColorUrgent = '" & FSBoardKioskGlobalColorUrgent & "'"
	
	Set cnn8 = Server.CreateObject("ADODB.Connection")
	cnn8.open (Session("ClientCnnString"))
	Set rs = Server.CreateObject("ADODB.Recordset")
	rs.CursorLocation = 3 
	Set rs = cnn8.Execute(SQL)
		
	'response.write(SQL)
		
	Response.Redirect("field-service.asp")

%><!--#include file="../../../inc/footer-main.asp"-->