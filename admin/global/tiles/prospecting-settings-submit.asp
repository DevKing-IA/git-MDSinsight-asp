<!--#include file="../../../inc/header.asp"-->

<%

	'***********************************************************
	'Get Values Of All Form Fields Posted
	'***********************************************************
		
	CRMTabLogColor = Request.Form("txtCRMTabLogColor")
	CRMTabProductsColor = Request.Form("txtCRMTabProductsColor")
	CRMTabEquipmentColor = Request.Form("txtCRMTabEquipmentColor")
	CRMTabDocumentsColor = Request.Form("txtCRMTabDocumentsColor")
	CRMTabLocationColor = Request.Form("txtCRMTabLocationColor")
	CRMTabContactsColor = Request.Form("txtCRMTabContactsColor")
	CRMTabCompetitorsColor = Request.Form("txtCRMTabCompetitorsColor")
	CRMTabOpportunityColor = Request.Form("txtCRMTabOpportunityColor")
	CRMTabAuditTrailColor = Request.Form("txtCRMTabAuditTrailColor")
	TabSocialMediaColor = Request.Form("txtTabSocialMediaColor")

	CRMMaxActivityDaysWarning = Request.Form("selCRMMaxActivityDaysWarning")
	CRMMaxActivityDaysPermitted = Request.Form("selCRMMaxActivityDaysPermitted")
	EWSDefaultApptDuration  = Request.Form("selEWSDefaultApptDuration")
	EWSDefaultMeetingDuration  = Request.Form("selEWSDefaultMeetingDuration")
	EWSPostURL = Request.Form("txtEWSPostURL")
	
	CRMHideLocationTab = Request.Form("selCRMShowHideLocationTab")
	CRMHideProductsTab = Request.Form("selCRMShowHideProductsTab")
	CRMHideEquipmentTab = Request.Form("selCRMShowHideEquipmentTab")
	
	If Request.Form("chkCRMAutoCoordinateColors") = "on" then CRMAutoCoordinateColors = 1 Else CRMAutoCoordinateColors = 0
	If Request.Form("chkShowLivePoolProspectSearchBox") = "on" then ShowLivePoolProspectSearchBox = 1 Else ShowLivePoolProspectSearchBox = 0

	If Request.Form("chkProspSnapshotOnOff") = "on" then ProspSnapshotOnOff = 1 Else ProspSnapshotOnOff = 0
	If Request.Form("chkProspSnapshotInsideSales") = "on" then ProspSnapshotInsideSales = 1 Else ProspSnapshotInsideSales = 0
	If Request.Form("chkProspSnapshotOutsideSales") = "on" then ProspSnapshotOutsideSales = 1 Else ProspSnapshotOutsideSales = 0
	ProspSnapshotAdditionalEmails = Request.Form("txtProspSnapshotAdditionalEmails")
	ProspSnapshotEmailSubject = Request.Form("txtProspSnapshotEmailSubject")
	ProspSnapshotUserNos = Request.Form("lstSelectedProspectingSnapshotReportUserIDs")
	ProspSnapshotSalesRepDisplayUserNos = Request.Form("lstSelectedProspectingSnapshotReportSalesRepUserIDs")
	
	If CRMAutoCoordinateColors = 1 Then
		CRMTileOfferingColor = CRMTabProductsColor
		CRMTileCompetitorColor = CRMTabCompetitorsColor
		CRMTileDollarsColor = CRMTabOpportunityColor
		CRMTileActivityColor = CRMTabLogColor
		CRMTileStageColor = CRMTabLogColor
	Else
		CRMTileOfferingColor = Request.Form("txtCRMTileOfferingColor")
		CRMTileCompetitorColor = Request.Form("txtCRMTileCompetitorColor")
		CRMTileDollarsColor = Request.Form("txtCRMTileDollarsColor")
		CRMTileActivityColor = Request.Form("txtCRMTileActivityColor")
		CRMTileStageColor = Request.Form("txtCRMTileStageColor")
	End If
	
	CRMTileOwnerColor = Request.Form("txtCRMTileOwnerColor")
	CRMTileCommentsColor = Request.Form("txtCRMTileCommentsColor")
	
	ProspectActivityDefaultDaysToShow = Request.Form("selProspectActivityDefaultDaysToShow")
	
	
	If Request.Form("chkProspectingWeeklyAgendaReportOnOff") = "on" then ProspectingWeeklyAgendaReportOnOff = 1 Else ProspectingWeeklyAgendaReportOnOff = 0
	ProspectingWeeklyAgendaReportEmailSubject = Request.Form("txtProspectingWeeklyAgendaReportEmailSubject")
	ProspectingWeeklyAgendaReportEmailSubject = Replace(ProspectingWeeklyAgendaReportEmailSubject, "'", "''")
	ProspectingWeeklyAgendaReportUserNos = Request.Form("lstSelectedProspectingWeeklyAgendaReportUserIDs")
	

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
		
		CRMTabLogColor_ORIG = rs("CRMTabLogColor")
		CRMTabProductsColor_ORIG = rs("CRMTabProductsColor")
		CRMTabEquipmentColor_ORIG = rs("CRMTabEquipmentColor")
		CRMTabDocumentsColor_ORIG = rs("CRMTabDocumentsColor")
		CRMTabLocationColor_ORIG = rs("CRMTabLocationColor")
		CRMTabContactsColor_ORIG = rs("CRMTabContactsColor")
		CRMTabCompetitorsColor_ORIG = rs("CRMTabCompetitorsColor")
		CRMTabOpportunityColor_ORIG = rs("CRMTabOpportunityColor")
		CRMTabAuditTrailColor_ORIG = rs("CRMTabAuditTrailColor")
		CRMTileOfferingColor_ORIG = rs("CRMTileOfferingColor")
		CRMTileCompetitorColor_ORIG = rs("CRMTileCompetitorColor")				
		CRMTileDollarsColor_ORIG = rs("CRMTileDollarsColor")		
		CRMTileActivityColor_ORIG = rs("CRMTileActivityColor")	
		CRMTileStageColor_ORIG = rs("CRMTileStageColor")	
		CRMTileOwnerColor_ORIG = rs("CRMTileOwnerColor")
		CRMTileCommentsColor_ORIG = rs("CRMTileCommentsColor")
		CRMMaxActivityDaysPermitted_ORIG = rs("CRMMaxActivityDaysPermitted")			
		CRMMaxActivityDaysWarning_ORIG = rs("CRMMaxActivityDaysWarning")
		EWSDefaultApptDuration_ORIG = rs("EWSDefaultApptDuration")
		EWSDefaultMeetingDuration_ORIG = rs("EWSDefaultMeetingDuration")
		EWSPostURL_ORIG = rs("EWSPostURL")
		CRMAutoCoordinateColors_ORIG = rs("CRMAutoCoordinateColors")
		CRMHideLocationTab_ORIG = rs("CRMHideLocationTab")
		CRMHideProductsTab_ORIG = rs("CRMHideProductsTab")
		CRMHideEquipmentTab_ORIG = rs("CRMHideEquipmentTab")
		ProspSnapshotOnOff_ORIG = rs("ProspSnapshotOnOff")			
		ProspSnapshotInsideSales_ORIG = rs("ProspSnapshotInsideSales")	
		ProspSnapshotOutsideSales_ORIG = rs("ProspSnapshotOutsideSales")	
		ProspSnapshotAdditionalEmails_ORIG = rs("ProspSnapshotAdditionalEmails")	
		ProspSnapshotEmailSubject_ORIG = rs("ProspSnapshotEmailSubject")	
		ProspSnapshotUserNos_ORIG = rs("ProspSnapshotUserNos")	
		ProspSnapshotSalesRepDisplayUserNos_ORIG = rs("ProspSnapshotSalesRepDisplayUserNos")		
		
	End If

		
	SQL = "SELECT * FROM Settings_Prospecting"
	
	Set rs = cnn8.Execute(SQL)
	
	If not rs.EOF Then
		ShowLivePoolProspectSearchBox_ORIG = rs("ShowLivePoolProspectSearchBox")			
		ProspectActivityDefaultDaysToShow_ORIG = rs("ProspectActivityDefaultDaysToShow")
		TabSocialMediaColor_ORIG = rs("TabSocialMediaColor")
		ProspectingWeeklyAgendaReportOnOff_ORIG = rs("ProspectingWeeklyAgendaReportOnOff")
		ProspectingWeeklyAgendaReportUserNos_ORIG = rs("ProspectingWeeklyAgendaReportUserNos")
		ProspectingWeeklyAgendaReportEmailSubject_ORIG = rs("ProspectingWeeklyAgendaReportEmailSubject")
		ProspectingWeeklyAgendaReportAdditionalEmails_ORIG = rs("ProspectingWeeklyAgendaReportAdditionalEmails")	
	End If

			
	set rs = Nothing
	cnn8.close
	set cnn8 = Nothing
	

	'***********************************************************
	'Perform Audit Trail Entry Inserts
	'***********************************************************
	
	If ProspectActivityDefaultDaysToShow <> ProspectActivityDefaultDaysToShow_ORIG Then
		CreateAuditLogEntry "Global Settings Change", "Global Settings Change", "Major", 1, GetTerm("Prospecting") & " activity defaults days to show change from " & ProspectActivityDefaultDaysToShow_ORIG & " to " & ProspectActivityDefaultDaysToShow
	End If
	
	If CRMTabLogColor <> CRMTabLogColor_ORIG Then
		CreateAuditLogEntry "Global Settings Change", "Global Settings Change", "Major", 1, GetTerm("Prospecting") & " - tab color for Log tab changed from  " & CRMTabLogColor_ORIG & " to " & CRMTabLogColor
	End If
	If CRMTabProductsColor <> CRMTabProductsColor_ORIG Then
		CreateAuditLogEntry "Global Settings Change", "Global Settings Change", "Major", 1, GetTerm("Prospecting") & " - tab color for Products tab changed from  " & CRMTabProductsColor_ORIG & " to " & CRMTabProductsColor
	End If
	If CRMTabEquipmentColor <> CRMTabEquipmentColor_ORIG Then
		CreateAuditLogEntry "Global Settings Change", "Global Settings Change", "Major", 1, GetTerm("Prospecting") & " - tab color for Equipment tab changed from  " & CRMTabEquipmentColor_ORIG & " to " & CRMTabEquipmentColor
	End If
	If CRMTabDocumentsColor <> CRMTabDocumentsColor_ORIG Then
		CreateAuditLogEntry "Global Settings Change", "Global Settings Change", "Major", 1, GetTerm("Prospecting") & " - tab color for Document tab changed from  " & CRMTabDocumentsColor_ORIG & " to " & CRMTabDocumentsColor
	End If
	If CRMTabLocationColor <> CRMTabLocationColor_ORIG Then
		CreateAuditLogEntry "Global Settings Change", "Global Settings Change", "Major", 1, GetTerm("Prospecting") & " - tab color for Add/Edit tab changed from  " & CRMTabLocationColor_ORIG & " to " & CRMTabLocationColor
	End If
	If CRMTabContactsColor <> CRMTabContactsColor_ORIG Then
		CreateAuditLogEntry "Global Settings Change", "Global Settings Change", "Major", 1, GetTerm("Prospecting") & " - tab color for Contacts tab changed from  " & CRMTabContactsColor_ORIG & " to " & CRMTabContactsColor
	End If
	If CRMTabCompetitorsColor <> CRMTabCompetitorsColor_ORIG Then
		CreateAuditLogEntry "Global Settings Change", "Global Settings Change", "Major", 1, GetTerm("Prospecting") & " - tab color for Competitors tab changed from  " & CRMTabCompetitorsColor_ORIG & " to " & CRMTabCompetitorsColor
	End If
	If CRMTabOpportunityColor <> CRMTabOpportunityColor_ORIG Then
		CreateAuditLogEntry "Global Settings Change", "Global Settings Change", "Major", 1, GetTerm("Prospecting") & " - tab color for Opportunities tab changed from  " & CRMTabOpportunityColor_ORIG & " to " & CRMTabOpportunityColor
	End If
	If CRMTabAuditTrailColor <> CRMTabAuditTrailColor_ORIG Then
		CreateAuditLogEntry "Global Settings Change", "Global Settings Change", "Major", 1, GetTerm("Prospecting") & " - tab color for Audit Trail tab changed from  " & CRMTabAuditTrailColor_ORIG & " to " & CRMTabAuditTrailColor
	End If
	If TabSocialMediaColor <> TabSocialMediaColor_ORIG Then
		CreateAuditLogEntry "Global Settings Change", "Global Settings Change", "Major", 1, GetTerm("Prospecting") & " - tab color for Social Media tab changed from  " & TabSocialMediaColor_ORIG & " to " & TabSocialMediaColor
	End If
	If CRMTileOfferingColor <> CRMTileOfferingColor_ORIG Then
		CreateAuditLogEntry "Global Settings Change", "Global Settings Change", "Major", 1, GetTerm("Prospecting") & " - tile color for Offering tile changed from  " & CRMTileOfferingColor_ORIG & " to " & CRMTileOfferingColor
	End If
	If CRMTileCompetitorColor <> CRMTileCompetitorColor_ORIG Then
		CreateAuditLogEntry "Global Settings Change", "Global Settings Change", "Major", 1, GetTerm("Prospecting") & " - tile color for Competitor tile changed from  " & CRMTileCompetitorColor_ORIG & " to " & CRMTileCompetitorColor
	End If
	If CRMTileDollarsColor <> CRMTileDollarsColor_ORIG Then
		CreateAuditLogEntry "Global Settings Change", "Global Settings Change", "Major", 1, GetTerm("Prospecting") & " - tile color for Dollars tile changed from  " & CRMTileDollarsColor_ORIG & " to " & CRMTileDollarsColor
	End If
	If CRMTileActivityColor <> CRMTileActivityColor_ORIG Then
		CreateAuditLogEntry "Global Settings Change", "Global Settings Change", "Major", 1, GetTerm("Prospecting") & " - tile color for Activity tile changed from  " & CRMTileActivityColor_ORIG & " to " & CRMTileActivityColor
	End If
	If CRMTileStageColor <> CRMTileStageColor_ORIG Then
		CreateAuditLogEntry "Global Settings Change", "Global Settings Change", "Major", 1, GetTerm("Prospecting") & " - tile color for Stage tile changed from  " & CRMTileStageColor_ORIG & " to " & CRMTileStageColor
	End If
	If CRMTileOwnerColor <> CRMTileOwnerColor_ORIG Then
		CreateAuditLogEntry "Global Settings Change", "Global Settings Change", "Major", 1, GetTerm("Prospecting") & " - tile color for Stage tile changed from  " & CRMTileOwnerColor_ORIG & " to " & CRMTileOwnerColor
	End If
	If CRMTileCommentsColor <> CRMTileCommentsColor_ORIG Then
		CreateAuditLogEntry "Global Settings Change", "Global Settings Change", "Major", 1, GetTerm("Prospecting") & " - tile color for Stage tile changed from  " & CRMTileCommentsColor_ORIG & " to " & CRMTileCommentsColor
	End If

	If Request.Form("chkCRMAutoCoordinateColors")="on" then CRMAutoCoordinateColors = 1 Else CRMAutoCoordinateColors = 0
	If Request.Form("chkCRMAutoCoordinateColors")="on" then CRMAutoCoordinateColorsMsg = "On" Else CRMAutoCoordinateColorsMsg = "Off"

	If Request.Form("chkShowLivePoolProspectSearchBox")="on" then ShowLivePoolProspectSearchBox = 1 Else ShowLivePoolProspectSearchBox = 0
	If Request.Form("chkShowLivePoolProspectSearchBox")="on" then ShowLivePoolProspectSearchBoxMsg = "On" Else ShowLivePoolProspectSearchBoxMsg = "Off"
	IF ShowLivePoolProspectSearchBox <> ShowLivePoolProspectSearchBox_ORIG Then
		CreateAuditLogEntry "Prospecting Settings Change", "Prospecting Settings Change", "Major", 1, "Live pool show all prospects search box changed from " & ShowLivePoolProspectSearchBox_ORIG  & " to " & ShowLivePoolProspectSearchBoxMsg 
	End If
	IF CRMAutoCoordinateColors <> CRMAutoCoordinateColors_ORIG Then
		CreateAuditLogEntry "Global Settings Change", "Global Settings Change", "Major", 1, "Auto coordinate tab/tile colors changed from " & CRMAutoCoordinateColors_ORIG & " to " & CRMAutoCoordinateColorsMsg 
	End If
	
	IF CRMAutoCoordinateColors <> CRMAutoCoordinateColors_ORIG Then
		CreateAuditLogEntry "Global Settings Change", "Global Settings Change", "Major", 1, "Auto coordinate tab/tile colors changed from " & ShowLivePoolProspectSearchBox_ORIG & " to " & ShowLivePoolProspectSearchBox
	End If
	If CRMMaxActivityDaysPermitted <> CRMMaxActivityDaysPermitted_ORIG Then
		CreateAuditLogEntry "Global Settings Change", "Global Settings Change", "Major", 1, GetTerm("Prospecting") & " - max days to permitted activity scheduling " & CRMMaxActivityDaysPermitted_ORIG & " to " & CRMMaxActivityDaysPermitted
	End If
	If CRMMaxActivityDaysWarning <> CRMMaxActivityDaysWarning_ORIG Then
		CreateAuditLogEntry "Global Settings Change", "Global Settings Change", "Major", 1, GetTerm("Prospecting") & " - max days before warning for  activity scheduling " & CRMMaxActivityDaysWarning_ORIG & " to " & CRMMaxActivityDaysWarning
	End If

	If Not IsNumeric(EWSDefaultApptDuration) Then EWSDefaultApptDuration = 30 'Default to 30 mins
	If cint(EWSDefaultApptDuration) <> EWSDefaultApptDuration_ORIG Then
		CreateAuditLogEntry "Global Settings Change", "Global Settings Change", "Major", 1, GetTerm("Prospecting") & " - default appointment duration (minutes) changed from  " & EWSDefaultApptDuration_ORIG & " to " & EWSDefaultApptDuration
	End If
	
	If Not IsNumeric(EWSDefaultMeetingDuration) Then EWSDefaultMeetingDuration = 45 'Default to 45 mins
	If cint(EWSDefaultMeetingDuration) <> EWSDefaultMeetingDuration_ORIG Then
		CreateAuditLogEntry "Global Settings Change", "Global Settings Change", "Major", 1, GetTerm("Prospecting") & " - default meeting duration (minutes) changed from  " & EWSDefaultMeetingDuration_ORIG & " to " & EWSDefaultMeetingDuration
	End If
	If EWSPostURL <> EWSPostURL_ORIG Then
		CreateAuditLogEntry "Global Settings Change", "Global Settings Change", "Major", 1, GetTerm("Prospecting") & " - EWS post URL (Exchange) changed from  " & EWSPostURL_ORIG & " to " & EWSPostURL
	End If

	If EWSPostURL <> EWSPostURL_ORIG Then
		CreateAuditLogEntry "Global Settings Change", "Global Settings Change", "Major", 1, GetTerm("Prospecting") & " - EWS post URL (Exchange) changed from  " & EWSPostURL_ORIG & " to " & EWSPostURL
	End If

    If ProspectingWeeklyAgendaReportOnOff <> ProspectingWeeklyAgendaReportOnOff_ORIG Then
		CreateAuditLogEntry "Prospecting Settings Change", "Prospecting Settings Change", "Major", 1, "Prospecting Weekly Agenda Report On/Off changed from " & ProspectingWeeklyAgendaReportOnOff_ORIG & " to " & ProspectingWeeklyAgendaReportOnOff
	End If

    If ProspectingWeeklyAgendaReportUserNos <> ProspectingWeeklyAgendaReportUserNos_ORIG Then
		CreateAuditLogEntry "Prospecting Settings Change", "Prospecting Settings Change", "Major", 1, "Prospecting Weekly Agenda Report User Nos changed from " & ProspectingWeeklyAgendaReportUserNos_ORIG & " to " & ProspectingWeeklyAgendaReportUserNos
	End If

    If ProspectingWeeklyAgendaReportEmailSubject <> ProspectingWeeklyAgendaReportEmailSubject_ORIG Then
		CreateAuditLogEntry "Prospecting Settings Change", "Prospecting Settings Change", "Major", 1, "Prospecting Weekly Agenda Report Email Subject changed from " & ProspectingWeeklyAgendaReportEmailSubject_ORIG & " to " & ProspectingWeeklyAgendaReportEmailSubject
	End If

	'***********************************************************
	'Update SQL with Request Form Field Data
	'***********************************************************	
	
	SQL = "UPDATE  " & MUV_Read("SQL_Owner") & ".Settings_Global SET "
	SQL = SQL & "CRMTabLogColor = '" & CRMTabLogColor & "',"
	SQL = SQL & "CRMTabProductsColor = '" & CRMTabProductsColor & "',"
	SQL = SQL & "CRMTabEquipmentColor = '" & CRMTabEquipmentColor & "',"
	SQL = SQL & "CRMTabDocumentsColor = '" & CRMTabDocumentsColor & "',"
	SQL = SQL & "CRMTabLocationColor = '" & CRMTabLocationColor & "',"
	SQL = SQL & "CRMTabContactsColor = '" & CRMTabContactsColor & "',"
	SQL = SQL & "CRMTabCompetitorsColor = '" & CRMTabCompetitorsColor & "',"
	SQL = SQL & "CRMTabOpportunityColor = '" & CRMTabOpportunityColor & "',"
	SQL = SQL & "CRMTabAuditTrailColor = '" & CRMTabAuditTrailColor & "',"
	SQL = SQL & "CRMTileOfferingColor = '" & CRMTileOfferingColor & "',"
	SQL = SQL & "CRMTileCompetitorColor = '" & CRMTileCompetitorColor & "',"
	SQL = SQL & "CRMTileDollarsColor = '" & CRMTileDollarsColor & "',"
	SQL = SQL & "CRMTileActivityColor = '" & CRMTileActivityColor & "',"
	SQL = SQL & "CRMTileStageColor = '" & CRMTileStageColor & "',"
	SQL = SQL & "CRMTileOwnerColor = '" & CRMTileOwnerColor & "',"
	SQL = SQL & "CRMTileCommentsColor = '" & CRMTileCommentsColor & "',"
	SQL = SQL & "CRMMaxActivityDaysPermitted = " & CRMMaxActivityDaysPermitted & ","	
	SQL = SQL & "CRMMaxActivityDaysWarning = " & CRMMaxActivityDaysWarning & ","	
	SQL = SQL & "EWSDefaultApptDuration = '" & EWSDefaultApptDuration  & "',"
	SQL = SQL & "EWSDefaultMeetingDuration = '" & EWSDefaultMeetingDuration & "',"
	SQL = SQL & "EWSPostURL = '" & EWSPostURL & "',"
	SQL = SQL & "CRMAutoCoordinateColors = " & CRMAutoCoordinateColors & ","
	SQL = SQL & "CRMHideLocationTab = " & CRMHideLocationTab & ","
	SQL = SQL & "CRMHideProductsTab = " & CRMHideProductsTab & ","
	SQL = SQL & "CRMHideEquipmentTab = " & CRMHideEquipmentTab & ","
	SQL = SQL & "ProspSnapshotOnOff = " & ProspSnapshotOnOff & ","
	SQL = SQL & "ProspSnapshotInsideSales = " & ProspSnapshotInsideSales & ","
	SQL = SQL & "ProspSnapshotOutsideSales = " & ProspSnapshotOutsideSales & ","
	SQL = SQL & "ProspSnapshotAdditionalEmails = '" & ProspSnapshotAdditionalEmails & "',"
	SQL = SQL & "ProspSnapshotEmailSubject = '" & ProspSnapshotEmailSubject & "',"
	SQL = SQL & "ProspSnapshotUserNos = '" & ProspSnapshotUserNos & "',"	
	SQL = SQL & "ProspSnapshotSalesRepDisplayUserNos = '" & ProspSnapshotSalesRepDisplayUserNos & "'"	

	Set cnn8 = Server.CreateObject("ADODB.Connection")
	cnn8.open (Session("ClientCnnString"))
	Set rs = Server.CreateObject("ADODB.Recordset")
	rs.CursorLocation = 3 
	Set rs = cnn8.Execute(SQL)

	SQL = "UPDATE  Settings_Prospecting SET "
	SQL = SQL & " ShowLivePoolProspectSearchBox = " & ShowLivePoolProspectSearchBox 
	SQL = SQL & ", ProspectActivityDefaultDaysToShow = " & ProspectActivityDefaultDaysToShow & " "
	SQL = SQL & ", TabSocialMediaColor = '" & TabSocialMediaColor & "' "
	SQL = SQL & ", ProspectingWeeklyAgendaReportOnOff = " & ProspectingWeeklyAgendaReportOnOff & " "
	SQL = SQL & ", ProspectingWeeklyAgendaReportEmailSubject = '" & ProspectingWeeklyAgendaReportEmailSubject & "' "
	SQL = SQL & ", ProspectingWeeklyAgendaReportUserNos = '" & ProspectingWeeklyAgendaReportUserNos & "' "	
	
	Set rs = cnn8.Execute(SQL)
	
	set rs = Nothing
	cnn8.close
	set cnn8 = Nothing



	Response.Redirect("prospecting-settings.asp")
%>
<!--#include file="../../../inc/footer-main.asp"-->