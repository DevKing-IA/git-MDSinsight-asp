<!--#include file="../../../inc/header.asp"-->

<%

	'***********************************************************
	'Get Values Of All Form Fields Posted
	'***********************************************************
		
	FilterChangeDays = Request.Form("selFilterChangeDays")
	FilterChangeDaysFieldService = Request.Form("selFilterChangeDaysFieldService")
	FilterChangeIndicatorAndButtonColor = Request.Form("txtFilterChangeIndicatorAndButtonColor")
	If Request.Form("chkShowSeparateFilterChangesTabOnServiceScreen") = "on" then ShowSeparateFilterChangesTabOnServiceScreen = 1 Else ShowSeparateFilterChangesTabOnServiceScreen = 0
	
	If Request.Form("chkAutoFilterChangeGenerationONOFF") = "on" then AutoFilterChangeGenerationONOFF = 1 Else AutoFilterChangeGenerationONOFF = 0
	If Request.Form("chkAutoFilterChangeUseRegions") = "on" then AutoFilterChangeUseRegions = 1 Else AutoFilterChangeUseRegions = 0
	AutoFilterChangeMaxNumTicketsPerDay = Request.Form("selAutoFilterChangeMaxNumTicketsPerDay")
		
	If Request.Form("chkFilterChangeEmail") = "on" then CompletedFilterChangeEmailOn = 1 Else CompletedFilterChangeEmailOn = 0
	If Request.Form("chkDoNotSendClientCompletedFilter") = "on" then DoNotSendClientCompletedFilter = 1 Else DoNotSendClientCompletedFilter = 0
	If Request.Form("chkFilterChangePDFIncludeServiceNotes") = "on" then FilterChangePDFIncludeServiceNotes = 1 Else FilterChangePDFIncludeServiceNotes = 0
	SendCompletedFilterChangesTo = Request.Form("txtFilterChangeEmailsTo")
	SendCompletedFilterChangesTo = Trim(SendCompletedFilterChangesTo)
	SendCompletedFilterChangesTo = Replace(SendCompletedFilterChangesTo," ","")
	SendCompletedFilterChangesTo = Replace(SendCompletedFilterChangesTo,vbCRLF,"")
	SendCompletedFilterChangesTo = Replace(SendCompletedFilterChangesTo,vbTab,"")
	If Trim(SendCompletedFilterChangesTo) <> "" Then
		If Right(SendCompletedFilterChangesTo,1)<>";" Then SendCompletedFilterChangesTo = SendCompletedFilterChangesTo & ";"
	End If
	
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
		FilterChangeDays_ORIG = rs("FilterChangeDays")
		FilterChangeDaysFieldService_ORIG = rs("FilterChangeDaysFieldService")	
	End If
	
	set rs = Nothing
	cnn8.close
	set cnn8 = Nothing
	
	
	
	SQL = "SELECT * FROM Settings_EmailService"
	
	Set cnn8 = Server.CreateObject("ADODB.Connection")
	cnn8.open (Session("ClientCnnString"))
	Set rs = Server.CreateObject("ADODB.Recordset")
	rs.CursorLocation = 3 
	Set rs = cnn8.Execute(SQL)
	
	If not rs.EOF Then
		FilterChangePDFIncludeServiceNotes_ORIG = rs("FilterChangePDFIncludeServiceNotes")
		CompletedFilterChangeEmailOn_ORIG = rs("CompletedFilterChangeEmailOn")
		DoNotSendClientCompletedFilter_ORIG = rs("DoNotSendClientCompletedFilter")
		SendCompletedFilterChangesTo_ORIG = rs("SendCompletedFilterChangesTo")
	End If
	
	set rs = Nothing
	cnn8.close
	set cnn8 = Nothing

	
	
	SQL = "SELECT * FROM Settings_FieldService"
	
	Set cnn8 = Server.CreateObject("ADODB.Connection")
	cnn8.open (Session("ClientCnnString"))
	Set rs = Server.CreateObject("ADODB.Recordset")
	rs.CursorLocation = 3 
	Set rs = cnn8.Execute(SQL)
	
	If not rs.EOF Then
		FilterChangeIndicatorAndButtonColor_ORIG = rs("FilterChangeIndicatorAndButtonColor")	
		ShowSeparateFilterChangesTabOnServiceScreen_ORIG = rs("ShowSeparateFilterChangesTabOnServiceScreen")
		AutoFilterChangeGenerationONOFF_ORIG = rs("AutoFilterChangeGenerationONOFF")
		AutoFilterChangeUseRegions_ORIG = rs("AutoFilterChangeUseRegions")
		AutoFilterChangeMaxNumTicketsPerDay_ORIG = rs("AutoFilterChangeMaxNumTicketsPerDay")
	End If
	
	set rs = Nothing
	cnn8.close
	set cnn8 = Nothing

	'***********************************************************
	'Perform Audit Trail Entry Inserts
	'***********************************************************
	
	IF cint(FilterChangeDays) <> cint(FilterChangeDays_ORIG) Then
		CreateAuditLogEntry "Global Settings Change", "Global Settings Change", "Major", 1, "Service screen - show filter changes within X days changed from  " & FilterChangeDays_ORIG & " to " & FilterChangeDays
	End If
	
	IF cint(FilterChangeDaysFieldService) <> cint(FilterChangeDaysFieldService_ORIG) Then
		CreateAuditLogEntry "Global Settings Change", "Global Settings Change", "Major", 1, "Field Service - show filter changes within X days changed from  " & FilterChangeDaysFieldService_ORIG & " to " & FilterChangeDaysFieldService 
	End If
	
	If FilterChangeIndicatorAndButtonColor_ORIG <> FilterChangeIndicatorAndButtonColor Then
		CreateAuditLogEntry "Global Settings Change", "Global Settings Change", "Major", 1, "Filter settings - Filter Change Indicator And Button Color changed from  " & FilterChangeIndicatorAndButtonColor_ORIG & " to " & FilterChangeIndicatorAndButtonColor
	End If

	If Request.Form("chkShowSeparateFilterChangesTabOnServiceScreen") = "on" then ShowSeparateFilterChangesTabMsg = "On" Else ShowSeparateFilterChangesTabMsg = "Off"
	If ShowSeparateFilterChangesTabOnServiceScreen_ORIG = 1 then ShowSeparateFilterChangesOrigTabMsg = "On" Else ShowSeparateFilterChangesOrigTabMsg = "Off"

	IF ShowSeparateFilterChangesTabOnServiceScreen <> ShowSeparateFilterChangesTabOnServiceScreen_ORIG Then
		CreateAuditLogEntry "Global Settings Change", "Global Settings Change", "Major", 1, "Filter settings - Show separate filter tab on service board changed from " & ShowSeparateFilterChangesOrigTabMsg_ORIG & " to " & ShowSeparateFilterChangesTabMsg
	End If
	
	If Request.Form("chkAutoFilterChangeGenerationONOFF") = "on" then AutoFilterChangeGenerationONOFFMsg = "On" Else AutoFilterChangeGenerationONOFFMsg = "Off"
	If AutoFilterChangeGenerationONOFF_ORIG = 1 then AutoFilterChangeGenerationONOFFMsgOrig = "On" Else AutoFilterChangeGenerationONOFFMsgOrig = "Off"
	IF AutoFilterChangeGenerationONOFF <> AutoFilterChangeGenerationONOFF_ORIG Then
		CreateAuditLogEntry "Global Settings Change", "Global Settings Change", "Major", 1, "Automated filter change ticket generation changed from " & AutoFilterChangeGenerationONOFFMsgOrig & " to " & AutoFilterChangeGenerationONOFFMsg
	End If


	If Request.Form("chkAutoFilterChangeUseRegions") = "on" then AutoFilterChangeUseRegionsMsg = "On" Else AutoFilterChangeUseRegionsMsg = "Off"
	If AutoFilterChangeUseRegions_ORIG = 1 then AutoFilterChangeUseRegionsMsgOrig = "On" Else AutoFilterChangeUseRegionsMsgOrig = "Off"
	IF AutoFilterChangeUseRegions <> AutoFilterChangeUseRegions_ORIG Then
		CreateAuditLogEntry "Global Settings Change", "Global Settings Change", "Major", 1, "Automated filter change use regions setting changed from " & AutoFilterChangeUseRegionsMsgOrig & " to " & AutoFilterChangeUseRegionsMsg
	End If

	If AutoFilterChangeMaxNumTicketsPerDay <> AutoFilterChangeMaxNumTicketsPerDay_ORIG Then
		CreateAuditLogEntry "Global Settings Change", "Global Settings Change", "Major", 1, "Automated filter change max tickets to generate per day changed from " & AutoFilterChangeMaxNumTicketsPerDay_ORIG & " to " & AutoFilterChangeMaxNumTicketsPerDay
	End If
	

	If Request.Form("chkFilterChangeEmail") = "on" then CompletedFilterChangeEmailOnMsg = "On" Else CompletedFilterChangeEmailOnMsg = "Off"
	If CompletedFilterChangeEmailOn_ORIG = 1 then CompletedFilterChangeEmailOnMsgOrig = "On" Else CompletedFilterChangeEmailOnMsgOrig = "Off"
	IF CompletedFilterChangeEmailOn <> CompletedFilterChangeEmailOn_ORIG Then
		CreateAuditLogEntry "Global Settings Change", "Global Settings Change", "Major", 1, "Filter settings - Completed filter change triggers email changed from " & CompletedFilterChangeEmailOnMsgOrig & " to " & CompletedFilterChangeEmailOnMsg
	End If
	
	If Request.Form("chkDoNotSendClientCompletedFilter") = "on" then DoNotSendClientCompletedFilterOnMsg = "On" Else DoNotSendClientCompletedFilterOnMsg = "Off"
	If DoNotSendClientCompletedFilterOn_ORIG = 1 then DoNotSendClientCompletedFilterOnMsgOrig = "On" Else DoNotSendClientCompletedFilterOnMsgOrig = "Off"
	IF DoNotSendClientCompletedFilterOn <> DoNotSendClientCompletedFilterOn_ORIG Then
		CreateAuditLogEntry "Global Settings Change", "Global Settings Change", "Major", 1, "Filter settings - Do not send an email to clients for completed filter changes changed from " & DoNotSendClientCompletedFilterOnMsgOrig & " to " & DoNotSendClientCompletedFilterOnMsg
	End If
	
	If Request.Form("chkFilterChangePDFIncludeServiceNotes") = "on" then FilterChangePDFIncludeServiceNotesOnMsg = "On" Else FilterChangePDFIncludeServiceNotesOnMsg = "Off"
	If FilterChangePDFIncludeServiceNotesOn_ORIG = 1 then FilterChangePDFIncludeServiceNotesOnMsgOrig = "On" Else FilterChangePDFIncludeServiceNotesOnMsgOrig = "Off"
	IF FilterChangePDFIncludeServiceNotesOn <> FilterChangePDFIncludeServiceNotesOn_ORIG Then
		CreateAuditLogEntry "Global Settings Change", "Global Settings Change", "Major", 1, "Filter settings - Include service notes in filter change emailed .pdf changed from " & FilterChangePDFIncludeServiceNotesOnMsgOrig & " to " & FilterChangePDFIncludeServiceNotesOnMsg
	End If


	If SendCompletedFilterChangesTo_ORIG <> SendCompletedFilterChangesTo Then
		CreateAuditLogEntry "Global Settings Change", "Global Settings Change", "Major", 1, "Filter settings - Send completed filter change email to the following additional addresses changed from  " & SendCompletedFilterChangesTo_ORIG & " to " & SendCompletedFilterChangesTo
	End If


	'***********************************************************
	'Update SQL with Request Form Field Data
	'***********************************************************
		
	SQL = "UPDATE Settings_Global SET "
	SQL = SQL & "FilterChangeDays = " & FilterChangeDays & ","
	SQL = SQL & "FilterChangeDaysFieldService = " & FilterChangeDaysFieldService

	Set cnn8 = Server.CreateObject("ADODB.Connection")
	cnn8.open (Session("ClientCnnString"))
	Set rs = Server.CreateObject("ADODB.Recordset")
	rs.CursorLocation = 3 
	Set rs = cnn8.Execute(SQL)
	
	set rs = Nothing
	cnn8.close
	set cnn8 = Nothing
	
	
	SQL = "UPDATE Settings_EmailService SET "
	SQL = SQL & "CompletedFilterChangeEmailOn = " & CompletedFilterChangeEmailOn & ","
	SQL = SQL & "DoNotSendClientCompletedFilter = " & DoNotSendClientCompletedFilter& ","
	SQL = SQL & "FilterChangePDFIncludeServiceNotes = " & FilterChangePDFIncludeServiceNotes & ","
	SQL = SQL & "SendCompletedFilterChangesTo = '" & SendCompletedFilterChangesTo & "'"

	Set cnn8 = Server.CreateObject("ADODB.Connection")
	cnn8.open (Session("ClientCnnString"))
	Set rs = Server.CreateObject("ADODB.Recordset")
	rs.CursorLocation = 3 
	Set rs = cnn8.Execute(SQL)
	
	set rs = Nothing
	cnn8.close
	set cnn8 = Nothing


	SQL = "UPDATE Settings_FieldService SET "
	SQL = SQL & "FilterChangeIndicatorAndButtonColor = '" & FilterChangeIndicatorAndButtonColor & "', "
	SQL = SQL & "ShowSeparateFilterChangesTabOnServiceScreen = " & ShowSeparateFilterChangesTabOnServiceScreen & ", "
	SQL = SQL & "AutoFilterChangeGenerationONOFF = " & AutoFilterChangeGenerationONOFF & ", "
	SQL = SQL & "AutoFilterChangeUseRegions = " & AutoFilterChangeUseRegions & ", "
	SQL = SQL & "AutoFilterChangeMaxNumTicketsPerDay = " & AutoFilterChangeMaxNumTicketsPerDay
	
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