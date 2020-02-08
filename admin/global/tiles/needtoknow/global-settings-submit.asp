<!--#include file="../../../../inc/header.asp"-->

<%

	'***********************************************************
	'Get Values Of All Form Fields Posted
	'***********************************************************
	
	N2KGlobalSettingsEmailToUserNos = Request.Form("lstSelectedN2KAPIEmailToUserNos")
	N2KGlobalSettingsUserNosToCC = Request.Form("lstSelectedN2KAPIUserNosToCC")
	If Request.Form("chkMissingClientLogoFile") = "on" then MissingClientLogoFile = 1 Else MissingClientLogoFile = 0
	If Request.Form("chkMissingHolidayinCompanyCalendar") = "on" then MissingHolidayinCompanyCalendar = 1 Else MissingHolidayinCompanyCalendar = 0
	If Request.Form("chkN2KGlobalSettingsReportONOFF") = "on" then N2KGlobalSettingsReportONOFF = 1 Else N2KGlobalSettingsReportONOFF = 0
		
	'***********************************************************
	'Get Original Values For Audit Trail Entries
	'***********************************************************
	
	SQL = "SELECT * FROM Settings_NeedToKnow"
	
	Set cnn8 = Server.CreateObject("ADODB.Connection")
	cnn8.open (Session("ClientCnnString"))
	Set rs = Server.CreateObject("ADODB.Recordset")
	rs.CursorLocation = 3 
	Set rs = cnn8.Execute(SQL)
	
	If not rs.EOF Then
		N2KGlobalSettingsEmailToUserNos_ORIG = rs("N2KGlobalSettingsEmailToUserNos")
		N2KGlobalSettingsUserNosToCC_ORIG = rs("N2KGlobalSettingsUserNosToCC")
		N2KGlobalSettingsReportONOFF_ORIG = rs("N2KGlobalSettingsReportONOFF")
		N2KGlobalIncludeMissingClientLogoFile_ORIG = rs("N2KGlobalIncludeMissingClientLogoFile")
		N2KGlobalIncludeMissingHolidayinCompanyCalendar_ORIG = rs("N2KGlobalIncludeMissingHolidayinCompanyCalendar")
	End If
				
	set rs = Nothing
	cnn8.close
	set cnn8 = Nothing
	
	'*************************************************************************
	'See if this is the first time entering data in Settings_NeedToKnow
	'*************************************************************************
	
	SQL = "SELECT * FROM Settings_NeedToKnow"
	
	Set cnn8 = Server.CreateObject("ADODB.Connection")
	cnn8.open (Session("ClientCnnString"))
	Set rs = Server.CreateObject("ADODB.Recordset")
	rs.CursorLocation = 3 
	Set rs = cnn8.Execute(SQL)
	
	If rs.EOF Then
		SettingsNeedToKnowHasRecords = false
	Else
		SettingsNeedToKnowHasRecords = true	
	End If
				
	set rs = Nothing
	cnn8.close
	set cnn8 = Nothing
	
	

	'***********************************************************
	'Update/Insert SQL with Request Form Field Data
	'***********************************************************

	If SettingsNeedToKnowHasRecords = true Then
	
		SQL = "UPDATE  " & MUV_Read("SQL_Owner") & ".Settings_NeedToKnow SET  "
		SQL = SQL & "N2KGlobalSettingsEmailToUserNos = '" & N2KGlobalSettingsEmailToUserNos & "',"
		SQL = SQL & "N2KGlobalSettingsUserNosToCC = '" & N2KGlobalSettingsUserNosToCC & "',"
		SQL = SQL & "N2KGlobalSettingsReportONOFF = " & N2KGlobalSettingsReportONOFF & ","
		SQL = SQL & "N2KGlobalIncludeMissingClientLogoFile = " & MissingClientLogoFile & ","
		SQL = SQL & "N2KGlobalIncludeMissingHolidayinCompanyCalendar = " & MissingHolidayinCompanyCalendar
	
	Else
	
		SQL = "INSERT INTO " & MUV_Read("SQL_Owner") & ".Settings_NeedToKnow "
		SQL = SQL & " (N2KGlobalSettingsEmailToUserNos, N2KGlobalSettingsUserNosToCC,N2KGlobalSettingsReportONOFF,N2KGlobalIncludeMissingClientLogoFile,N2KGlobalIncludeMissingHolidayinCompanyCalendar) "
		SQL = SQL & " VALUES "
		SQL = SQL & " ('" & N2KGlobalSettingsEmailToUserNos & "', '" & N2KGlobalSettingsUserNosToCC & "'," & N2KGlobalSettingsReportONOFF & "," & MissingClientLogoFile & "," & MissingHolidayinCompanyCalendar & ") "
	
	End If
	
	'Response.write("<br><br><br>" & SQL)
	
	Set cnn8 = Server.CreateObject("ADODB.Connection")
	cnn8.open (Session("ClientCnnString"))
	Set rs = Server.CreateObject("ADODB.Recordset")
	rs.CursorLocation = 3 
	
	Set rs = cnn8.Execute(SQL)

	set rs = Nothing
	cnn8.close
	set cnn8 = Nothing
	

	'***********************************************************
	'Perform Audit Trail Entry Inserts
	'***********************************************************

	If N2KGlobalSettingsEmailToUserNos <> N2KGlobalSettingsEmailToUserNos_ORIG Then

		userNames = ""
		userNamesOrig = ""
		
	
		IndividualN2KGlobalSettingsEmailToUserNos = Split(N2KGlobalSettingsEmailToUserNos,",")
		
		for i=0 to Ubound(IndividualN2KGlobalSettingsEmailToUserNos)
		     userNames = userNames & " " & GetUserFirstAndLastNameByUserNo(IndividualN2KGlobalSettingsEmailToUserNos (i))
		next
		
		
		IndividualN2KGlobalSettingsEmailToUserNos_ORIG  = Split(N2KGlobalSettingsEmailToUserNos_ORIG,",")
		
		for i=0 to Ubound(IndividualN2KGlobalSettingsEmailToUserNos_ORIG)
		     userNamesOrig = userNamesOrig & " " & GetUserFirstAndLastNameByUserNo(IndividualN2KGlobalSettingsEmailToUserNos_ORIG(i))
		next
		
		
		CreateAuditLogEntry "Global Settings Change", "Global Settings Change", "Major", 1, "Global Settings Need To Know Reports send email to changed from " & userNamesOrig & " to " & userNames
	End If


	If N2KGlobalSettingsUserNosToCC <> N2KGlobalSettingsUserNosToCC_ORIG Then

		userNames = ""
		userNamesOrig = ""
		
	
		IndividualN2KGlobalSettingsUserNosToCC  = Split(N2KGlobalSettingsUserNosToCC,",")
		
		for i=0 to Ubound(IndividualN2KGlobalSettingsUserNosToCC)
		     userNames = userNames & " " & GetUserFirstAndLastNameByUserNo(IndividualN2KGlobalSettingsUserNosToCC (i))
		next
		
		
		IndividualN2KGlobalSettingsUserNosToCC_ORIG  = Split(IndividualN2KGlobalSettingsUserNosToCC_ORIG,",")
		
		for i=0 to Ubound(IndividualN2KGlobalSettingsUserNosToCC_ORIG)
		     userNamesOrig = userNamesOrig & " " & GetUserFirstAndLastNameByUserNo(IndividualN2KGlobalSettingsUserNosToCC_ORIG(i))
		next
		
		
		CreateAuditLogEntry "Global Settings Change", "Global Settings Change", "Major", 1, "Global Settings Need To Know Reports users to Cc changed from " & userNamesOrig & " to " & userNames
	End If

	If Request.Form("chkN2KGlobalSettingsReportONOFF")="on" then N2KGlobalSettingsReportONOFFMsg = "On" Else N2KGlobalSettingsReportONOFFMsg = "Off"
	If N2KGlobalSettingsReportONOFF_ORIG = 1 then N2KGlobalSettingsReportONOFFMsg_ORIGFMsg = "On" Else N2KGlobalSettingsReportONOFFMsg_ORIGFMsg = "Off"
	
	IF N2KGlobalSettingsReportONOFF <> N2KGlobalSettingsReportONOFF_ORIG Then
		CreateAuditLogEntry "Global Settings Change", "Global Settings Change", "Major", 1, "Global Settings need to know report changed from " & N2KGlobalSettingsReportONOFFMsg_ORIGFMsg & " to " & N2KGlobalSettingsReportONOFFMsg 
	End If



	If Request.Form("chkMissingClientLogoFile")="on" then MissingClientLogoFileMsg = "On" Else MissingClientLogoFileMsg = "Off"
	If N2KGlobalIncludeMissingClientLogoFile_ORIG = 1 then N2KGlobalIncludeMissingClientLogoFile_ORIGMsg = "On" Else N2KGlobalIncludeMissingClientLogoFile_ORIGMsg = "Off"
	
	IF MissingClientLogoFile <> N2KGlobalIncludeMissingClientLogoFile_ORIG Then
		CreateAuditLogEntry "Global Settings Change", "Global Settings Change", "Major", 1, "Global Settings need to know report Missing Client Logo File changed from " & N2KGlobalIncludeMissingClientLogoFile_ORIGMsg & " to " & MissingClientLogoFileMsg 
	End If


	If Request.Form("chkMissingHolidayinCompanyCalendar")="on" then MissingHolidayinCompanyCalendarMsg = "On" Else MissingHolidayinCompanyCalendarMsg = "Off"
	If N2KGlobalIncludeMissingHolidayinCompanyCalendar_ORIG = 1 then N2KGlobalIncludeMissingHolidayinCompanyCalendar_ORIGMsg = "On" Else N2KGlobalIncludeMissingHolidayinCompanyCalendar_ORIGMsg = "Off"
	
	IF MissingHolidayinCompanyCalendar <> N2KGlobalIncludeMissingHolidayinCompanyCalendar_ORIG Then
		CreateAuditLogEntry "Global Settings Change", "Global Settings Change", "Major", 1, "Global Settings need to know report Missing Holiday in Company Calendar changed from " & N2KGlobalIncludeMissingHolidayinCompanyCalendar_ORIGMsg & " to " & MissingHolidayinCompanyCalendarMsg 
	End If
	
	Response.Redirect("global-settings.asp")
%><!--#include file="../../../../inc/footer-main.asp"-->