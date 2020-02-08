<!--#include file="../../../../inc/header.asp"-->

<%

	'***********************************************************
	'Get Values Of All Form Fields Posted
	'***********************************************************
	
	N2KAccountsReceivableEmailToUserNos = Request.Form("lstSelectedN2KAPIEmailToUserNos")
	N2KAccountsReceivableUserNosToCC = Request.Form("lstSelectedN2KAPIUserNosToCC")

	If Request.Form("chkEmptyAddress2") = "on" then EmptyAddress2 = 1 Else EmptyAddress2 = 0
	If Request.Form("chkEmptyCity") = "on" then EmptyCity = 1 Else EmptyCity = 0
	If Request.Form("chkEmptyCityStateZip") = "on" then EmptyCityStateZip = 1 Else EmptyCityStateZip = 0
	If Request.Form("chkEmptyCustomerName") = "on" then EmptyCustomerName = 1 Else EmptyCustomerName = 0
	If Request.Form("chkEmptyPhoneNumber") = "on" then EmptyPhoneNumber = 1 Else EmptyPhoneNumber = 0
	If Request.Form("chkEmptyState") = "on" then EmptyState = 1 Else EmptyState = 0
	If Request.Form("chkEmptyZip") = "on" then EmptyZip = 1 Else EmptyZip = 0
	If Request.Form("chkInvalidCityStateZip") = "on" then InvalidCityStateZip = 1 Else InvalidCityStateZip = 0
	If Request.Form("chkInvalidPhoneNumber") = "on" then InvalidPhoneNumber = 1 Else InvalidPhoneNumber = 0
	If Request.Form("chkInvalidState") = "on" then InvalidState = 1 Else InvalidState = 0
	If Request.Form("chkInvalidZipCode") = "on" then InvalidZipCode = 1 Else InvalidZipCode = 0
	If Request.Form("chkMissingcustomertype") = "on" then Missingcustomertype = 1 Else Missingcustomertype = 0
	If Request.Form("chkMissingprimarysalesman") = "on" then Missingprimarysalesman = 1 Else Missingprimarysalesman = 0
	If Request.Form("chkMissingsecondarysalesman") = "on" then Missingsecondarysalesman = 1 Else Missingsecondarysalesman = 0
	If Request.Form("chkNotAssignedToRegion") = "on" then NotAssignedToRegion = 1 Else NotAssignedToRegion = 0

	If Request.Form("chkN2KARReportONOFF") = "on" then N2KARReportONOFF = 1 Else N2KARReportONOFF = 0
		
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
		
		N2KAccountsReceivableEmailToUserNos_ORIG = rs("N2KAREmailToUserNos")
		N2KAccountsReceivableUserNosToCC_ORIG = rs("N2KARUserNosToCC")
		N2KARReportONOFF_ORIG = rs("N2KARReportONOFF")
		N2KARIncludeEmptyAddress2_ORIG = rs("N2KARIncludeEmptyAddress2")
		N2KARIncludeEmptyCity_ORIG = rs("N2KARIncludeEmptyCity")
		N2KARIncludeEmptyCityStateZip_ORIG = rs("N2KARIncludeEmptyCityStateZip")
		N2KARIncludeEmptyCustomerName_ORIG = rs("N2KARIncludeEmptyCustomerName")
		N2KARIncludeEmptyPhoneNumber_ORIG = rs("N2KARIncludeEmptyPhoneNumber")
		N2KARIncludeEmptyState_ORIG = rs("N2KARIncludeEmptyState")
		N2KARIncludeEmptyZip_ORIG = rs("N2KARIncludeEmptyZip")
		N2KARIncludeInvalidCityStateZip_ORIG = rs("N2KARIncludeInvalidCityStateZip")
		N2KARIncludeInvalidPhoneNumber_ORIG = rs("N2KARIncludeInvalidPhoneNumber")
		N2KARIncludeInvalidState_ORIG = rs("N2KARIncludeInvalidState")
		N2KARIncludeInvalidZipCode_ORIG = rs("N2KARIncludeInvalidZipCode")
		N2KARIncludeMissingcustomertype_ORIG = rs("N2KARIncludeMissingcustomertype")
		N2KARIncludeMissingprimarysalesman_ORIG = rs("N2KARIncludeMissingprimarysalesman")
		N2KARIncludeMissingsecondarysalesman_ORIG = rs("N2KARIncludeMissingsecondarysalesman")
		
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
		SQL = SQL & "N2KAREmailToUserNos = '" & N2KAccountsReceivableEmailToUserNos & "',"
		SQL = SQL & "N2KARUserNosToCC = '" & N2KAccountsReceivableUserNosToCC & "',"
		SQL = SQL & "N2KARReportONOFF = " & N2KARReportONOFF & ","
		SQL = SQL & "N2KARIncludeEmptyAddress2 = " & EmptyAddress2 & ","
		SQL = SQL & "N2KARIncludeEmptyCity = " & EmptyCity & ","
		SQL = SQL & "N2KARIncludeEmptyCityStateZip = " & EmptyCityStateZip & ","
		SQL = SQL & "N2KARIncludeEmptyCustomerName = " & EmptyCustomerName & ","
		SQL = SQL & "N2KARIncludeEmptyPhoneNumber = " & EmptyPhoneNumber & ","
		SQL = SQL & "N2KARIncludeEmptyState = " & EmptyState & ","
		SQL = SQL & "N2KARIncludeEmptyZip = " & EmptyZip & ","
		SQL = SQL & "N2KARIncludeInvalidCityStateZip = " & InvalidCityStateZip & ","
		SQL = SQL & "N2KARIncludeInvalidPhoneNumber = " & InvalidPhoneNumber & ","
		SQL = SQL & "N2KARIncludeInvalidState = " & InvalidState & ","
		SQL = SQL & "N2KARIncludeInvalidZipCode = " & InvalidZipCode & ","
		SQL = SQL & "N2KARIncludeMissingcustomertype = " & Missingcustomertype & ","
		SQL = SQL & "N2KARIncludeMissingprimarysalesman = " & Missingprimarysalesman & ","
		SQL = SQL & "N2KARIncludeMissingsecondarysalesman = " & Missingsecondarysalesman & ","
		SQL = SQL & "N2KARIncludeNotAssignedToRegion= " & NotAssignedToRegion
	
	Else
	
		SQL = "INSERT INTO " & MUV_Read("SQL_Owner") & ".Settings_NeedToKnow "
		SQL = SQL & " (N2KAREmailToUserNos, N2KARUserNosToCC,N2KARReportONOFF,	N2KARIncludeEmptyAddress2, N2KARIncludeEmptyCity, N2KARIncludeEmptyCityStateZip, N2KARIncludeEmptyCustomerName, N2KARIncludeEmptyPhoneNumber, N2KARIncludeEmptyState, N2KARIncludeEmptyZip, N2KARIncludeInvalidCityStateZip, N2KARIncludeInvalidPhoneNumber, N2KARIncludeInvalidState, N2KARIncludeInvalidZipCode, N2KARIncludeMissingcustomertype, N2KARIncludeMissingprimarysalesman, N2KARIncludeMissingsecondarysalesman, N2KARIncludeNotAssignedToRegion) "
		SQL = SQL & " VALUES "
		SQL = SQL & " ('" & N2KAccountsReceivableEmailToUserNos & "','" & N2KAccountsReceivableUserNosToCC & "'," & N2KARReportONOFF & "," & EmptyAddress2 & "," & EmptyCity & "," & EmptyCityStateZip & "," & EmptyCustomerName & "," & EmptyPhoneNumber & "," & EmptyState & "," & EmptyZip & "," & InvalidCityStateZip & "," & InvalidPhoneNumber & "," & InvalidState & "," & InvalidZipCode & "," & Missingcustomertype & "," & Missingprimarysalesman & "," & Missingsecondarysalesman & "," & NotAssignedToRegion & ") "	
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

	If N2KAccountsReceivableEmailToUserNos <> N2KAccountsReceivableEmailToUserNos_ORIG Then

		userNames = ""
		userNamesOrig = ""
		
	
		IndividualN2KAccountsReceivableEmailToUserNos = Split(N2KAccountsReceivableEmailToUserNos,",")
		
		for i=0 to Ubound(IndividualN2KAccountsReceivableEmailToUserNos)
		     userNames = userNames & " " & GetUserFirstAndLastNameByUserNo(IndividualN2KAccountsReceivableEmailToUserNos (i))
		next
		
		
		IndividualN2KAccountsReceivableEmailToUserNos_ORIG  = Split(N2KAccountsReceivableEmailToUserNos_ORIG,",")
		
		for i=0 to Ubound(IndividualN2KAccountsReceivableEmailToUserNos_ORIG)
		     userNamesOrig = userNamesOrig & " " & GetUserFirstAndLastNameByUserNo(IndividualN2KAccountsReceivableEmailToUserNos_ORIG(i))
		next
		
		
		CreateAuditLogEntry "Global Settings Change", "Global Settings Change", "Major", 1, GetTerm("Accounts Receivable") & " Need To Know Reports send email to changed from " & userNamesOrig & " to " & userNames
	End If


	If N2KAccountsReceivableUserNosToCC <> N2KAccountsReceivableUserNosToCC_ORIG Then

		userNames = ""
		userNamesOrig = ""
		
	
		IndividualN2KARUserNosToCC  = Split(N2KARUserNosToCC,",")
		
		for i=0 to Ubound(IndividualN2KARUserNosToCC)
		     userNames = userNames & " " & GetUserFirstAndLastNameByUserNo(IndividualN2KARUserNosToCC (i))
		next
		
		
		IndividualN2KARUserNosToCC_ORIG  = Split(IndividualN2KARUserNosToCC_ORIG,",")
		
		for i=0 to Ubound(IndividualN2KARUserNosToCC_ORIG)
		     userNamesOrig = userNamesOrig & " " & GetUserFirstAndLastNameByUserNo(IndividualN2KARUserNosToCC_ORIG(i))
		next
		
		
		CreateAuditLogEntry "Global Settings Change", "Global Settings Change", "Major", 1, GetTerm("Accounts Receivable") & " Need To Know Reports users to Cc changed from " & userNamesOrig & " to " & userNames
	End If

		
	If Request.Form("chkN2KARReportONOFF")="on" then N2KARReportONOFFMsg = "On" Else N2KARReportONOFFMsg = "Off"
	If N2KARReportONOFFMsg_ORIG = 1 then N2KARReportONOFFMsg_ORIGFMsg = "On" Else N2KARReportONOFFMsg_ORIGFMsg = "Off"
	
	IF N2KARReportONOFF <> N2KARReportONOFFMsg_ORIG Then
		CreateAuditLogEntry "Global Settings Change", "Global Settings Change", "Major", 1, GetTerm("Accounts Receivable") & " need to know report changed from " & N2KARReportONOFFMsg_ORIGFMsg & " to " & N2KARReportONOFFMsg 
	End If


	If Request.Form("chkEmptyAddress2") = "on" then EmptyAddress2Msg = "On" Else EmptyAddress2Msg = "Off"
	If N2KARIncludeEmptyAddress2_ORIG = 1 then N2KARIncludeEmptyAddress2_ORIGMsg = "On" Else N2KARIncludeEmptyAddress2_ORIGMsg = "Off"

	IF EmptyAddress2 <> N2KARIncludeEmptyAddress2_ORIG Then
		CreateAuditLogEntry "Global Settings Change", "Global Settings Change", "Major", 1, GetTerm("Accounts Receivable") & " need to know report Emtry Address changed from " & N2KARIncludeEmptyAddress2_ORIGMsg & " to " & EmptyAddress2Msg 
	End If

	If Request.Form("chkEmptyCity") = "on" then EmptyCityMsg = "On" Else EmptyCityMsg = "Off"
	If N2KARIncludeEmptyCity_ORIG = 1 then N2KARIncludeEmptyCity_ORIGMsg = "On" Else N2KARIncludeEmptyCity_ORIGMsg = "Off"

	IF EmptyCity <> N2KARIncludeEmptyCity_ORIG Then
		CreateAuditLogEntry "Global Settings Change", "Global Settings Change", "Major", 1, GetTerm("Accounts Receivable") & " need to know report Empty City changed from " & N2KARIncludeEmptyCity_ORIGMsg & " to " & EmptyCityMsg 
	End If

	If Request.Form("chkEmptyCityStateZip") = "on" then EmptyCityStateZipMsg = "On" Else EmptyCityStateZipMsg = "Off"
	If N2KARIncludeEmptyCityStateZip_ORIG = 1 then N2KARIncludeEmptyCityStateZip_ORIGMsg = "On" Else N2KARIncludeEmptyCityStateZip_ORIGMsg = "Off"

	IF EmptyCityStateZip <> N2KARIncludeEmptyCityStateZip_ORIG Then
		CreateAuditLogEntry "Global Settings Change", "Global Settings Change", "Major", 1, GetTerm("Accounts Receivable") & " need to know report Empty City State Zip changed from " & N2KARIncludeEmptyCityStateZip_ORIGMsg & " to " & EmptyCityStateZipMsg 
	End If

	If Request.Form("chkEmptyCustomerName") = "on" then EmptyCustomerNameMsg = "On" Else EmptyCustomerNameMsg = "Off"
	If N2KARIncludeEmptyCustomerName_ORIG = 1 then N2KARIncludeEmptyCustomerName_ORIGMsg = "On" Else N2KARIncludeEmptyCustomerName_ORIGMsg = "Off"

	IF EmptyCustomerName <> N2KARIncludeEmptyCustomerName_ORIG Then
		CreateAuditLogEntry "Global Settings Change", "Global Settings Change", "Major", 1, GetTerm("Accounts Receivable") & " need to know report Emptry Customer Name changed from " & N2KARIncludeEmptyCustomerName_ORIGMsg & " to " & EmptyCustomerNameMsg 
	End If

	If Request.Form("chkEmptyPhoneNumber") = "on" then EmptyPhoneNumberMsg = "On" Else EmptyPhoneNumberMsg = "Off"
	If N2KARIncludeEmptyPhoneNumber_ORIG = 1 then N2KARIncludeEmptyPhoneNumber_ORIGMsg = "On" Else N2KARIncludeEmptyPhoneNumber_ORIGMsg = "Off"

	IF EmptyPhoneNumber <> N2KARIncludeEmptyPhoneNumber_ORIG Then
		CreateAuditLogEntry "Global Settings Change", "Global Settings Change", "Major", 1, GetTerm("Accounts Receivable") & " need to know report Empty Phone Number changed from " & N2KARIncludeEmptyPhoneNumber_ORIGMsg & " to " & EmptyPhoneNumberMsg 
	End If

	If Request.Form("chkEmptyState") = "on" then EmptyStateMsg = "On" Else EmptyStateMsg = "Off"
	If N2KARIncludeEmptyState_ORIG = 1 then N2KARIncludeEmptyState_ORIGMsg = "On" Else N2KARIncludeEmptyState_ORIGMsg = "Off"

	IF EmptyState <> N2KARIncludeEmptyState_ORIG Then
		CreateAuditLogEntry "Global Settings Change", "Global Settings Change", "Major", 1, GetTerm("Accounts Receivable") & " need to know report Empty State changed from " & N2KARIncludeEmptyState_ORIGMsg & " to " & EmptyStateMsg 
	End If

	If Request.Form("chkEmptyZip") = "on" then EmptyZipMsg = "On" Else EmptyZipMsg = "Off"
	If N2KARIncludeEmptyZip_ORIG = 1 then N2KARIncludeEmptyZip_ORIGMsg = "On" Else N2KARIncludeEmptyZip_ORIGMsg = "Off"

	IF EmptyZip <> N2KARIncludeEmptyZip_ORIG Then
		CreateAuditLogEntry "Global Settings Change", "Global Settings Change", "Major", 1, GetTerm("Accounts Receivable") & " need to know report Empty Zip changed from " & N2KARIncludeEmptyZip_ORIGMsg & " to " & EmptyZipMsg 
	End If

	If Request.Form("chkInvalidCityStateZip") = "on" then InvalidCityStateZipMsg = "On" Else InvalidCityStateZipMsg = "Off"
	If N2KARIncludeInvalidCityStateZip_ORIG = 1 then N2KARIncludeInvalidCityStateZip_ORIGMsg = "On" Else N2KARIncludeInvalidCityStateZip_ORIGMsg = "Off"

	IF InvalidCityStateZip <> N2KARIncludeInvalidCityStateZip_ORIG Then
		CreateAuditLogEntry "Global Settings Change", "Global Settings Change", "Major", 1, GetTerm("Accounts Receivable") & " need to know report Invalid CIty State Zip changed from " & N2KARIncludeInvalidCityStateZip_ORIGMsg & " to " & InvalidCityStateZipMsg 
	End If

	If Request.Form("chkInvalidPhoneNumber") = "on" then InvalidPhoneNumberMsg = "On" Else InvalidPhoneNumberMsg = "Off"
	If N2KARIncludeInvalidPhoneNumber_ORIG = 1 then N2KARIncludeInvalidPhoneNumber_ORIGMsg = "On" Else N2KARIncludeInvalidPhoneNumber_ORIGMsg = "Off"

	IF InvalidPhoneNumber <> N2KARIncludeInvalidPhoneNumber_ORIG Then
		CreateAuditLogEntry "Global Settings Change", "Global Settings Change", "Major", 1, GetTerm("Accounts Receivable") & " need to know report Invalid Phone Number changed from " & N2KARIncludeInvalidPhoneNumber_ORIGMsg & " to " & InvalidPhoneNumberMsg 
	End If

	If Request.Form("chkInvalidState") = "on" then InvalidStateMsg = "On" Else InvalidStateMsg = "Off"
	If N2KARIncludeInvalidState_ORIG = 1 then N2KARIncludeInvalidState_ORIGMsg = "On" Else N2KARIncludeInvalidState_ORIGMsg = "Off"

	IF InvalidState <> N2KARIncludeInvalidState_ORIG Then
		CreateAuditLogEntry "Global Settings Change", "Global Settings Change", "Major", 1, GetTerm("Accounts Receivable") & " need to know report Invalid State changed from " & N2KARIncludeInvalidState_ORIGMsg & " to " & InvalidStateMsg 
	End If

	If Request.Form("chkInvalidZipCode") = "on" then InvalidZipCodeMsg = "On" Else InvalidZipCodeMsg = "Off"
	If N2KARIncludeInvalidZipCode_ORIG = 1 then N2KARIncludeInvalidZipCode_ORIGMsg = "On" Else N2KARIncludeInvalidZipCode_ORIGMsg = "Off"

	IF InvalidZipCode <> N2KARIncludeInvalidZipCode_ORIG Then
		CreateAuditLogEntry "Global Settings Change", "Global Settings Change", "Major", 1, GetTerm("Accounts Receivable") & " need to know report Invalid Zip Code changed from " & N2KARIncludeInvalidZipCode_ORIGMsg & " to " & InvalidZipCodeMsg 
	End If

	If Request.Form("chkMissingcustomertype") = "on" then MissingcustomertypeMsg = "On" Else MissingcustomertypeMsg = "Off"
	If N2KARIncludeMissingcustomertype_ORIG = 1 then N2KARIncludeMissingcustomertype_ORIGMsg = "On" Else N2KARIncludeMissingcustomertype_ORIGMsg = "Off"

	IF Missingcustomertype <> N2KARIncludeMissingcustomertype_ORIG Then
		CreateAuditLogEntry "Global Settings Change", "Global Settings Change", "Major", 1, GetTerm("Accounts Receivable") & " need to know report Missing Customer Type changed from " & N2KARIncludeMissingcustomertype_ORIGMsg & " to " & MissingcustomertypeMsg 
	End If


	If Request.Form("chkMissingprimarysalesman") = "on" then MissingprimarysalesmanMsg = "On" Else MissingprimarysalesmanMsg = "Off"
	If N2KARIncludeMissingprimarysalesman_ORIG = 1 then N2KARIncludeMissingprimarysalesman_ORIGMsg = "On" Else N2KARIncludeMissingprimarysalesman_ORIGMsg = "Off"
	
	IF Missingprimarysalesman <> N2KARIncludeMissingprimarysalesman_ORIG Then
		CreateAuditLogEntry "Global Settings Change", "Global Settings Change", "Major", 1, GetTerm("Accounts Receivable") & " need to know report Missing Primary Salesman changed from " & N2KARIncludeMissingprimarysalesman_ORIGMsg & " to " & MissingprimarysalesmanMsg 
	End If

	If Request.Form("chkMissingsecondarysalesman") = "on" then MissingsecondarysalesmanMsg = "On" Else MissingsecondarysalesmanMsg = "Off"	
	If N2KARIncludeMissingsecondarysalesman_ORIG = 1 then N2KARIncludeMissingsecondarysalesman_ORIGMsg = "On" Else N2KARIncludeMissingsecondarysalesman_ORIGMsg = "Off"

	IF Missingsecondarysalesman <> N2KARIncludeMissingsecondarysalesman_ORIG Then
		CreateAuditLogEntry "Global Settings Change", "Global Settings Change", "Major", 1, GetTerm("Accounts Receivable") & " need to know report Missing Secondry Salesman changed from " & N2KARIncludeMissingsecondarysalesman_ORIGMsg & " to " & MissingsecondarysalesmanMsg 
	End If

	If Request.Form("chkN2KARReportONOFF") = "on" then N2KARReportONOFF = 1 Else N2KARReportONOFF = 0
	Response.Redirect("accounts-receivable.asp")
%><!--#include file="../../../../inc/footer-main.asp"-->