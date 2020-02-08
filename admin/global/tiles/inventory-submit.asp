<!--#include file="../../../inc/header.asp"-->

<%

	'***********************************************************
	'Get Values Of All Form Fields Posted
	'***********************************************************
	
	InventoryAPIRepostURL = Request.Form("txtInventoryAPIRepostURL")
	If Request.Form("chkInventoryOrderAPIRepostONOFF") = "on" then InventoryAPIRepostONOFF = 1 Else InventoryAPIRepostONOFF = 0
	InventoryAPIRepostMode = Request.Form("selInventoryOrderAPIRepostMode")

	InventoryAPIRepostOnHandURL = Request.Form("txtInventoryAPIRepostOnHandURL")
	If Request.Form("chkInventoryAPIRepostOnHandONOFF") = "on" then InventoryAPIRepostOnHandONOFF = 1 Else InventoryAPIRepostOnHandONOFF = 0
	InventoryAPIRepostOnHandMode = Request.Form("selInventoryOrderAPIRepostOnHandMode")
	InventoryWebAppPostOnHandURL	= Request.Form("txtInventoryWebAppPostOnHandURL")
	InventoryWebAppPostOnHandMode= Request.Form("selInventoryWebAppPostOnHandMode")

	If Request.Form("chkInventoryAPIDailyActivityReportOnOff") = "on" then InventoryAPIDailyActivityReportOnOff = 1 Else InventoryAPIDailyActivityReportOnOff = 0
	InventoryAPIDailyActivityReportAdditionalEmails = Request.Form("txtInventoryAPIDailyActivityReportAdditionalEmails")
	InventoryAPIDailyActivityReportEmailSubject = Request.Form("txtInventoryAPIDailyActivityReportEmailSubject")
	InventoryAPIDailyActivityReportUserNos = Request.Form("lstSelectedInventoryAPIDailyActivityReportUserIDs")

	If Request.Form("chkInventoryProductChangesReportOnOff") = "on" then InventoryProductChangesReportOnOff = 1 Else InventoryProductChangesReportOnOff = 0
	InventoryProductChangesReportAdditionalEmails = Request.Form("txtInventoryProductChangesReportAdditionalEmails")
	InventoryProductChangesReportEmailSubject = Request.Form("txtInventoryProductChangesReportEmailSubject")
	InventoryProductChangesReportUserNos = Request.Form("lstSelectedInventoryProductChangesReportUserIDs")

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
		InventoryAPIRepostONOFF_ORIG = rs("InventoryAPIRepostONOFF")			
		InventoryAPIRepostMode_ORIG = rs("InventoryAPIRepostMode")	
		InventoryAPIRepostURL_ORIG = rs("InventoryAPIRepostURL")	
		InventoryAPIDailyActivityReportOnOff_ORIG	= rs("InventoryAPIDailyActivityReportOnOff")			
		InventoryAPIDailyActivityReportAdditionalEmails_ORIG = rs("InventoryAPIDailyActivityReportAdditionalEmails")	
		InventoryAPIDailyActivityReportEmailSubject_ORIG = rs("InventoryAPIDailyActivityReportEmailSubject")	
		InventoryAPIDailyActivityReportUserNos_ORIG = rs("InventoryAPIDailyActivityReportUserNos")		
		InventoryAPIRepostOnHandONOFF_ORIG = rs("InventoryAPIRepostOnHandONOFF")			
		InventoryAPIRepostOnHandMode_ORIG = rs("InventoryAPIRepostOnHandMode")	
		InventoryAPIRepostOnHandURL_ORIG = rs("InventoryAPIRepostOnHandURL")	
		InventoryWebAppPostOnHandURL_ORIG = rs("InventoryWebAppPostOnHandURL")
		InventoryWebAppPostOnHandMode_ORIG = rs("InventoryWebAppPostOnHandMode")
	End If


	SQL = "SELECT * FROM Settings_InventoryControl"
	
	Set cnn8 = Server.CreateObject("ADODB.Connection")
	cnn8.open (Session("ClientCnnString"))
	Set rs = Server.CreateObject("ADODB.Recordset")
	rs.CursorLocation = 3 
	Set rs = cnn8.Execute(SQL)
	
	If not rs.EOF Then
		InventoryProductChangesReportOnOff_ORIG = rs("InventoryProductChangesReportOnOff")			
		InventoryProductChangesReportAdditionalEmails_ORIG = rs("InventoryProductChangesReportAdditionalEmails")	
		InventoryProductChangesReportEmailSubject_ORIG = rs("InventoryProductChangesReportEmailSubject")	
		InventoryProductChangesReportUserNos_ORIG = rs("InventoryProductChangesReportUserNos")	
	End If

	
	set rs = Nothing
	cnn8.close
	set cnn8 = Nothing
	

	'***********************************************************
	'Update SQL with Request Form Field Data
	'***********************************************************

	SQL = "UPDATE  " & MUV_Read("SQL_Owner") & ".Settings_Global SET  "
	SQL = SQL & "InventoryAPIDailyActivityReportOnOff = " & InventoryAPIDailyActivityReportOnOff & ","
	SQL = SQL & "InventoryAPIDailyActivityReportAdditionalEmails = '" & InventoryAPIDailyActivityReportAdditionalEmails & "',"
	SQL = SQL & "InventoryAPIDailyActivityReportEmailSubject = '" & InventoryAPIDailyActivityReportEmailSubject & "',"
	SQL = SQL & "InventoryAPIDailyActivityReportUserNos = '" & InventoryAPIDailyActivityReportUserNos & "',"
	SQL = SQL & "InventoryAPIRepostONOFF = " & InventoryAPIRepostONOFF & ","
	SQL = SQL & "InventoryAPIRepostURL = '" & InventoryAPIRepostURL & "',"
	SQL = SQL & "InventoryAPIRepostMode = '" & InventoryAPIRepostMode & "', "
	SQL = SQL & "InventoryAPIRepostOnHandONOFF = " & InventoryAPIRepostOnHandONOFF & ","
	SQL = SQL & "InventoryAPIRepostOnHandURL = '" & InventoryAPIRepostOnHandURL & "',"
	SQL = SQL & "InventoryAPIRepostOnHandMode = '" & InventoryAPIRepostOnHandMode & "', "
	SQL = SQL & "InventoryWebAppPostOnHandMode = '" & InventoryWebAppPostOnHandMode & "', "
	SQL = SQL & "InventoryWebAppPostOnHandURL = '" & InventoryWebAppPostOnHandURL & "' "

	
	'Response.write("<br><br><br>" & SQL)
	
	Set cnn8 = Server.CreateObject("ADODB.Connection")
	cnn8.open (Session("ClientCnnString"))
	Set rs = Server.CreateObject("ADODB.Recordset")
	rs.CursorLocation = 3 
	
	Set rs = cnn8.Execute(SQL)

	SQL = "UPDATE  " & MUV_Read("SQL_Owner") & ".Settings_InventoryControl SET  "
	SQL = SQL & "InventoryProductChangesReportOnOff = " & InventoryProductChangesReportOnOff & ","
	SQL = SQL & "InventoryProductChangesReportAdditionalEmails = '" & InventoryProductChangesReportAdditionalEmails & "',"
	SQL = SQL & "InventoryProductChangesReportEmailSubject = '" & InventoryProductChangesReportEmailSubject & "',"
	SQL = SQL & "InventoryProductChangesReportUserNos = '" & InventoryProductChangesReportUserNos & "' "
	
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


	If Request.Form("chkInventoryOrderAPIRepostONOFF")="on" then InventoryAPIRepostONOFF = "On" Else InventoryAPIRepostONOFF = "Off"

	If InventoryAPIRepostONOFF <> InventoryAPIRepostONOFF_ORIG Then
		CreateAuditLogEntry "Global Settings Change", "Global Settings Change", "Major", 1, "Inventory API Post Mode changed from " & InventoryAPIRepostONOFF_ORIG & " to " & InventoryAPIRepostONOFF
	End If
	
	If InventoryAPIRepostMode <> InventoryAPIRepostMode_ORIG Then
		CreateAuditLogEntry "Global Settings Change", "Global Settings Change", "Major", 1, "Inventory API posting URL changed from " & InventoryAPIRepostMode_ORIG & " to " & InventoryAPIRepostMode
	End If

	If InventoryAPIRepostURL <> InventoryAPIRepostURL_ORIG Then
		CreateAuditLogEntry "Global Settings Change", "Global Settings Change", "Major", 1, "Inventory API posting URL changed from " & InventoryAPIRepostURL_ORIG & " to " & InventoryAPIRepostURL 
	End If
		
	If Request.Form("chkInventoryOrderAPIRepostOnHandONOFF")="on" then InventoryAPIRepostOnHandONOFF = "On" Else InventoryAPIRepostOnHandONOFF = "Off"

	If InventoryAPIRepostOnHandONOFF <> InventoryAPIRepostOnHandONOFF_ORIG Then
		CreateAuditLogEntry "Global Settings Change", "Global Settings Change", "Major", 1, "Inventory API Post On Hand Mode changed from " & InventoryAPIRepostOnHandONOFF_ORIG & " to " & InventoryAPIRepostOnHandONOFF
	End If
	
	If InventoryAPIRepostOnHandMode <> InventoryAPIRepostOnHandMode_ORIG Then
		CreateAuditLogEntry "Global Settings Change", "Global Settings Change", "Major", 1, "Inventory API posting On Hand MODE changed from " & InventoryAPIRepostOnHandMode_ORIG & " to " & InventoryAPIRepostOnHandMode
	End If

	If InventoryAPIRepostOnHandURL <> InventoryAPIRepostOnHandURL_ORIG Then
		CreateAuditLogEntry "Global Settings Change", "Global Settings Change", "Major", 1, "Inventory API posting On Hand URL changed from " & InventoryAPIRepostOnHandURL_ORIG & " to " & InventoryAPIRepostOnHandURL 
	End If

	If InventoryWebAppPostOnHandMode <> InventoryWebAppPostOnHandMode_ORIG Then
		CreateAuditLogEntry "Global Settings Change", "Global Settings Change", "Major", 1, "Inventory API posting On Hand MODE changed from " & InventoryWebAppPostOnHandMode_ORIG & " to " & InventoryWebAppPostOnHandMode
	End If

	If InventoryWebAppPostOnHandURL<> InventoryWebAppPostOnHandURL_ORIG Then
		CreateAuditLogEntry "Global Settings Change", "Global Settings Change", "Major", 1, "Inventory API posting On Hand URL changed from " & InventoryWebAppPostOnHandURL_ORIG & " to " & InventoryWebAppPostOnHandURL
	End If


	If Request.Form("chkInventoryAPIDailyActivityReportOnOff")="on" then InventoryAPIDailyActivityReportOnOff = "On" Else InventoryAPIDailyActivityReportOnOff = "Off"
	
	IF InventoryAPIDailyActivityReportOnOff <> InventoryAPIDailyActivityReportOnOff_ORIG Then
		CreateAuditLogEntry "Global Settings Change", "Global Settings Change", "Major", 1, "Email Inventory API Daily Activity Report changed from " & InventoryAPIDailyActivityReportOnOff_ORIG & " to " & InventoryAPIDailyActivityReportOnOff
	End If

	If InventoryAPIDailyActivityReportEmailSubject <> InventoryAPIDailyActivityReportEmailSubject_ORIG Then
		CreateAuditLogEntry "Global Settings Change", "Global Settings Change", "Major", 1, "Inventory API Daily Activity Report Email Subject changed from " & InventoryAPIDailyActivityReportEmailSubject_ORIG & " to " & InventoryAPIDailyActivityReportEmailSubject
	End If

	If InventoryAPIDailyActivityReportAdditionalEmails <> InventoryAPIDailyActivityReportAdditionalEmails_ORIG Then
		CreateAuditLogEntry "Global Settings Change", "Global Settings Change", "Major", 1, "Inventory API Daily Activity Report Additional Emails changed from " & InventoryAPIDailyActivityReportAdditionalEmails_ORIG & " to " & InventoryAPIDailyActivityReportAdditionalEmails
	End If


	If InventoryAPIDailyActivityReportUserNos <> InventoryAPIDailyActivityReportUserNos_ORIG Then

		userNames = ""
		userNamesOrig = ""
		
	
		IndividualInventoryAPIDailyActivityReportUserNos = Split(InventoryAPIDailyActivityReportUserNos,",")
		
		for i=0 to Ubound(IndividualInventoryAPIDailyActivityReportUserNos)
		     userNames = userNames & " " & GetUserFirstAndLastNameByUserNo(IndividualInventoryAPIDailyActivityReportUserNos(i))
		next
		
		
		IndividualInventoryAPIDailyActivityReportUserNos_ORIG  = Split(InventoryAPIDailyActivityReportUserNos_ORIG,",")
		
		for i=0 to Ubound(IndividualInventoryAPIDailyActivityReportUserNos_ORIG)
		     userNamesOrig = userNamesOrig & " " & GetUserFirstAndLastNameByUserNo(IndividualInventoryAPIDailyActivityReportUserNos_ORIG(i))
		next
		
		
		CreateAuditLogEntry "Global Settings Change", "Global Settings Change", "Major", 1, "Inventory API Daily Activity Report Additional Send Report To changed from " & userNamesOrig & " to " & userNames
	End If



	If Request.Form("chkInventoryProductChangesReportOnOff")="on" then InventoryProductChangesReportOnOff = "On" Else InventoryProductChangesReportOnOff = "Off"
	
	IF InventoryProductChangesReportOnOff <> InventoryProductChangesReportOnOff_ORIG Then
		CreateAuditLogEntry "Global Settings Change", "Global Settings Change", "Major", 1, "Email Inventory Product Changes Report changed from " & InventoryProductChangesReportOnOff_ORIG & " to " & InventoryProductChangesReportOnOff
	End If

	If InventoryProductChangesReportEmailSubject <> InventoryProductChangesReportEmailSubject_ORIG Then
		CreateAuditLogEntry "Global Settings Change", "Global Settings Change", "Major", 1, "Inventory Product Changes Report Email Subject changed from " & InventoryProductChangesReportEmailSubject_ORIG & " to " & InventoryProductChangesReportEmailSubject
	End If

	If InventoryProductChangesReportAdditionalEmails <> InventoryProductChangesReportAdditionalEmails_ORIG Then
		CreateAuditLogEntry "Global Settings Change", "Global Settings Change", "Major", 1, "Inventory Product Changes Report Additional Emails changed from " & InventoryProductChangesReportAdditionalEmails_ORIG & " to " & InventoryProductChangesReportAdditionalEmails
	End If

	If InventoryProductChangesReportUserNos <> InventoryProductChangesReportUserNos_ORIG Then

		userNames = ""
		userNamesOrig = ""
	
		IndividualInventoryProductChangesReportUserNos = Split(InventoryProductChangesReportUserNos,",")
		
		for i=0 to Ubound(IndividualInventoryProductChangesReportUserNos)
		     userNames = userNames & " " & GetUserFirstAndLastNameByUserNo(IndividualInventoryProductChangesReportUserNos(i))
		next
		
		IndividualInventoryProductChangesReportUserNos_ORIG  = Split(IndividualInventoryProductChangesReportUserNos_ORIG,",")
		
		for i=0 to Ubound(IndividualInventoryProductChangesReportUserNos_ORIG)
		     userNamesOrig = userNamesOrig & " " & GetUserFirstAndLastNameByUserNo(IndividualInventoryProductChangesReportUserNos_ORIG(i))
		next
		
		CreateAuditLogEntry "Global Settings Change", "Global Settings Change", "Major", 1, "Inventory Product Changes Report Additional Send Report To changed from " & userNamesOrig & " to " & userNames
	End If
	


	Response.Redirect("inventory.asp")
	
%><!--#include file="../../../inc/footer-main.asp"-->