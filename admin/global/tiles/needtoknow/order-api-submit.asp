<!--#include file="../../../../inc/header.asp"-->
<!--#include file="../../../../inc/InSightFuncs.asp"-->
<%

	'***********************************************************
	'Get Values Of All Form Fields Posted
	'***********************************************************
	
	N2KAPIEmailToUserNos = Request.Form("lstSelectedN2KAPIEmailToUserNos")
	N2KAPIUserNosToCC = Request.Form("lstSelectedN2KAPIUserNosToCC")
	If Request.Form("chkN2KAPIReportONOFF") = "on" then N2KAPIReportONOFF = 1 Else N2KAPIReportONOFF = 0
		
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
		
		N2KAPIEmailToUserNos_ORIG = rs("N2KAPIEmailToUserNos")
		N2KAPIUserNosToCC_ORIG = rs("N2KAPIUserNosToCC")
		N2KAPIReportONOFF_ORIG = rs("N2KAPIReportONOFF")
		
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
		SQL = SQL & "N2KAPIEmailToUserNos = '" & N2KAPIEmailToUserNos & "',"
		SQL = SQL & "N2KAPIUserNosToCC = '" & N2KAPIUserNosToCC & "',"
		SQL = SQL & "N2KAPIReportONOFF = " & N2KAPIReportONOFF
	
	Else
	
		SQL = "INSERT INTO " & MUV_Read("SQL_Owner") & ".Settings_NeedToKnow "
		SQL = SQL & " (N2KAPIEmailToUserNos, N2KAPIUserNosToCC,N2KAPIReportONOFF) "
		SQL = SQL & " VALUES "
		SQL = SQL & " ('" & N2KAPIEmailToUserNos & "', '" & N2KAPIUserNosToCC & "', " & N2KAPIReportONOFF & ") "
	
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

	If N2KAPIEmailToUserNos <> N2KAPIEmailToUserNos_ORIG Then

		userNames = ""
		userNamesOrig = ""
		
	
		IndividualN2KAPIEmailToUserNos = Split(N2KAPIEmailToUserNos,",")
		
		for i=0 to Ubound(IndividualN2KAPIEmailToUserNos)
		     userNames = userNames & " " & GetUserFirstAndLastNameByUserNo(IndividualN2KAPIEmailToUserNos (i))
		next
		
		
		IndividualN2KAPIEmailToUserNos_ORIG  = Split(N2KAPIEmailToUserNos_ORIG,",")
		
		for i=0 to Ubound(IndividualN2KAPIEmailToUserNos_ORIG)
		     userNamesOrig = userNamesOrig & " " & GetUserFirstAndLastNameByUserNo(IndividualN2KAPIEmailToUserNos_ORIG(i))
		next
		
		
		CreateAuditLogEntry "Global Settings Change", "Global Settings Change", "Major", 1, "API Need To Know Reports send email to changed from " & userNamesOrig & " to " & userNames
	End If


	If N2KAPIUserNosToCC <> N2KAPIUserNosToCC_ORIG Then

		userNames = ""
		userNamesOrig = ""
		
	
		IndividualN2KAPIUserNosToCC  = Split(N2KAPIUserNosToCC,",")
		
		for i=0 to Ubound(IndividualN2KAPIUserNosToCC)
		     userNames = userNames & " " & GetUserFirstAndLastNameByUserNo(IndividualN2KAPIUserNosToCC (i))
		next
		
		
		IndividualN2KAPIUserNosToCC_ORIG  = Split(IndividualN2KAPIUserNosToCC_ORIG,",")
		
		for i=0 to Ubound(IndividualN2KAPIUserNosToCC_ORIG)
		     userNamesOrig = userNamesOrig & " " & GetUserFirstAndLastNameByUserNo(IndividualN2KAPIUserNosToCC_ORIG(i))
		next
		
		
		CreateAuditLogEntry "Global Settings Change", "Global Settings Change", "Major", 1, "API Need To Know Reports users to Cc changed from " & userNamesOrig & " to " & userNames
	End If


	If Request.Form("chkN2KAPIReportONOFF")="on" then N2KAPIReportONOFFMsg = "On" Else N2KAPIReportONOFFMsg = "Off"
	If N2KAPIReportONOFFMsg_ORIG = 1 then N2KAPIReportONOFFMsg_ORIGFMsg = "On" Else N2KAPIReportONOFFMsg_ORIGFMsg = "Off"
	
	IF N2KAPIReportONOFF <> N2KAPIReportONOFFMsg_ORIG Then
		CreateAuditLogEntry "Global Settings Change", "Global Settings Change", "Major", 1, "API to know report changed from " & N2KAPIReportONOFFMsg_ORIGFMsg & " to " & N2KAPIReportONOFFMsg 
	End If

	Response.Redirect("order-api.asp")
	
%><!--#include file="../../../../inc/footer-main.asp"-->