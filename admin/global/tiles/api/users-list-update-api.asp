<!--#include file="../../../../inc/header.asp"-->

<%

Dim userListName: userListName = Request.Form("userListName")
	'*************************************************************************
	'See if this is the first time entering data in Settings_Global
	'*************************************************************************
	
	SQL = "SELECT * FROM Settings_Global"
	
	Set cnn8 = Server.CreateObject("ADODB.Connection")
	cnn8.open (Session("ClientCnnString"))
	Set rs = Server.CreateObject("ADODB.Recordset")
	rs.CursorLocation = 3 
	Set rs = cnn8.Execute(SQL)
	
	If rs.EOF Then
		SettingsGlobalHasRecords = false
	Else
		SettingsGlobalHasRecords = true	
	End If
				
	set rs = Nothing
	cnn8.close
	set cnn8 = Nothing

'Just to insert NULL default values so further update query will not fail
If NOT SettingsGlobalHasRecords Then
    SQL = "INSERT INTO " & MUV_Read("SQL_Owner") & ".Settings_Global "
	SQL = SQL & " (APIDailyActivityReportAdditionalEmails) "
	SQL = SQL & " VALUES "
	SQL = SQL & " ('') "	
	
	Set cnn8 = Server.CreateObject("ADODB.Connection")
	cnn8.open (Session("ClientCnnString"))
	Set rs = Server.CreateObject("ADODB.Recordset")
	rs.CursorLocation = 3 
	
	Set rs = cnn8.Execute(SQL)

	set rs = Nothing
	cnn8.close
	set cnn8 = Nothing
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
		APIDailyActivityReportAdditionalEmails_ORIG = rs("APIDailyActivityReportAdditionalEmails")		
	End If	
	
	set rs = Nothing
	cnn8.close
	set cnn8 = Nothing


If userListName = "APIDailyActivityReportAdditionalEmails" Then
	APIDailyActivityReportAdditionalEmails = Request.Form("txtAPIDailyActivityReportAdditionalEmails")

	Set cnn8 = Server.CreateObject("ADODB.Connection")
	cnn8.open (Session("ClientCnnString"))
	Set rs = Server.CreateObject("ADODB.Recordset")
	rs.CursorLocation = 3 

	SQL = "UPDATE  " & MUV_Read("SQL_Owner") & ".Settings_Global SET "
	SQL = SQL & "APIDailyActivityReportAdditionalEmails = '" & APIDailyActivityReportAdditionalEmails & "'"
	Set rs = cnn8.Execute(SQL)
 	set rs = Nothing
	cnn8.close
	set cnn8 = Nothing

    If APIDailyActivityReportAdditionalEmails <> APIDailyActivityReportAdditionalEmails_ORIG Then
		CreateAuditLogEntry "Global Settings Change", "Global Settings Change", "Major", 1, "API Daily Activity Report Additional Emails changed from " & APIDailyActivityReportAdditionalEmails_ORIG & " to " & APIDailyActivityReportAdditionalEmails
	End If
    Response.Redirect("order-api.asp")
End If

%><!--#include file="../../../../inc/footer-main.asp"-->