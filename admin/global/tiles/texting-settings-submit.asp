<!--#include file="../../../inc/header.asp"-->

<%

	'***********************************************************
	'Get Values Of All Form Fields Posted
	'***********************************************************
			
	EZTextID = Request.Form("txtEZTextingID")
	EZTextPassword = Request.Form("txtEZTextingPassword")



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
		EZTextID_ORIG = rs("EZTextingID")
		EZTextPassword_ORIG = rs("EZTextingPassword")	
	End If
	
	set rs = Nothing
	cnn8.close
	set cnn8 = Nothing
	
	
	'***********************************************************
	'Perform Audit Trail Entry Inserts
	'***********************************************************
	
	If EZTextID <> EZTextID_ORIG Then
		CreateAuditLogEntry "Global Settings Change", "Global Settings Change", "Major", 1, "Ez-Texting user id changed from  " & EZTextID_ORIG & " to " & EZTextID
	End If
	If EZTextPassword <> EZTextPassword_ORIG Then
		CreateAuditLogEntry "Global Settings Change", "Global Settings Change", "Major", 1, "Ez-Texting password changed from  " & EZTextPassword_ORIG & " to " & EZTextPassword
	End If



	'***********************************************************
	'Update SQL with Request Form Field Data
	'***********************************************************
		
	SQL = "UPDATE  " & MUV_Read("SQL_Owner") & ".Settings_Global SET "
	SQL = SQL & "EZTextingID = '" & EZTextID & "',"
	SQL = SQL & "EZTextingPassword = '" & EZTextPassword & "'"
									

	Set cnn8 = Server.CreateObject("ADODB.Connection")
	cnn8.open (Session("ClientCnnString"))
	Set rs = Server.CreateObject("ADODB.Recordset")
	rs.CursorLocation = 3 
	Set rs = cnn8.Execute(SQL)
	
	set rs = Nothing
	cnn8.close
	set cnn8 = Nothing



	Response.Redirect("texting-settings.asp")
%>
<!--#include file="../../../inc/footer-main.asp"-->