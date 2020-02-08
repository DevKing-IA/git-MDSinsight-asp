<!--#include file="../../../../inc/header.asp"-->

<%

	'***********************************************************
	'Get Values Of All Form Fields Posted
	'***********************************************************
		
	POST_Serno = Request.Form("txtSerno")
	POST_Mode = Request.Form("selMode")
	EmailForNon200Responses = Request.Form("txtEmailForNon200Responses")
	POST_CustomerURL1 = Request.Form("txtCustomerURL1")
	POST_CustomerURL2 = Request.Form("txtCustomerURL2")
	POST_CustomerURL1ONOFF = Request.Form("chkPOST_CustomerURL1ONOFF")
	POST_CustomerURL2ONOFF = Request.Form("chkPOST_CustomerURL2ONOFF")
	
	
	'***********************************************************
	'Get Original Values For Audit Trail Entries
	'***********************************************************
		
	SQL = "SELECT * FROM Settings_AR"
	
	Set cnn8 = Server.CreateObject("ADODB.Connection")
	cnn8.open (Session("ClientCnnString"))
	Set rs = Server.CreateObject("ADODB.Recordset")
	rs.CursorLocation = 3 
	Set rs = cnn8.Execute(SQL)
	
	If not rs.EOF Then
		POST_Serno_ORIG = rs("POST_Serno")
		POST_Mode_ORIG = rs("POST_Mode")
		EmailForNon200Responses_ORIG = rs("EmailForNon200Responses")
		POST_CustomerURL1_ORIG = rs("POST_CustomerURL1")
		POST_CustomerURL2_ORIG = rs("POST_CustomerURL2")		
		POST_CustomerURL1ONOFF_ORIG = rs("POST_CustomerURL1ONOFF")		
		POST_CustomerURL2ONOFF_ORIG = rs("POST_CustomerURL2ONOFF")
	End If
	
	set rs = Nothing
	cnn8.close
	set cnn8 = Nothing
	

	'***********************************************************
	'Perform Audit Trail Entry Inserts
	'***********************************************************
	
	
	'Post Parameter Audit Trail Entries
	If POST_Serno <> POST_Serno_ORIG Then
		CreateAuditLogEntry "Global Settings AR API Change", "Global Settings AR API Change", "Major", 1, "POST Serno changed from " & POST_Serno_ORIG & " to " & POST_Serno
	End If
	If POST_Mode <> POST_Mode_ORIG Then
		CreateAuditLogEntry "Global Settings AR API Change", "Global Settings AR API Change", "Major", 1, "POST Mode changed from " & POST_Mode_ORIG & " to " & POST_Mode
	End If
	If EmailForNon200Responses <> EmailForNon200Responses_ORIG Then
		CreateAuditLogEntry "Global Settings AR API Change", "Global Settings AR API Change", "Major", 1, "Email For Non 200 Responses changed from " & EmailForNon200Responses_ORIG & " to " & EmailForNon200Responses
	End If
	If POST_CustomerURL1 <> POST_CustomerURL1_ORIG Then
		CreateAuditLogEntry "Global Settings AR API Change", "Global Settings AR API Change", "Major", 1, "POST Customer URL changed from " & POST_CustomerURL1_ORIG & " to " & POST_CustomerURL1 
	End If
	If POST_CustomerURL2 <> POST_CustomerURL2_ORIG Then
		CreateAuditLogEntry "Global Settings AR API Change", "Global Settings AR API Change", "Major", 1, "POST Customer URL changed from " & POST_CustomerURL2_ORIG & " to " & POST_CustomerURL2
	End If
	If Request.Form("chkPOST_CustomerURL1ONOFF")="on" then POST_CustomerURL1ONOFF = 1 Else POST_CustomerURL1ONOFF = 0
	IF POST_CustomerURL1ONOFF <> POST_CustomerURL1ONOFF_ORIG Then
		CreateAuditLogEntry "Global Settings AR API Change", "Global Settings AR API Change", "Major", 1, "POST Customer URL1 ON/OFF changed from " & POST_CustomerURL1ONOFF_ORIG & " to " & POST_CustomerURL1ONOFF
	End If
	If Request.Form("chkPOST_CustomerURL2ONOFF")="on" then POST_CustomerURL2ONOFF = 1 Else POST_CustomerURL2ONOFF = 0
	IF POST_CustomerURL2ONOFF <> POST_CustomerURL2ONOFF_ORIG Then
		CreateAuditLogEntry "Global Settings AR API Change", "Global Settings AR API Change", "Major", 1, "POST Customer ULR2 ON/OFF changed from " & POST_CustomerURL2ONOFF_ORIG & " to " & POST_CustomerURL2ONOFF
	End If



	'***********************************************************
	'Update SQL with Request Form Field Data
	'***********************************************************

	SQL = "UPDATE  " & MUV_Read("SQL_Owner") & ".Settings_AR SET "
	SQL = SQL & "POST_Serno = '" & POST_Serno & "',"
	SQL = SQL & "POST_Mode = '" & POST_Mode & "',"
	SQL = SQL & "EmailForNon200Responses = '" & EmailForNon200Responses & "',"		
	SQL = SQL & "POST_CustomerURL1 = '" & POST_CustomerURL1 & "',"
	SQL = SQL & "POST_CustomerURL2 = '" & POST_CustomerURL2 & "',"
	SQL = SQL & "POST_CustomerURL1ONOFF = " & POST_CustomerURL1ONOFF & ", "	
	SQL = SQL & "POST_CustomerURL2ONOFF = " & POST_CustomerURL2ONOFF	

	Set cnn8 = Server.CreateObject("ADODB.Connection")
	cnn8.open (Session("ClientCnnString"))
	Set rs = Server.CreateObject("ADODB.Recordset")
	rs.CursorLocation = 3 
	Set rs = cnn8.Execute(SQL)
	
	set rs = Nothing
	cnn8.close
	set cnn8 = Nothing
		
	Response.Redirect("accounts-receivable.asp")
%>
<!--#include file="../../../../inc/footer-main.asp"-->