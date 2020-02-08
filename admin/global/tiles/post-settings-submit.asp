<!--#include file="../../../inc/header.asp"-->

<%

	'***********************************************************
	'Get Values Of All Form Fields Posted
	'***********************************************************
		
	POST_Serno = Request.Form("txtSerno")
	POST_Mode = Request.Form("selMode")
	EmailForNon200Responses = Request.Form("txtEmailForNon200Responses")
	POST_ServiceMemoURL1 = Request.Form("txtServiceMemoURL1")
	POST_AssetLocationURL1 = Request.Form("txtAssetLocationURL1")
	POST_ServiceMemoURL2 = Request.Form("txtServiceMemoURL2")
	POST_AssetLocationURL2 = Request.Form("txtAssetLocationURL2")
	InternalEmail_MailDomain = Request.Form("txtInternalEmail_MailDomain")
	NeverPutOnHold = Request.Form("chkNeverPutOnHold")
	POST_ServiceMemoURL1ONOFF = Request.Form("chkPOST_ServiceMemoURL1ONOFF")
	POST_ServiceMemoURL1_MplexFormat = Request.Form("chkPOST_ServiceMemoURL1_MplexFormat")
	POST_AssetLocationURL1ONOFF = Request.Form("chkPOST_AssetLocationURL1ONOFF")
	POST_ServiceMemoURL2ONOFF = Request.Form("chkPOST_ServiceMemoURL2ONOFF")
	POST_AssetLocationURL2ONOFF = Request.Form("chkPOST_AssetLocationURL2ONOFF")
	
	
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
		POST_Serno_ORIG = rs("POST_Serno")
		POST_Mode_ORIG = rs("POST_Mode")
		EmailForNon200Responses_ORIG = rs("EmailForNon200Responses")
		POST_ServiceMemoURL1_ORIG = rs("POST_ServiceMemoURL1")
		POST_AssetLocationURL1_ORIG = rs("POST_AssetLocationURL1")
		POST_ServiceMemoURL2_ORIG = rs("POST_ServiceMemoURL2")
		POST_AssetLocationURL2_ORIG = rs("POST_AssetLocationURL2")
		InternalEmail_MailDomain_ORIG = rs("InternalEmail_MailDomain")
		ShowOpenPopupMessage_ORIG =	rs("NotesScreenShowPopup")
		NeverPutOnHold_ORIG = rs("NeverPutOnHold")
		POST_ServiceMemoURL1ONOFF_ORIG = rs("POST_ServiceMemoURL1ONOFF")		
		POST_ServiceMemoURL1_MplexFormat_ORIG = rs("POST_ServiceMemoURL1_MplexFormat")			
		POST_AssetLocationURL1ONOFF_ORIG = rs("POST_AssetLocationURL1ONOFF")		
		POST_ServiceMemoURL2ONOFF_ORIG = rs("POST_ServiceMemoURL2ONOFF")		
		POST_AssetLocationURL2ONOFF_ORIG = rs("POST_AssetLocationURL2ONOFF")		
	End If
	
	set rs = Nothing
	cnn8.close
	set cnn8 = Nothing
	

	'***********************************************************
	'Perform Audit Trail Entry Inserts
	'***********************************************************
	
	
	'Post Parameter Audit Trail Entries
	If POST_Serno <> POST_Serno_ORIG Then
		CreateAuditLogEntry "Global Settings Change", "Global Settings Change", "Major", 1, "POST Serno changed from " & POST_Serno_ORIG & " to " & POST_Serno
	End If
	If POST_Mode <> POST_Mode_ORIG Then
		CreateAuditLogEntry "Global Settings Change", "Global Settings Change", "Major", 1, "POST Mode changed from " & POST_Mode_ORIG & " to " & POST_Mode
	End If
	If EmailForNon200Responses <> EmailForNon200Responses_ORIG Then
		CreateAuditLogEntry "Global Settings Change", "Global Settings Change", "Major", 1, "Email For Non 200 Responses changed from " & EmailForNon200Responses_ORIG & " to " & EmailForNon200Responses
	End If
	If POST_ServiceMemoURL <> POST_ServiceMemoURL1_ORIG Then
		CreateAuditLogEntry "Global Settings Change", "Global Settings Change", "Major", 1, "POST Service Memo URL#1 changed from " & POST_ServiceMemoURL1_ORIG & " to " & POST_ServiceMemoURL1 
	End If
	If POST_ServiceMemoURL2 <> POST_ServiceMemoURL2_ORIG Then
		CreateAuditLogEntry "Global Settings Change", "Global Settings Change", "Major", 1, "POST Service Memo URL#2 changed from " & POST_ServiceMemoURL2_ORIG & " to " & POST_ServiceMemoURL2
	End If
	If POST_AssetLocationURL1 <> POST_AssetLocationURL1_ORIG Then
		CreateAuditLogEntry "Global Settings Change", "Global Settings Change", "Major", 1, "POST Asset Location URL changed from " & POST_AssetLocationURL1_ORIG & " to " & POST_AssetLocationURL1 
	End If
	If POST_AssetLocationURL2 <> POST_AssetLocationURL2_ORIG Then
		CreateAuditLogEntry "Global Settings Change", "Global Settings Change", "Major", 1, "POST Asset Location URL changed from " & POST_AssetLocationURL2_ORIG & " to " & POST_AssetLocationURL2
	End If
	IF InternalEmail_MailDomain <> InternalEmail_MailDomain_ORIG Then
		CreateAuditLogEntry "Global Settings Change", "Global Settings Change", "Major", 1, "Internal email domain changed from " & InternalEmail_MailDomain_ORIG & " to " & InternalEmail_MailDomain
	End If
	If Request.Form("chkNeverPutOnHold")="on" then NeverPutOnHold = 1 Else NeverPutOnHold = 0
	IF NeverPutOnHold  <> NeverPutOnHold_ORIG Then
		CreateAuditLogEntry "Global Settings Change", "Global Settings Change", "Major", 1, "Never put tickets on hold changed from " & NeverPutOnHold_ORIG & " to " & NeverPutOnHold  
	End If
	If Request.Form("chkPOST_ServiceMemoURL1ONOFF") ="on" then POST_ServiceMemoURL1ONOFF = 1 Else POST_ServiceMemoURL1ONOFF = 0
	IF POST_ServiceMemoURL1ONOFF  <> POST_ServiceMemoURL1ONOFF_ORIG Then
		CreateAuditLogEntry "Global Settings Change", "Global Settings Change", "Major", 1, "POST Service Memo URL1 ON/OFF changed from " & POST_ServiceMemoURL1ONOFF_ORIG & " to " & POST_ServiceMemoURL1ONOFF  
	End If
	If Request.Form("chkPOST_ServiceMemoURL1_MplexFormat") ="on" then POST_ServiceMemoURL1_MplexFormat = 1 Else POST_ServiceMemoURL1_MplexFormat = 0
	IF POST_ServiceMemoURL1_MplexFormat  <> POST_ServiceMemoURL1_MplexFormat_ORIG Then
		CreateAuditLogEntry "Global Settings Change", "Global Settings Change", "Major", 1, "POST Service Memo URL1 use Metroplex format changed from " & POST_ServiceMemoURL1_MplexFormat_ORIG & " to " & POST_ServiceMemoURL1_MplexFormat  
	End If
	If Request.Form("chkPOST_AssetLocationURL1ONOFF")="on" then POST_AssetLocationURL1ONOFF = 1 Else POST_AssetLocationURL1ONOFF = 0
	IF POST_AssetLocationURL1ONOFF <> POST_AssetLocationURL1ONOFF_ORIG Then
		CreateAuditLogEntry "Global Settings Change", "Global Settings Change", "Major", 1, "POST Asset Location URL1 ON/OFF changed from " & POST_AssetLocationURL1ONOFF_ORIG & " to " & POST_AssetLocationURL1ONOFF
	End If
	If Request.Form("chkPOST_ServiceMemoURL2ONOFF")="on" then POST_ServiceMemoURL2ONOFF = 1 Else POST_ServiceMemoURL2ONOFF = 0
	IF POST_ServiceMemoURL2ONOFF <> POST_ServiceMemoURL2ONOFF_ORIG Then
		CreateAuditLogEntry "Global Settings Change", "Global Settings Change", "Major", 1, "POST Service Memo URL2 ON/OFF changed from " & POST_ServiceMemoURL2ONOFF_ORIG & " to " & POST_ServiceMemoURL2ONOFF  
	End If
	If Request.Form("chkPOST_AssetLocationURL2ONOFF")="on" then POST_AssetLocationURL2ONOFF = 1 Else POST_AssetLocationURL2ONOFF = 0
	IF POST_AssetLocationURL2ONOFF <> POST_AssetLocationURL2ONOFF_ORIG Then
		CreateAuditLogEntry "Global Settings Change", "Global Settings Change", "Major", 1, "POST Asset Location URL2 ON/OFF changed from " & POST_AssetLocationURL2ONOFF_ORIG & " to " & POST_AssetLocationURL2ONOFF
	End If


	'***********************************************************
	'Update SQL with Request Form Field Data
	'***********************************************************

	POST_ServiceMemoURL1 = trim(POST_ServiceMemoURL1)
	
	If left(POST_ServiceMemoURL1,1) = "," Then 
		POST_ServiceMemoURL1 = right(POST_ServiceMemoURL1,len(POST_ServiceMemoURL1)-1)
	End If
	
	POST_ServiceMemoURL1 = trim(POST_ServiceMemoURL1)

	SQL = "UPDATE  " & MUV_Read("SQL_Owner") & ".Settings_Global SET "
	SQL = SQL & "NeverPutOnHold = " & NeverPutOnHold & ", "
	SQL = SQL & "POST_Serno = '" & POST_Serno & "',"
	SQL = SQL & "POST_Mode = '" & POST_Mode & "',"
	SQL = SQL & "EmailForNon200Responses = '" & EmailForNon200Responses & "',"
	SQL = SQL & "POST_ServiceMemoURL1 = '" & POST_ServiceMemoURL1 & "',"
	SQL = SQL & "POST_AssetLocationURL1 = '" & POST_AssetLocationURL1 & "',"		
	SQL = SQL & "POST_ServiceMemoURL2 = '" & POST_ServiceMemoURL2 & "',"
	SQL = SQL & "POST_AssetLocationURL2 = '" & POST_AssetLocationURL2 & "',"
	SQL = SQL & "InternalEmail_MailDomain = '" & InternalEmail_MailDomain & "',"
	SQL = SQL & "POST_ServiceMemoURL1ONOFF = " & POST_ServiceMemoURL1ONOFF & ", "	
	SQL = SQL & "POST_ServiceMemoURL1_MplexFormat = " & POST_ServiceMemoURL1_MplexFormat & ", "	
	SQL = SQL & "POST_AssetLocationURL1ONOFF = " & POST_AssetLocationURL1ONOFF & ", "	
	SQL = SQL & "POST_ServiceMemoURL2ONOFF = " & POST_ServiceMemoURL2ONOFF & ", "	
	SQL = SQL & "POST_AssetLocationURL2ONOFF = " & POST_AssetLocationURL2ONOFF & " "	
	

	Set cnn8 = Server.CreateObject("ADODB.Connection")
	cnn8.open (Session("ClientCnnString"))
	Set rs = Server.CreateObject("ADODB.Recordset")
	rs.CursorLocation = 3 
	Set rs = cnn8.Execute(SQL)
	
	set rs = Nothing
	cnn8.close
	set cnn8 = Nothing

	
		
	Response.Redirect("post-settings.asp")
%>
<!--#include file="../../../inc/footer-main.asp"-->