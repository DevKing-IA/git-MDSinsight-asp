<!--#include file="../../../inc/header.asp"-->

<%

	'***********************************************************
	'Get Values Of All Form Fields Posted
	'***********************************************************
	
	ShowOpenPopupMessage = Request.Form("chkShowOpenPopupMessage")
	
	

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
		NotesScreenShowPopup_ORIG = rs("NotesScreenShowPopup")
	End If
			
	set rs = Nothing
	cnn8.close
	set cnn8 = Nothing
	
	
	'***********************************************************
	'Perform Audit Trail Entry Inserts
	'***********************************************************
	
	If Request.Form("chkShowOpenPopupMessage")="on" then ShowOpenPopupMessage = 1 Else ShowOpenPopupMessage = 0
	
	IF ShowOpenPopupMessage  <> NotesScreenShowPopup_ORIG Then
		CreateAuditLogEntry "Global Settings Change", "Global Settings Change", "Major", 1, "Show popup on notes screen for open tickets " & NotesScreenShowPopup_ORIG & " to " & ShowOpenPopupMessage  
	End If
	
	
	'***********************************************************
	'Update SQL with Request Form Field Data
	'***********************************************************
	
	SQL = "UPDATE  " & MUV_Read("SQL_Owner") & ".Settings_Global SET NotesScreenShowPopup = " & ShowOpenPopupMessage
	
									
	Set cnn8 = Server.CreateObject("ADODB.Connection")
	cnn8.open (Session("ClientCnnString"))
	Set rs = Server.CreateObject("ADODB.Recordset")
	rs.CursorLocation = 3 
	Set rs = cnn8.Execute(SQL)
	

	set rs = Nothing
	cnn8.close
	set cnn8 = Nothing
	
	
	
	
	Response.Redirect("client-care-settings.aspv")
%>
<!--#include file="../../../inc/footer-main.asp"-->