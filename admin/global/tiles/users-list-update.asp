<!--#include file="../../../inc/header.asp"-->

<%

Dim userListName: userListName = Request.Form("userListName")


	SQL = "SELECT * FROM Settings_Global"
	
	Set cnn8 = Server.CreateObject("ADODB.Connection")
	cnn8.open (Session("ClientCnnString"))
	Set rs = Server.CreateObject("ADODB.Recordset")
	rs.CursorLocation = 3 
	Set rs = cnn8.Execute(SQL)
	
	If not rs.EOF Then
		ProspSnapshotAdditionalEmails_ORIG = rs("ProspSnapshotAdditionalEmails")	
	End If	
	
	set rs = Nothing
	cnn8.close
	set cnn8 = Nothing
	
	

	SQL = "SELECT * FROM Settings_Prospecting"
	
	Set cnn8 = Server.CreateObject("ADODB.Connection")
	cnn8.open (Session("ClientCnnString"))
	Set rs = Server.CreateObject("ADODB.Recordset")
	rs.CursorLocation = 3 
	Set rs = cnn8.Execute(SQL)
	
	If not rs.EOF Then
		ProspectingWeeklyAgendaReportAdditionalEmails_ORIG = rs("ProspectingWeeklyAgendaReportAdditionalEmails")
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
		FieldServiceNotesReportAdditionalEmails_ORIG = rs("FieldServiceNotesReportAdditionalEmails")	
        ServiceTicketCarryoverReportAdditionalEmails_ORIG = rs("ServiceTicketCarryoverReportAdditionalEmails")	
        ServiceTicketthresholdReportAdditionalEmails_ORIG = rs("ServiceTicketthresholdReportAdditionalEmails")	
	End If	
	
	set rs = Nothing
	cnn8.close
	set cnn8 = Nothing
	
	
	If userListName = "InventoryProductChangesReportAdditionalEmails"  Then

	'*************************************************************************
	'See if this is the first time entering data in Settings_InventoryControl
	'*************************************************************************
	
	SQL = "SELECT * FROM Settings_InventoryControl"
	
	Set cnn8 = Server.CreateObject("ADODB.Connection")
	cnn8.open (Session("ClientCnnString"))
	Set rs = Server.CreateObject("ADODB.Recordset")
	rs.CursorLocation = 3 
	Set rs = cnn8.Execute(SQL)
	
	If rs.EOF Then
		SettingsInventoryControlHasRecords = false
	Else
		SettingsInventoryControlHasRecords = true	
	End If
				
	set rs = Nothing
	cnn8.close
	set cnn8 = Nothing
	
	

	'Just to insert NULL default values so further update query will not fail
	If NOT SettingsInventoryControlHasRecords Then
	    SQL = "INSERT INTO " & MUV_Read("SQL_Owner") & ".Settings_InventoryControl "
		SQL = SQL & " (InventoryAPIRepostONOFF,InventoryAPIRepostOnHandONOFF,InventoryAPIDailyActivityReportOnOff,InventoryProductChangesReportOnOff) "
		SQL = SQL & " VALUES "
		SQL = SQL & " (0,0,0,0) "	
		
		Set cnn8 = Server.CreateObject("ADODB.Connection")
		cnn8.open (Session("ClientCnnString"))
		Set rs = Server.CreateObject("ADODB.Recordset")
		rs.CursorLocation = 3 
		
		Set rs = cnn8.Execute(SQL)
	
		set rs = Nothing
		cnn8.close
		set cnn8 = Nothing
	End If
	
	    SQL = "SELECT * FROM Settings_InventoryControl"
		
		Set cnn8 = Server.CreateObject("ADODB.Connection")
		cnn8.open (Session("ClientCnnString"))
		Set rs = Server.CreateObject("ADODB.Recordset")
		rs.CursorLocation = 3 
		Set rs = cnn8.Execute(SQL)
		
		If not rs.EOF Then
	        InventoryProductChangesReportAdditionalEmails_ORIG = rs("InventoryProductChangesReportAdditionalEmails")
		End If	
		
		set rs = Nothing
		cnn8.close
		set cnn8 = Nothing
	End If

	
	If userListName = "CustAnalSum1EmailAddressesToCC" OR userListName = "MCSActivitySummaryEmailAddressesToCC" Then
	  	
	  	'*************************************************************************
		'See if this is the first time entering data in Settings_BizIntel
		'*************************************************************************
		
		SQL = "SELECT * FROM Settings_BizIntel"
		
		Set cnn8 = Server.CreateObject("ADODB.Connection")
		cnn8.open (Session("ClientCnnString"))
		Set rs = Server.CreateObject("ADODB.Recordset")
		rs.CursorLocation = 3 
		Set rs = cnn8.Execute(SQL)
		
		If rs.EOF Then
			SettingsBizIntelHasRecords = false
		Else
			SettingsBizIntelHasRecords = true	
		End If
					
		set rs = Nothing
		cnn8.close
		set cnn8 = Nothing
	
	'Just to insert NULL default values so further update query will not fail
	If NOT SettingsBizIntelHasRecords Then
	    SQL = "INSERT INTO " & MUV_Read("SQL_Owner") & ".Settings_BizIntel "
		SQL = SQL & " (CustAnalSum1EmailToUserNos) "
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
	  
	    SQL = "SELECT * FROM Settings_BizIntel"
		
		Set cnn8 = Server.CreateObject("ADODB.Connection")
		cnn8.open (Session("ClientCnnString"))
		Set rs = Server.CreateObject("ADODB.Recordset")
		rs.CursorLocation = 3 
		Set rs = cnn8.Execute(SQL)
		
		If not rs.EOF Then
	        CustAnalSum1EmailAddressesToCC_ORIG = rs("CustAnalSum1EmailAddressesToCC")
	        MCSActivitySummaryEmailAddressesToCC_ORIG = rs("MCSActivitySummaryEmailAddressesToCC")
		End If	
		
		set rs = Nothing
		cnn8.close
		set cnn8 = Nothing
	End If

	If userListName = "NotesReportAdditionalEmails" Then	
		FieldServiceNotesReportAdditionalEmails = Request.Form("txtFieldServiceNotesReportAdditionalEmails")
	
	    If FieldServiceNotesReportAdditionalEmails <> FieldServiceNotesReportAdditionalEmails_ORIG Then
			CreateAuditLogEntry "Field Service Settings Change", "Field Service Settings Change", "Major", 1, "FieldServiceNotesReportAdditionalEmails changed from " & FieldServiceNotesReportAdditionalEmails_ORIG & " to " & FieldServiceNotesReportAdditionalEmails
		End If
		
		Set cnn8 = Server.CreateObject("ADODB.Connection")
		cnn8.open (Session("ClientCnnString"))
		Set rs = Server.CreateObject("ADODB.Recordset")
		rs.CursorLocation = 3 
	
		SQL = "UPDATE  " & MUV_Read("SQL_Owner") & ".Settings_FieldService SET "
		SQL = SQL & "FieldServiceNotesReportAdditionalEmails = '" & FieldServiceNotesReportAdditionalEmails & "'"
		Set rs = cnn8.Execute(SQL)
	 	set rs = Nothing
		cnn8.close
		set cnn8 = Nothing
	     Response.Redirect("field-service.asp")
	End If

	If userListName = "CarryoverReportAdditionalEmails" Then
		ServiceTicketCarryoverReportAdditionalEmails = Request.Form("txtServiceTicketCarryoverReportAdditionalEmails")
	
	    If ServiceTicketCarryoverReportAdditionalEmails <> ServiceTicketCarryoverReportAdditionalEmails_ORIG Then
			CreateAuditLogEntry "Field Service Settings Change", "Field Service Settings Change", "Major", 1, "ServiceTicketCarryoverReportAdditionalEmails changed from " & ServiceTicketCarryoverReportAdditionalEmails_ORIG & " to " & ServiceTicketCarryoverReportAdditionalEmails
		End If
		
		Set cnn8 = Server.CreateObject("ADODB.Connection")
		cnn8.open (Session("ClientCnnString"))
		Set rs = Server.CreateObject("ADODB.Recordset")
		rs.CursorLocation = 3 
	
		SQL = "UPDATE  " & MUV_Read("SQL_Owner") & ".Settings_FieldService SET "
		SQL = SQL & "ServiceTicketCarryoverReportAdditionalEmails = '" & ServiceTicketCarryoverReportAdditionalEmails & "'"
		Set rs = cnn8.Execute(SQL)
	 	set rs = Nothing
		cnn8.close
		set cnn8 = Nothing
	     Response.Redirect("field-service.asp")
	End If
	
	If userListName = "ThresholdReportAdditionalEmails" Then
	    ServiceTicketthresholdReportAdditionalEmails = Request.Form("txtServiceTicketthresholdReportAdditionalEmails")
	
	    If ServiceTicketthresholdReportAdditionalEmails <> ServiceTicketthresholdReportAdditionalEmails_ORIG Then
			CreateAuditLogEntry "Field Service Settings Change", "Field Service Settings Change", "Major", 1, "ServiceTicketthresholdReportAdditionalEmails changed from " & ServiceTicketthresholdReportAdditionalEmails_ORIG & " to " & ServiceTicketthresholdReportAdditionalEmails
		End If
		
		Set cnn8 = Server.CreateObject("ADODB.Connection")
		cnn8.open (Session("ClientCnnString"))
		Set rs = Server.CreateObject("ADODB.Recordset")
		rs.CursorLocation = 3 
	
		SQL = "UPDATE  " & MUV_Read("SQL_Owner") & ".Settings_FieldService SET "
		SQL = SQL & "ServiceTicketthresholdReportAdditionalEmails = '" & ServiceTicketthresholdReportAdditionalEmails & "'"
		Set rs = cnn8.Execute(SQL)
	 	set rs = Nothing
		cnn8.close
		set cnn8 = Nothing
	     Response.Redirect("field-service.asp")
	End If

	If userListName = "ProspSnapshotAdditionalEmails" Then
	    ProspSnapshotAdditionalEmails = Request.Form("txtProspSnapshotAdditionalEmails")
	
	    If ProspSnapshotAdditionalEmails <> ProspSnapshotAdditionalEmails_ORIG Then
			CreateAuditLogEntry "Global Settings Change", "Global Settings Change", "Major", 1, "ProspSnapshotAdditionalEmails changed from " & ProspSnapshotAdditionalEmails_ORIG & " to " & ProspSnapshotAdditionalEmails
		End If
		
		Set cnn8 = Server.CreateObject("ADODB.Connection")
		cnn8.open (Session("ClientCnnString"))
		Set rs = Server.CreateObject("ADODB.Recordset")
		rs.CursorLocation = 3 
	
		SQL = "UPDATE  " & MUV_Read("SQL_Owner") & ".Settings_Global SET "
		SQL = SQL & "ProspSnapshotAdditionalEmails = '" & ProspSnapshotAdditionalEmails & "'"
		Set rs = cnn8.Execute(SQL)
	 	set rs = Nothing
		cnn8.close
		set cnn8 = Nothing
	     Response.Redirect("prospecting-settings.asp")
	End If

	If userListName = "ProspectingWeeklyAgendaReportAdditionalEmails" Then
	
	    ProspectingWeeklyAgendaReportAdditionalEmails = Request.Form("txtProspectingWeeklyAgendaReportAdditionalEmails")
	
	    If ProspectingWeeklyAgendaReportAdditionalEmails <> ProspectingWeeklyAgendaReportAdditionalEmails_ORIG Then
			CreateAuditLogEntry "Global Settings Change", "Global Settings Change", "Major", 1, "ProspectingWeeklyAgendaReportAdditionalEmails changed from " & ProspectingWeeklyAgendaReportAdditionalEmails_ORIG & " to " & ProspectingWeeklyAgendaReportAdditionalEmails
		End If
		
		Set cnn8 = Server.CreateObject("ADODB.Connection")
		cnn8.open (Session("ClientCnnString"))
		Set rs = Server.CreateObject("ADODB.Recordset")
		rs.CursorLocation = 3 
	
		SQL = "UPDATE  " & MUV_Read("SQL_Owner") & ".Settings_Prospecting SET "
		SQL = SQL & "ProspectingWeeklyAgendaReportAdditionalEmails = '" & ProspectingWeeklyAgendaReportAdditionalEmails & "'"
		Set rs = cnn8.Execute(SQL)
	 	set rs = Nothing
		cnn8.close
		set cnn8 = Nothing
	     Response.Redirect("prospecting-settings.asp")
	End If
	
	If userListName = "InventoryAPIDailyActivityReportAdditionalEmails" Then
	    InventoryAPIDailyActivityReportAdditionalEmails = Request.Form("txtInventoryAPIDailyActivityReportAdditionalEmails")
	
		Set cnn8 = Server.CreateObject("ADODB.Connection")
		cnn8.open (Session("ClientCnnString"))
		Set rs = Server.CreateObject("ADODB.Recordset")
		rs.CursorLocation = 3 
	
		SQL = "UPDATE  " & MUV_Read("SQL_Owner") & ".Settings_Global SET "
		SQL = SQL & "InventoryAPIDailyActivityReportAdditionalEmails = '" & InventoryAPIDailyActivityReportAdditionalEmails & "'"
		Set rs = cnn8.Execute(SQL)
	 	set rs = Nothing
		cnn8.close
		set cnn8 = Nothing
	
	    If InventoryAPIDailyActivityReportAdditionalEmails <> InventoryAPIDailyActivityReportAdditionalEmails_ORIG Then
			CreateAuditLogEntry "Global Settings Change", "Global Settings Change", "Major", 1, "Inventory API Daily Activity Report Additional Emails changed from " & InventoryAPIDailyActivityReportAdditionalEmails_ORIG & " to " & InventoryAPIDailyActivityReportAdditionalEmails
		End If
	
	    Response.Redirect("inventory.asp")
	End If
		
	If userListName = "InventoryProductChangesReportAdditionalEmails" Then
	    InventoryProductChangesReportAdditionalEmails = Request.Form("txtInventoryProductChangesReportAdditionalEmails")
	
		Set cnn8 = Server.CreateObject("ADODB.Connection")
		cnn8.open (Session("ClientCnnString"))
		Set rs = Server.CreateObject("ADODB.Recordset")
		rs.CursorLocation = 3 
	
		SQL = "UPDATE  " & MUV_Read("SQL_Owner") & ".Settings_InventoryControl SET "
		SQL = SQL & "InventoryProductChangesReportAdditionalEmails = '" & InventoryProductChangesReportAdditionalEmails & "'"
		Set rs = cnn8.Execute(SQL)
	 	set rs = Nothing
		cnn8.close
		set cnn8 = Nothing
	
	    If InventoryProductChangesReportAdditionalEmails <> InventoryProductChangesReportAdditionalEmails_ORIG Then
			CreateAuditLogEntry "Global Settings Change", "Global Settings Change", "Major", 1, "Inventory Product Changes Report Additional Emails changed from " & InventoryProductChangesReportAdditionalEmails_ORIG & " to " & InventoryProductChangesReportAdditionalEmails
		End If
	
	    Response.Redirect("inventory.asp")
	End If
	
	If userListName = "CustAnalSum1EmailAddressesToCC" Then
	    CustAnalSum1EmailAddressesToCC = Request.Form("txtCustAnalSum1EmailAddressesToCC")
	
		Set cnn8 = Server.CreateObject("ADODB.Connection")
		cnn8.open (Session("ClientCnnString"))
		Set rs = Server.CreateObject("ADODB.Recordset")
		rs.CursorLocation = 3 
	
		SQL = "UPDATE  " & MUV_Read("SQL_Owner") & ".Settings_BizIntel SET "
		SQL = SQL & "CustAnalSum1EmailAddressesToCC = '" & CustAnalSum1EmailAddressesToCC & "'"
		Set rs = cnn8.Execute(SQL)
	 	set rs = Nothing
		cnn8.close
		set cnn8 = Nothing
	
	  	IF CustAnalSum1EmailAddressesToCC <> CustAnalSum1EmailAddressesToCC_ORIG Then
			CreateAuditLogEntry "Global Settings Change", "Global Settings Change", "Major", 1, "Biz Intel Automatic Customer Analysis Summary 1 user numbers to Cc changed from " & CustAnalSum1EmailAddressesToCC_ORIG & " to " & CustAnalSum1EmailAddressesToCC 
		End If
	
	    Response.Redirect("bizintel.asp")
	End If
	

	If userListName = "MCSActivitySummaryEmailAddressesToCC" Then
	    MCSActivitySummaryEmailAddressesToCC = Request.Form("txtMCSActivitySummaryEmailAddressesToCC")
	
		Set cnn8 = Server.CreateObject("ADODB.Connection")
		cnn8.open (Session("ClientCnnString"))
		Set rs = Server.CreateObject("ADODB.Recordset")
		rs.CursorLocation = 3 
	
		SQL = "UPDATE  " & MUV_Read("SQL_Owner") & ".Settings_BizIntel SET "
		SQL = SQL & "MCSActivitySummaryEmailAddressesToCC = '" & MCSActivitySummaryEmailAddressesToCC & "'"
		Set rs = cnn8.Execute(SQL)
	 	set rs = Nothing
		cnn8.close
		set cnn8 = Nothing
	
	  	IF MCSActivitySummaryEmailAddressesToCC <> MCSActivitySummaryEmailAddressesToCC_ORIG Then
			CreateAuditLogEntry "Global Settings Change", "Global Settings Change", "Major", 1, "Biz Intel Automatic MCS Activity report user numbers to Cc changed from " & MCSActivitySummaryEmailAddressesToCC_ORIG & " to " & MCSActivitySummaryEmailAddressesToCC
		End If
	
	    Response.Redirect("bizintel.asp")
	End If


%><!--#include file="../../../inc/footer-main.asp"-->