<!--#include file="../../../../inc/header.asp"-->

<%

	Dim userListName: userListName = Request.Form("userListName")
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
	
	If NOT SettingsNeedToKnowHasRecords Then
	    SQL = "INSERT INTO " & MUV_Read("SQL_Owner") & ".Settings_NeedToKnow "
		SQL = SQL & " (N2KAREmailToUserNos, N2KARUserNosToCC, N2KAREmailAddressesToCC) "
		SQL = SQL & " VALUES "
		SQL = SQL & " ('','','') "	
		
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
	
	SQL = "SELECT * FROM Settings_NeedToKnow"
	
	Set cnn8 = Server.CreateObject("ADODB.Connection")
	cnn8.open (Session("ClientCnnString"))
	Set rs = Server.CreateObject("ADODB.Recordset")
	rs.CursorLocation = 3 
	Set rs = cnn8.Execute(SQL)
	
	If not rs.EOF Then
		N2KAPIEmailAddressesToCC_ORIG = rs("N2KAPIEmailAddressesToCC")	
		N2KAccountsReceivableEmailAddressesToCC_ORIG = rs("N2KAREmailAddressesToCC")
		N2KEquipmentEmailAddressesToCC_ORIG = rs("N2KEquipmentEmailAddressesToCC")
		N2KGlobalSettingsEmailAddressesToCC_ORIG = rs("N2KGlobalSettingsEmailAddressesToCC")
		N2KInventoryEmailAddressesToCC_ORIG = rs("N2KInventoryEmailAddressesToCC")
	End If	
	
	set rs = Nothing
	cnn8.close
	set cnn8 = Nothing
	
	
	If userListName = "N2KAPIEmailAddressesToCC" Then
		N2KAPIEmailAddressesToCC = Request.Form("txtN2KAPIEmailAddressesToCC")
	
		Set cnn8 = Server.CreateObject("ADODB.Connection")
		cnn8.open (Session("ClientCnnString"))
		Set rs = Server.CreateObject("ADODB.Recordset")
		rs.CursorLocation = 3 
	
		SQL = "UPDATE  " & MUV_Read("SQL_Owner") & ".Settings_NeedToKnow SET "
		SQL = SQL & "N2KAPIEmailAddressesToCC = '" & N2KAPIEmailAddressesToCC & "'"
		Set rs = cnn8.Execute(SQL)
	 	set rs = Nothing
		cnn8.close
		set cnn8 = Nothing
	
	    IF N2KAPIEmailAddressesToCC <> N2KAPIEmailAddressesToCC_ORIG Then
			CreateAuditLogEntry "Global Settings Change", "Global Settings Change", "Major", 1, "API Need To Know Reports user numbers to Cc changed from " & N2KAPIEmailAddressesToCC_ORIG & " to " & N2KAPIEmailAddressesToCC 
		End If
	    Response.Redirect("order-api.asp")
	End If
	

	If userListName = "N2KAccountsReceivableEmailAddressesToCC" Then
		N2KAccountsReceivableEmailAddressesToCC = Request.Form("txtN2KAccountsReceivableEmailAddressesToCC")
	
		Set cnn8 = Server.CreateObject("ADODB.Connection")
		cnn8.open (Session("ClientCnnString"))
		Set rs = Server.CreateObject("ADODB.Recordset")
		rs.CursorLocation = 3 
	
		SQL = "UPDATE  " & MUV_Read("SQL_Owner") & ".Settings_NeedToKnow SET "
		SQL = SQL & "N2KAREmailAddressesToCC = '" & N2KAccountsReceivableEmailAddressesToCC & "'"
		Set rs = cnn8.Execute(SQL)
	 	set rs = Nothing
		cnn8.close
		set cnn8 = Nothing
	
	    IF N2KAccountsReceivableEmailAddressesToCC <> N2KAccountsReceivableEmailAddressesToCC_ORIG Then
			CreateAuditLogEntry "Global Settings Change", "Global Settings Change", "Major", 1, GetTerm("Accounts Receivable") & " Need To Know Reports user numbers to Cc changed from " & N2KAccountsReceivableEmailAddressesToCC_ORIG & " to " & N2KAccountsReceivableEmailAddressesToCC 
		End If
	    Response.Redirect("accounts-receivable.asp")
	End If


	
	If userListName = "N2KEquipmentEmailAddressesToCC" Then
		N2KEquipmentEmailAddressesToCC = Request.Form("txtN2KEquipmentEmailAddressesToCC")
	
		Set cnn8 = Server.CreateObject("ADODB.Connection")
		cnn8.open (Session("ClientCnnString"))
		Set rs = Server.CreateObject("ADODB.Recordset")
		rs.CursorLocation = 3 
	
		SQL = "UPDATE  " & MUV_Read("SQL_Owner") & ".Settings_NeedToKnow SET "
		SQL = SQL & "N2KEquipmentEmailAddressesToCC = '" & N2KEquipmentEmailAddressesToCC & "'"
		Set rs = cnn8.Execute(SQL)
	 	set rs = Nothing
		cnn8.close
		set cnn8 = Nothing
	
	  	IF N2KEquipmentEmailAddressesToCC <> N2KEquipmentEmailAddressesToCC_ORIG Then
			CreateAuditLogEntry "Global Settings Change", "Global Settings Change", "Major", 1, "Equipment Need To Know Reports user numbers to Cc changed from " & N2KEquipmentEmailAddressesToCC_ORIG & " to " & N2KEquipmentEmailAddressesToCC 
		End If
	    Response.Redirect("equipment.asp")
	End If


	
	If userListName = "N2KGlobalSettingsEmailAddressesToCC" Then
		N2KGlobalSettingsEmailAddressesToCC = Request.Form("txtN2KGlobalSettingsEmailAddressesToCC")
	
		Set cnn8 = Server.CreateObject("ADODB.Connection")
		cnn8.open (Session("ClientCnnString"))
		Set rs = Server.CreateObject("ADODB.Recordset")
		rs.CursorLocation = 3 
	
		SQL = "UPDATE  " & MUV_Read("SQL_Owner") & ".Settings_NeedToKnow SET "
		SQL = SQL & "N2KGlobalSettingsEmailAddressesToCC = '" & N2KGlobalSettingsEmailAddressesToCC & "'"
		Set rs = cnn8.Execute(SQL)
	 	set rs = Nothing
		cnn8.close
		set cnn8 = Nothing
	
		IF N2KGlobalSettingsEmailAddressesToCC <> N2KGlobalSettingsEmailAddressesToCC_ORIG Then
			CreateAuditLogEntry "Global Settings Change", "Global Settings Change", "Major", 1, "Global Settings Need To Know Reports user numbers to Cc changed from " & N2KGlobalSettingsEmailAddressesToCC_ORIG & " to " & N2KGlobalSettingsEmailAddressesToCC 
		End If
	    Response.Redirect("global-settings.asp")
	End If

	
	If userListName = "N2KInventoryEmailAddressesToCC" Then
		N2KInventoryEmailAddressesToCC = Request.Form("txtN2KInventoryEmailAddressesToCC")
	
		Set cnn8 = Server.CreateObject("ADODB.Connection")
		cnn8.open (Session("ClientCnnString"))
		Set rs = Server.CreateObject("ADODB.Recordset")
		rs.CursorLocation = 3 
	
		SQL = "UPDATE  " & MUV_Read("SQL_Owner") & ".Settings_NeedToKnow SET "
		SQL = SQL & "N2KInventoryEmailAddressesToCC = '" & N2KInventoryEmailAddressesToCC & "'"
		Set rs = cnn8.Execute(SQL)
	 	set rs = Nothing
		cnn8.close
		set cnn8 = Nothing
				
		IF N2KInventoryEmailAddressesToCC <> N2KInventoryEmailAddressesToCC_ORIG Then
			CreateAuditLogEntry "Global Settings Change", "Global Settings Change", "Major", 1, "Inventory Need To Know Reports user numbers to Cc changed from " & N2KInventoryEmailAddressesToCC_ORIG & " to " & N2KInventoryEmailAddressesToCC 
		End If
	    Response.Redirect("inventory.asp")
	End If

%><!--#include file="../../../../inc/footer-main.asp"-->