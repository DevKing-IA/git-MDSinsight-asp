<!--#include file="../../../../inc/header.asp"-->

<%

	'***********************************************************
	'Get Values Of All Form Fields Posted
	'***********************************************************
	
	N2KEquipmentEmailToUserNos = Request.Form("lstSelectedN2KAPIEmailToUserNos")
	N2KEquipmentUserNosToCC = Request.Form("lstSelectedN2KAPIUserNosToCC")

	If Request.Form("chkBlankInsightAssetTagBrandPrefix") = "on" then BlankInsightAssetTagBrandPrefix = 1 Else BlankInsightAssetTagBrandPrefix = 0
	If Request.Form("chkBlankInsightAssetTagClassPrefix") = "on" then BlankInsightAssetTagClassPrefix = 1 Else BlankInsightAssetTagClassPrefix = 0
	If Request.Form("chkBlankInsightAssetTagManufacturerPrefix") = "on" then BlankInsightAssetTagManufacturerPrefix = 1 Else BlankInsightAssetTagManufacturerPrefix = 0
	If Request.Form("chkBlankInsightAssetTagModelPrefix") = "on" then BlankInsightAssetTagModelPrefix = 1 Else BlankInsightAssetTagModelPrefix = 0
	If Request.Form("chkUndefinedBrandExistsforEqp") = "on" then UndefinedBrandExistsforEqp = 1 Else UndefinedBrandExistsforEqp = 0
	If Request.Form("chkUndefinedClassExistsforEqp") = "on" then UndefinedClassExistsforEqp = 1 Else UndefinedClassExistsforEqp = 0
	If Request.Form("chkUndefinedConditionCodeExistsforEqp") = "on" then UndefinedConditionCodeExistsforEqp = 1 Else UndefinedConditionCodeExistsforEqp = 0
	If Request.Form("chkUndefinedGroupExistsforEqp") = "on" then UndefinedGroupExistsforEqp = 1 Else UndefinedGroupExistsforEqp = 0
	If Request.Form("chkUndefinedManufacturerExistsforEqp") = "on" then UndefinedManufacturerExistsforEqp = 1 Else UndefinedManufacturerExistsforEqp = 0
	If Request.Form("chkUndefinedModelExistsforEqp") = "on" then UndefinedModelExistsforEqp = 1 Else UndefinedModelExistsforEqp = 0
	If Request.Form("chkUndefinedStatusCodeExistsforEqp") = "on" then UndefinedStatusCodeExistsforEqp = 1 Else UndefinedStatusCodeExistsforEqp = 0
	If Request.Form("chkZeroDollarRentalsExistforEqp") = "on" then ZeroDollarRentalsExistforEqp = 1 Else ZeroDollarRentalsExistforEqp = 0
	
	If Request.Form("chkN2KEquipmentReportONOFF") = "on" then N2KEquipmentReportONOFF = 1 Else N2KEquipmentReportONOFF = 0
		
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
		
		N2KEquipmentEmailToUserNos_ORIG = rs("N2KEquipmentEmailToUserNos")
		N2KEquipmentUserNosToCC_ORIG = rs("N2KEquipmentUserNosToCC")
		N2KEquipmentReportONOFF_ORIG = rs("N2KEquipmentReportONOFF")
		N2KEqpIncludeBlankInsightAssetTagBrandPrefix_ORIG = rs("N2KEqpIncludeBlankInsightAssetTagBrandPrefix")
		N2KEqpIncludeBlankInsightAssetTagClassPrefix_ORIG = rs("N2KEqpIncludeBlankInsightAssetTagClassPrefix")
		N2KEqpIncludeBlankInsightAssetTagManufacturerPrefix_ORIG = rs("N2KEqpIncludeBlankInsightAssetTagManufacturerPrefix")
		N2KEqpIncludeBlankInsightAssetTagModelPrefix_ORIG = rs("N2KEqpIncludeBlankInsightAssetTagModelPrefix")
		N2KEqpIncludeUndefinedBrandExistsforEqp_ORIG = rs("N2KEqpIncludeUndefinedBrandExistsforEqp")
		N2KEqpIncludeUndefinedClassExistsforEqp_ORIG = rs("N2KEqpIncludeUndefinedClassExistsforEqp")
		N2KEqpIncludeUndefinedConditionCodeExistsforEqp_ORIG = rs("N2KEqpIncludeUndefinedConditionCodeExistsforEqp")
		N2KEqpIncludeUndefinedGroupExistsforEqp_ORIG = rs("N2KEqpIncludeUndefinedGroupExistsforEqp")
		N2KEqpIncludeUndefinedManufacturerExistsforEqp_ORIG = rs("N2KEqpIncludeUndefinedManufacturerExistsforEqp")
		N2KEqpIncludeUndefinedModelExistsforEqp_ORIG = rs("N2KEqpIncludeUndefinedModelExistsforEqp")
		N2KEqpIncludeUndefinedStatusCodeExistsforEqp_ORIG = rs("N2KEqpIncludeUndefinedStatusCodeExistsforEqp")
		N2KEqpIncludeZeroDollarRentalsExistforEqp_ORIG = rs("N2KEqpIncludeZeroDollarRentalsExistforEqp")

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
		SQL = SQL & "N2KEquipmentEmailToUserNos = '" & N2KEquipmentEmailToUserNos & "',"
		SQL = SQL & "N2KEquipmentUserNosToCC = '" & N2KEquipmentUserNosToCC & "',"
		SQL = SQL & "N2KEquipmentReportONOFF = " & N2KEquipmentReportONOFF & ","
		SQL = SQL & "N2KEqpIncludeBlankInsightAssetTagBrandPrefix = " & BlankInsightAssetTagBrandPrefix & ","
		SQL = SQL & "N2KEqpIncludeBlankInsightAssetTagClassPrefix = " & BlankInsightAssetTagClassPrefix & ","
		SQL = SQL & "N2KEqpIncludeBlankInsightAssetTagManufacturerPrefix = " & BlankInsightAssetTagManufacturerPrefix & ","
		SQL = SQL & "N2KEqpIncludeBlankInsightAssetTagModelPrefix = " & BlankInsightAssetTagModelPrefix & ","
		SQL = SQL & "N2KEqpIncludeUndefinedBrandExistsforEqp = " & UndefinedBrandExistsforEqp & ","
		SQL = SQL & "N2KEqpIncludeUndefinedClassExistsforEqp = " & UndefinedClassExistsforEqp & ","
		SQL = SQL & "N2KEqpIncludeUndefinedConditionCodeExistsforEqp = " & UndefinedConditionCodeExistsforEqp & ","
		SQL = SQL & "N2KEqpIncludeUndefinedGroupExistsforEqp = " & UndefinedGroupExistsforEqp & ","
		SQL = SQL & "N2KEqpIncludeUndefinedManufacturerExistsforEqp = " & UndefinedManufacturerExistsforEqp & ","
		SQL = SQL & "N2KEqpIncludeUndefinedModelExistsforEqp = " & UndefinedModelExistsforEqp & ","
		SQL = SQL & "N2KEqpIncludeUndefinedStatusCodeExistsforEqp = " & UndefinedStatusCodeExistsforEqp & ","
		SQL = SQL & "N2KEqpIncludeZeroDollarRentalsExistforEqp = " & ZeroDollarRentalsExistforEqp

	Else
	
		SQL = "INSERT INTO " & MUV_Read("SQL_Owner") & ".Settings_NeedToKnow "
		SQL = SQL & " (N2KEquipmentEmailToUserNos, N2KEquipmentUserNosToCC,N2KEquipmentReportONOFF, N2KEqpIncludeBlankInsightAssetTagBrandPrefix, N2KEqpIncludeBlankInsightAssetTagClassPrefix, N2KEqpIncludeBlankInsightAssetTagManufacturerPrefix, N2KEqpIncludeBlankInsightAssetTagModelPrefix, N2KEqpIncludeUndefinedBrandExistsforEqp, N2KEqpIncludeUndefinedClassExistsforEqp, N2KEqpIncludeUndefinedConditionCodeExistsforEqp, N2KEqpIncludeUndefinedGroupExistsforEqp, N2KEqpIncludeUndefinedManufacturerExistsforEqp, N2KEqpIncludeUndefinedModelExistsforEqp, N2KEqpIncludeUndefinedStatusCodeExistsforEqp, N2KEqpIncludeZeroDollarRentalsExistforEqp) "
		SQL = SQL & " VALUES "
		SQL = SQL & " ('" & N2KEquipmentEmailToUserNos & "', '" & N2KEquipmentUserNosToCC & "'," & N2KEquipmentReportONOFF & "," & BlankInsightAssetTagBrandPrefix & "," & BlankInsightAssetTagClassPrefix & "," & BlankInsightAssetTagManufacturerPrefix & "," & BlankInsightAssetTagModelPrefix & "," & UndefinedBrandExistsforEqp & "," & UndefinedClassExistsforEqp & "," & UndefinedConditionCodeExistsforEqp & "," & UndefinedGroupExistsforEqp & "," & UndefinedManufacturerExistsforEqp & "," & UndefinedModelExistsforEqp & "," & UndefinedStatusCodeExistsforEqp & "," & ZeroDollarRentalsExistforEqp & ") "
	
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

	If N2KEquipmentEmailToUserNos <> N2KEquipmentEmailToUserNos_ORIG Then

		userNames = ""
		userNamesOrig = ""
		
	
		IndividualN2KEquipmentEmailToUserNos = Split(N2KEquipmentEmailToUserNos,",")
		
		for i=0 to Ubound(IndividualN2KEquipmentEmailToUserNos)
		     userNames = userNames & " " & GetUserFirstAndLastNameByUserNo(IndividualN2KEquipmentEmailToUserNos (i))
		next
		
		
		IndividualN2KEquipmentEmailToUserNos_ORIG  = Split(N2KEquipmentEmailToUserNos_ORIG,",")
		
		for i=0 to Ubound(IndividualN2KEquipmentEmailToUserNos_ORIG)
		     userNamesOrig = userNamesOrig & " " & GetUserFirstAndLastNameByUserNo(IndividualN2KEquipmentEmailToUserNos_ORIG(i))
		next
		
		
		CreateAuditLogEntry "Global Settings Change", "Global Settings Change", "Major", 1, "Equipment Need To Know Reports send email to changed from " & userNamesOrig & " to " & userNames
	End If


	If N2KEquipmentUserNosToCC <> N2KEquipmentUserNosToCC_ORIG Then

		userNames = ""
		userNamesOrig = ""
		
	
		IndividualN2KEquipmentUserNosToCC  = Split(N2KEquipmentUserNosToCC,",")
		
		for i=0 to Ubound(IndividualN2KEquipmentUserNosToCC)
		     userNames = userNames & " " & GetUserFirstAndLastNameByUserNo(IndividualN2KEquipmentUserNosToCC (i))
		next
		
		
		IndividualN2KEquipmentUserNosToCC_ORIG  = Split(IndividualN2KEquipmentUserNosToCC_ORIG,",")
		
		for i=0 to Ubound(IndividualN2KEquipmentUserNosToCC_ORIG)
		     userNamesOrig = userNamesOrig & " " & GetUserFirstAndLastNameByUserNo(IndividualN2KEquipmentUserNosToCC_ORIG(i))
		next
		
		
		CreateAuditLogEntry "Global Settings Change", "Global Settings Change", "Major", 1, "Equipment Need To Know Reports users to Cc changed from " & userNamesOrig & " to " & userNames
	End If


	If Request.Form("chkN2KEquipmentReportONOFF")="on" then N2KEquipmentReportONOFFMsg = "On" Else N2KEquipmentReportONOFFMsg = "Off"
	If N2KEquipmentReportONOFF_ORIG = 1 then N2KEquipmentReportONOFFMsg_ORIGFMsg = "On" Else N2KEquipmentReportONOFFMsg_ORIGFMsg = "Off"
	
	IF N2KEquipmentReportONOFF <> N2KEquipmentReportONOFF_ORIG Then
		CreateAuditLogEntry "Global Settings Change", "Global Settings Change", "Major", 1, "Equipment need to know report changed from " & N2KEquipmentReportONOFFMsg_ORIGFMsg & " to " & N2KEquipmentReportONOFFMsg 
	End If


	If Request.Form("chkBlankInsightAssetTagBrandPrefix") = "on" then BlankInsightAssetTagBrandPrefixMsg = "On" Else BlankInsightAssetTagBrandPrefixMsg = "Off"
	If N2KEqpIncludeBlankInsightAssetTagBrandPrefix_ORIG = 1 then N2KEqpIncludeBlankInsightAssetTagBrandPrefix_ORIGMsg = "On" Else N2KEqpIncludeBlankInsightAssetTagBrandPrefix_ORIGMsg = "Off"

	IF BlankInsightAssetTagBrandPrefix <> N2KEqpIncludeBlankInsightAssetTagBrandPrefix_ORIG Then
		CreateAuditLogEntry "Global Settings Change", "Global Settings Change", "Major", 1, "Equipment need to know report Blank Insight Asset Tag Brand Prefix changed from " & N2KEqpIncludeBlankInsightAssetTagBrandPrefix_ORIGMsg & " to " & BlankInsightAssetTagBrandPrefixMsg 
	End If

	
	If Request.Form("chkBlankInsightAssetTagClassPrefix") = "on" then BlankInsightAssetTagClassPrefixMsg = "On" Else BlankInsightAssetTagClassPrefixMsg = "Off"
	If N2KEqpIncludeBlankInsightAssetTagClassPrefix_ORIG = 1 then N2KEqpIncludeBlankInsightAssetTagClassPrefix_ORIGMsg = "On" Else N2KEqpIncludeBlankInsightAssetTagClassPrefix_ORIGMsg = "Off"

	IF BlankInsightAssetTagClassPrefix <> N2KEqpIncludeBlankInsightAssetTagClassPrefix_ORIG Then
		CreateAuditLogEntry "Global Settings Change", "Global Settings Change", "Major", 1, "Equipment need to know report Blank Insight Asset Tag Class Prefix changed from " & N2KEqpIncludeBlankInsightAssetTagClassPrefix_ORIGMsg & " to " & BlankInsightAssetTagClassPrefixMsg 
	End If

	
	If Request.Form("chkBlankInsightAssetTagManufacturerPrefix") = "on" then BlankInsightAssetTagManufacturerPrefixMsg = "On" Else BlankInsightAssetTagManufacturerPrefixMsg = "Off"
	If N2KEqpIncludeBlankInsightAssetTagManufacturerPrefix_ORIG = 1 then N2KEqpIncludeBlankInsightAssetTagManufacturerPrefix_ORIGMsg = "On" Else N2KEqpIncludeBlankInsightAssetTagManufacturerPrefix_ORIGMsg = "Off"

	IF BlankInsightAssetTagManufacturerPrefix <> N2KEqpIncludeBlankInsightAssetTagManufacturerPrefix_ORIG Then
		CreateAuditLogEntry "Global Settings Change", "Global Settings Change", "Major", 1, "Equipment need to know report Blank Insight Asset Tag Manufacturer Prefix changed from " & N2KEqpIncludeBlankInsightAssetTagManufacturerPrefix_ORIGMsg & " to " & BlankInsightAssetTagManufacturerPrefixMsg 
	End If

	
	If Request.Form("chkBlankInsightAssetTagModelPrefix") = "on" then BlankInsightAssetTagModelPrefixMsg = "On" Else BlankInsightAssetTagModelPrefixMsg = "Off"
	If N2KEqpIncludeBlankInsightAssetTagModelPrefix_ORIG = 1 then N2KEqpIncludeBlankInsightAssetTagModelPrefix_ORIGMsg = "On" Else N2KEqpIncludeBlankInsightAssetTagModelPrefix_ORIGMsg = "Off"

	IF BlankInsightAssetTagModelPrefix <> N2KEqpIncludeBlankInsightAssetTagModelPrefix_ORIG Then
		CreateAuditLogEntry "Global Settings Change", "Global Settings Change", "Major", 1, "Equipment need to know report Blank Insight Asset Tag Model Prefix changed from " & N2KEqpIncludeBlankInsightAssetTagModelPrefix_ORIGMsg & " to " & BlankInsightAssetTagModelPrefixMsg 
	End If

	
	If Request.Form("chkUndefinedBrandExistsforEqp") = "on" then UndefinedBrandExistsforEqpMsg = "On" Else UndefinedBrandExistsforEqpMsg = "Off"
	If N2KEqpIncludeUndefinedBrandExistsforEqp_ORIG = 1 then N2KEqpIncludeUndefinedBrandExistsforEqp_ORIGMsg = "On" Else N2KEqpIncludeUndefinedBrandExistsforEqp_ORIGMsg = "Off"

	IF UndefinedBrandExistsforEqp <> N2KEqpIncludeUndefinedBrandExistsforEqp_ORIG Then
		CreateAuditLogEntry "Global Settings Change", "Global Settings Change", "Major", 1, "Equipment need to know report Undefined Brand Exists for Eqp changed from " & N2KEqpIncludeUndefinedBrandExistsforEqp_ORIGMsg & " to " & UndefinedBrandExistsforEqpMsg 
	End If

	
	If Request.Form("chkUndefinedClassExistsforEqp") = "on" then UndefinedClassExistsforEqpMsg = "On" Else UndefinedClassExistsforEqpMsg = "Off"
	If N2KEqpIncludeUndefinedClassExistsforEqp_ORIG = 1 then N2KEqpIncludeUndefinedClassExistsforEqp_ORIGMsg = "On" Else N2KEqpIncludeUndefinedClassExistsforEqp_ORIGMsg = "Off"

	IF UndefinedClassExistsforEqp <> N2KEqpIncludeUndefinedClassExistsforEqp_ORIG Then
		CreateAuditLogEntry "Global Settings Change", "Global Settings Change", "Major", 1, "Equipment need to know report Undefined Class Exists for Eqp changed from " & N2KEqpIncludeUndefinedClassExistsforEqp_ORIGMsg & " to " & UndefinedClassExistsforEqpMsg 
	End If

	
	If Request.Form("chkUndefinedConditionCodeExistsforEqp") = "on" then UndefinedConditionCodeExistsforEqpMsg = "On" Else UndefinedConditionCodeExistsforEqpMsg = "Off"
	If N2KEqpIncludeUndefinedConditionCodeExistsforEqp_ORIG = 1 then N2KEqpIncludeUndefinedConditionCodeExistsforEqp_ORIGMsg = "On" Else N2KEqpIncludeUndefinedConditionCodeExistsforEqp_ORIGMsg = "Off"

	IF UndefinedConditionCodeExistsforEqp <> N2KEqpIncludeUndefinedConditionCodeExistsforEqp_ORIG Then
		CreateAuditLogEntry "Global Settings Change", "Global Settings Change", "Major", 1, "Equipment need to know report Undefined Condition Code Exists for Eqp changed from " & N2KEqpIncludeUndefinedConditionCodeExistsforEqp_ORIGMsg & " to " & UndefinedConditionCodeExistsforEqpMsg 
	End If

	
	If Request.Form("chkUndefinedGroupExistsforEqp") = "on" then UndefinedGroupExistsforEqpMsg = "On" Else UndefinedGroupExistsforEqpMsg = "Off"
	If N2KEqpIncludeUndefinedGroupExistsforEqp_ORIG = 1 then N2KEqpIncludeUndefinedGroupExistsforEqp_ORIGMsg = "On" Else N2KEqpIncludeUndefinedGroupExistsforEqp_ORIGMsg = "Off"

	IF UndefinedGroupExistsforEqp <> N2KEqpIncludeUndefinedGroupExistsforEqp_ORIG Then
		CreateAuditLogEntry "Global Settings Change", "Global Settings Change", "Major", 1, "Equipment need to know report Undefined Group Exists for Eqp changed from " & N2KEqpIncludeUndefinedGroupExistsforEqp_ORIGMsg & " to " & UndefinedGroupExistsforEqpMsg 
	End If

	
	If Request.Form("chkUndefinedManufacturerExistsforEqp") = "on" then UndefinedManufacturerExistsforEqpMsg = "On" Else UndefinedManufacturerExistsforEqpMsg = "Off"
	If N2KEqpIncludeUndefinedManufacturerExistsforEqp_ORIG = 1 then N2KEqpIncludeUndefinedManufacturerExistsforEqp_ORIGMsg = "On" Else N2KEqpIncludeUndefinedManufacturerExistsforEqp_ORIGMsg = "Off"

	IF UndefinedManufacturerExistsforEqp <> N2KEqpIncludeUndefinedManufacturerExistsforEqp_ORIG Then
		CreateAuditLogEntry "Global Settings Change", "Global Settings Change", "Major", 1, "Equipment need to know report Undefined Manufacturer Exists for Eqp changed from " & N2KEqpIncludeUndefinedManufacturerExistsforEqp_ORIGMsg & " to " & UndefinedManufacturerExistsforEqpMsg
	End If

	
	If Request.Form("chkUndefinedModelExistsforEqp") = "on" then UndefinedModelExistsforEqpMsg = "On" Else UndefinedModelExistsforEqpMsg = "Off"
	If N2KEqpIncludeUndefinedModelExistsforEqp_ORIG = 1 then N2KEqpIncludeUndefinedModelExistsforEqp_ORIGMsg = "On" Else N2KEqpIncludeUndefinedModelExistsforEqp_ORIGMsg = "Off"

	IF UndefinedModelExistsforEqp <> N2KEqpIncludeUndefinedModelExistsforEqp_ORIG Then
		CreateAuditLogEntry "Global Settings Change", "Global Settings Change", "Major", 1, "Equipment need to know report Undefined Model Exists for Eqp changed from " & N2KEqpIncludeUndefinedModelExistsforEqp_ORIGMsg & " to " & UndefinedModelExistsforEqpMsg 
	End If

	
	If Request.Form("chkUndefinedStatusCodeExistsforEqp") = "on" then UndefinedStatusCodeExistsforEqpMsg = "On" Else UndefinedStatusCodeExistsforEqpMsg = "Off"
	If N2KEqpIncludeUndefinedStatusCodeExistsforEqp_ORIG = 1 then N2KEqpIncludeUndefinedStatusCodeExistsforEqp_ORIGMsg = "On" Else N2KEqpIncludeUndefinedStatusCodeExistsforEqp_ORIGMsg = "Off"

	IF UndefinedStatusCodeExistsforEqp <> N2KEqpIncludeUndefinedStatusCodeExistsforEqp_ORIG Then
		CreateAuditLogEntry "Global Settings Change", "Global Settings Change", "Major", 1, "Equipment need to know report Undefined Status Code Exists for Eqp changed from " & N2KEqpIncludeUndefinedStatusCodeExistsforEqp_ORIGMsg & " to " & UndefinedStatusCodeExistsforEqpMsg 
	End If

	
	If Request.Form("chkZeroDollarRentalsExistforEqp") = "on" then ZeroDollarRentalsExistforEqpMsg = "On" Else ZeroDollarRentalsExistforEqpMsg = "Off"
	If N2KEqpIncludeZeroDollarRentalsExistforEqp_ORIG = 1 then N2KEqpIncludeZeroDollarRentalsExistforEqp_ORIGMsg = "On" Else N2KEqpIncludeZeroDollarRentalsExistforEqp_ORIGMsg = "Off"

	IF ZeroDollarRentalsExistforEqp <> N2KEqpIncludeZeroDollarRentalsExistforEqp_ORIG Then
		CreateAuditLogEntry "Global Settings Change", "Global Settings Change", "Major", 1, "Equipment need to know report Zero Dollar Rentals Exist for Eqp changed from " & N2KEqpIncludeZeroDollarRentalsExistforEqp_ORIGMsg & " to " & ZeroDollarRentalsExistforEqpMsg 
	End If
	
	Response.Redirect("equipment.asp")
%><!--#include file="../../../../inc/footer-main.asp"-->