<!--#include file="../../../../inc/header.asp"-->

<%

	'***********************************************************
	'Get Values Of All Form Fields Posted
	'***********************************************************
	
	N2KInventoryEmailToUserNos = Request.Form("lstSelectedN2KAPIEmailToUserNos")
	N2KInventoryUserNosToCC = Request.Form("lstSelectedN2KAPIUserNosToCC")
	
	If Request.Form("chkBlankCaseBin") = "on" then BlankCaseBin = 1 Else BlankCaseBin = 0
	If Request.Form("chkBlankCaseUPCCode") = "on" then BlankCaseUPCCode = 1 Else BlankCaseUPCCode = 0
	If Request.Form("chkBlankUnitandCaseUPCCode") = "on" then BlankUnitandCaseUPCCode = 1 Else BlankUnitandCaseUPCCode = 0
	If Request.Form("chkBlankUnitBin") = "on" then BlankUnitBin = 1 Else BlankUnitBin = 0
	If Request.Form("chkBlankUnitUPCCode") = "on" then BlankUnitUPCCode = 1 Else BlankUnitUPCCode = 0
	If Request.Form("chkDuplicateUnitorCaseBin") = "on" then DuplicateUnitorCaseBin = 1 Else DuplicateUnitorCaseBin = 0
	If Request.Form("chkDuplicateUPCCode") = "on" then DuplicateUPCCode = 1 Else DuplicateUPCCode = 0
	
	If Request.Form("chkN2KInventoryReportONOFF") = "on" then N2KInventoryReportONOFF = 1 Else N2KInventoryReportONOFF = 0
	N2KInventoryAllowedDuplicateBins = Request.Form("txtN2KInventoryAllowedDuplicateBins")
	
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
		
		N2KInventoryEmailToUserNos_ORIG = rs("N2KInventoryEmailToUserNos")
		N2KInventoryUserNosToCC_ORIG = rs("N2KInventoryUserNosToCC")
		N2KInventoryReportONOFF_ORIG = rs("N2KInventoryReportONOFF")
		N2KInventoryAllowedDuplicateBins_ORIG = rs("N2KInventoryAllowedDuplicateBins")
		N2KInventoryIncludeBlankCaseBin_ORIG = rs("N2KInventoryIncludeBlankCaseBin")
		N2KInventoryIncludeBlankCaseUPCCode_ORIG = rs("N2KInventoryIncludeBlankCaseUPCCode")
		N2KInventoryIncludeBlankUnitandCaseUPCCode_ORIG = rs("N2KInventoryIncludeBlankUnitandCaseUPCCode")
		N2KInventoryIncludeBlankUnitBin_ORIG = rs("N2KInventoryIncludeBlankUnitBin")
		N2KInventoryIncludeBlankUnitUPCCode_ORIG = rs("N2KInventoryIncludeBlankUnitUPCCode")
		N2KInventoryIncludeDuplicateUnitorCaseBin_ORIG = rs("N2KInventoryIncludeDuplicateUnitorCaseBin")
		N2KInventoryIncludeDuplicateUPCCode_ORIG = rs("N2KInventoryIncludeDuplicateUPCCode")
		
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
		SQL = SQL & "N2KInventoryEmailToUserNos = '" & N2KInventoryEmailToUserNos & "',"
		SQL = SQL & "N2KInventoryUserNosToCC = '" & N2KInventoryUserNosToCC & "',"
		SQL = SQL & "N2KInventoryReportONOFF = " & N2KInventoryReportONOFF & ", "
		SQL = SQL & "N2KInventoryAllowedDuplicateBins = '" & N2KInventoryAllowedDuplicateBins & "',"
		SQL = SQL & "N2KInventoryIncludeBlankCaseBin = '" & BlankCaseBin & "',"
		SQL = SQL & "N2KInventoryIncludeBlankCaseUPCCode = '" & BlankCaseUPCCode & "',"
		SQL = SQL & "N2KInventoryIncludeBlankUnitandCaseUPCCode = '" & BlankUnitandCaseUPCCode & "',"
		SQL = SQL & "N2KInventoryIncludeBlankUnitBin = '" & BlankUnitBin & "',"
		SQL = SQL & "N2KInventoryIncludeBlankUnitUPCCode = '" & BlankUnitUPCCode & "',"
		SQL = SQL & "N2KInventoryIncludeDuplicateUnitorCaseBin = '" & DuplicateUnitorCaseBin & "',"
		SQL = SQL & "N2KInventoryIncludeDuplicateUPCCode = '" & DuplicateUPCCode & "'"

	Else
	
		SQL = "INSERT INTO " & MUV_Read("SQL_Owner") & ".Settings_NeedToKnow "
		SQL = SQL & " (N2KInventoryEmailToUserNos, N2KInventoryUserNosToCC,N2KInventoryReportONOFF, N2KInventoryAllowedDuplicateBins, N2KInventoryIncludeBlankCaseBin, N2KInventoryIncludeBlankCaseUPCCode, N2KInventoryIncludeBlankUnitandCaseUPCCode, N2KInventoryIncludeBlankUnitBin, N2KInventoryIncludeBlankUnitUPCCode, N2KInventoryIncludeDuplicateUnitorCaseBin, N2KInventoryIncludeDuplicateUPCCode) "
		SQL = SQL & " VALUES "
		SQL = SQL & " ('" & N2KInventoryEmailToUserNos & "','" & N2KInventoryUserNosToCC & "'," & N2KInventoryReportONOFF & ",'" & N2KInventoryAllowedDuplicateBins & "',"	& BlankCaseBin & "," & BlankCaseUPCCode & "," & BlankUnitandCaseUPCCode & "," & BlankUnitBin & "," & BlankUnitUPCCode & "," & DuplicateUnitorCaseBin & "," & DuplicateUPCCode & " ) "
	
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

	If N2KInventoryEmailToUserNos <> N2KInventoryEmailToUserNos_ORIG Then

		userNames = ""
		userNamesOrig = ""
		
	
		IndividualN2KInventoryEmailToUserNos = Split(N2KInventoryEmailToUserNos,",")
		
		for i=0 to Ubound(IndividualN2KInventoryEmailToUserNos)
		     userNames = userNames & " " & GetUserFirstAndLastNameByUserNo(IndividualN2KInventoryEmailToUserNos (i))
		next
		
		
		IndividualN2KInventoryEmailToUserNos_ORIG  = Split(N2KInventoryEmailToUserNos_ORIG,",")
		
		for i=0 to Ubound(IndividualN2KInventoryEmailToUserNos_ORIG)
		     userNamesOrig = userNamesOrig & " " & GetUserFirstAndLastNameByUserNo(IndividualN2KInventoryEmailToUserNos_ORIG(i))
		next
		
		
		CreateAuditLogEntry "Global Settings Change", "Global Settings Change", "Major", 1, "Inventory Need To Know Reports send email to changed from " & userNamesOrig & " to " & userNames
	End If


	If N2KInventoryUserNosToCC <> N2KInventoryUserNosToCC_ORIG Then

		userNames = ""
		userNamesOrig = ""
		
	
		IndividualN2KInventoryUserNosToCC  = Split(N2KInventoryUserNosToCC,",")
		
		for i=0 to Ubound(IndividualN2KInventoryUserNosToCC)
		     userNames = userNames & " " & GetUserFirstAndLastNameByUserNo(IndividualN2KInventoryUserNosToCC (i))
		next
		
		
		IndividualN2KInventoryUserNosToCC_ORIG  = Split(IndividualN2KInventoryUserNosToCC_ORIG,",")
		
		for i=0 to Ubound(IndividualN2KInventoryUserNosToCC_ORIG)
		     userNamesOrig = userNamesOrig & " " & GetUserFirstAndLastNameByUserNo(IndividualN2KInventoryUserNosToCC_ORIG(i))
		next
		
		
		CreateAuditLogEntry "Global Settings Change", "Global Settings Change", "Major", 1, "Inventory Need To Know Reports users to Cc changed from " & userNamesOrig & " to " & userNames
	End If

		
	If Request.Form("chkN2KInventoryReportONOFF")="on" then N2KInventoryReportONOFFMsg = "On" Else N2KInventoryReportONOFFMsg = "Off"
	If N2KInventoryReportONOFFMsg_ORIG = 1 then N2KInventoryReportONOFFMsg_ORIGFMsg = "On" Else N2KInventoryReportONOFFMsg_ORIGFMsg = "Off"
	
	IF N2KInventoryReportONOFF <> N2KInventoryReportONOFFMsg_ORIG Then
		CreateAuditLogEntry "Global Settings Change", "Global Settings Change", "Major", 1, "Inventory need to know report changed from " & N2KInventoryReportONOFFMsg_ORIGFMsg & " to " & N2KInventoryReportONOFFMsg 
	End If

	IF N2KInventoryAllowedDuplicateBins <> N2KInventoryAllowedDuplicateBins_ORIG Then
		CreateAuditLogEntry "Global Settings Change", "Global Settings Change", "Major", 1, "Inventory Need To Know Reports duplicate bin locations changed from " & N2KInventoryAllowedDuplicateBins_ORIG & " to " & N2KInventoryAllowedDuplicateBins
	End If


	If Request.Form("chkBlankCaseBin") = "on" then BlankCaseBinMsg = "On" Else BlankCaseBinMsg = "Off"
	If N2KInventoryIncludeBlankCaseBin_ORIG = 1 then N2KInventoryIncludeBlankCaseBin_ORIGMsg = "On" Else N2KInventoryIncludeBlankCaseBin_ORIGMsg = "Off"

	IF BlankCaseBin <> N2KInventoryIncludeBlankCaseBin_ORIG Then
		CreateAuditLogEntry "Global Settings Change", "Global Settings Change", "Major", 1, "Inventory need to know report Blank Case Bin changed from " & N2KInventoryIncludeBlankCaseBin_ORIGMsg & " to " & BlankCaseBinMsg 
	End If
	

	If Request.Form("chkBlankCaseUPCCode") = "on" then BlankCaseUPCCodeMsg = "On" Else BlankCaseUPCCodeMsg = "Off"
	If N2KInventoryIncludeBlankCaseUPCCode_ORIG = 1 then N2KInventoryIncludeBlankCaseUPCCode_ORIGMsg = "On" Else N2KInventoryIncludeBlankCaseUPCCode_ORIGMsg = "Off"

	IF BlankCaseUPCCode <> N2KInventoryIncludeBlankCaseUPCCode_ORIG Then
		CreateAuditLogEntry "Global Settings Change", "Global Settings Change", "Major", 1, "Inventory need to know report Blank Case UPC Code changed from " & N2KInventoryIncludeBlankCaseUPCCode_ORIGMsg & " to " & BlankCaseUPCCodeMsg 
	End If

	
	If Request.Form("chkBlankUnitandCaseUPCCode") = "on" then BlankUnitandCaseUPCCodeMsg = "On" Else BlankUnitandCaseUPCCodeMsg = "Off"
	If N2KInventoryIncludeBlankUnitandCaseUPCCode_ORIG = 1 then N2KInventoryIncludeBlankUnitandCaseUPCCode_ORIGMsg = "On" Else N2KInventoryIncludeBlankUnitandCaseUPCCode_ORIGMsg = "Off"

	IF BlankUnitandCaseUPCCode <> N2KInventoryIncludeBlankUnitandCaseUPCCode_ORIG Then
		CreateAuditLogEntry "Global Settings Change", "Global Settings Change", "Major", 1, "Inventory need to know report Blank Unit and Case UPC Code changed from " & N2KInventoryIncludeBlankUnitandCaseUPCCode_ORIGMsg & " to " & BlankUnitandCaseUPCCodeMsg 
	End If

	
	If Request.Form("chkBlankUnitBin") = "on" then BlankUnitBinMsg = "On" Else BlankUnitBinMsg = "Off"
	If N2KInventoryIncludeBlankUnitBin_ORIG = 1 then N2KInventoryIncludeBlankUnitBin_ORIGMsg = "On" Else N2KInventoryIncludeBlankUnitBin_ORIGMsg = "Off"

	IF BlankUnitBin <> N2KInventoryIncludeBlankUnitBin_ORIG Then
		CreateAuditLogEntry "Global Settings Change", "Global Settings Change", "Major", 1, "Inventory need to know report Blank Unit Bin changed from " & N2KInventoryIncludeBlankUnitBin_ORIGMsg & " to " & BlankUnitBinMsg 
	End If

	
	If Request.Form("chkBlankUnitUPCCode") = "on" then BlankUnitUPCCodeMsg = "On" Else BlankUnitUPCCodeMsg = "Off"
	If N2KInventoryIncludeBlankUnitUPCCode_ORIG = 1 then N2KInventoryIncludeBlankUnitUPCCode_ORIGMsg = "On" Else N2KInventoryIncludeBlankUnitUPCCode_ORIGMsg = "Off"

	IF BlankUnitUPCCode <> N2KInventoryIncludeBlankUnitUPCCode_ORIG Then
		CreateAuditLogEntry "Global Settings Change", "Global Settings Change", "Major", 1, "Inventory need to know report Blank Unit UPC Code changed from " & N2KInventoryIncludeBlankUnitUPCCode_ORIGMsg & " to " & BlankUnitUPCCodeMsg 
	End If

	
	If Request.Form("chkDuplicateUnitorCaseBin") = "on" then DuplicateUnitorCaseBinMsg = "On" Else DuplicateUnitorCaseBinMsg = "Off"
	If N2KInventoryIncludeDuplicateUnitorCaseBin_ORIG = 1 then N2KInventoryIncludeDuplicateUnitorCaseBin_ORIGMsg = "On" Else N2KInventoryIncludeDuplicateUnitorCaseBin_ORIGMsg = "Off"

	IF DuplicateUnitorCaseBin <> N2KInventoryIncludeDuplicateUnitorCaseBin_ORIG Then
		CreateAuditLogEntry "Global Settings Change", "Global Settings Change", "Major", 1, "Inventory need to know report Duplicate Unit or Case Bin changed from " & N2KInventoryIncludeDuplicateUnitorCaseBin_ORIGMsg & " to " & DuplicateUnitorCaseBinMsg 
	End If

	
	If Request.Form("chkDuplicateUPCCode") = "on" then DuplicateUPCCodeMsg = "On" Else DuplicateUPCCodeMsg = "Off"
	If N2KInventoryIncludeDuplicateUPCCode_ORIG = 1 then N2KInventoryIncludeDuplicateUPCCode_ORIGMsg = "On" Else N2KInventoryIncludeDuplicateUPCCode_ORIGMsg = "Off"

	IF DuplicateUPCCode <> N2KInventoryIncludeDuplicateUPCCode_ORIG Then
		CreateAuditLogEntry "Global Settings Change", "Global Settings Change", "Major", 1, "Inventory need to know report Duplicate UPC Code changed from " & N2KInventoryIncludeDuplicateUPCCode_ORIGMsg & " to " & DuplicateUPCCodeMsg 
	End If
	
	Response.Redirect("inventory.asp")
%><!--#include file="../../../../inc/footer-main.asp"-->