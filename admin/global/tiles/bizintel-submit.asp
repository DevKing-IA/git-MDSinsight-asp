<!--#include file="../../../inc/header.asp"-->

<%
	'***********************************************************
	'Get Values Of All Form Fields Posted
	'***********************************************************
	
	If Request.Form("chkCustAnalSum1OnOff") = "on" then CustAnalSum1OnOff = 1 Else CustAnalSum1OnOff = 0
	CustAnalSum1EmailToUserNos = Request.Form("lstSelectedCustAnalSum1EmailToUserIDs")
	CustAnalSum1UserNosToCC = Request.Form("lstSelectedCustAnalSum1EmailToUserIDsCC")
	MCSUserNosToCC = Request.Form("lstSelectedMCSAnalysisCCUserIDs")

	If Request.Form("chkMCSActivitySummaryOnOff") = "on" then MCSActivitySummaryOnOff = 1 Else MCSActivitySummaryOnOff = 0
	MCSActivitySummaryEmailToUserNos = Request.Form("lstSelectedMCSActivitySummaryEmailToUserIDs")
	MCSActivitySummaryUserNosToCC = Request.Form("lstSelectedMCSActivitySummaryCCToUserIDs")

	If Request.Form("chkMCSUseAlternateHeader") = "on" then MCSUseAlternateHeader = 1 Else MCSUseAlternateHeader = 0
		
	'***********************************************************
	'Get Original Values For Audit Trail Entries
	'***********************************************************
	
	SQL = "SELECT * FROM Settings_BizIntel"
	
	Set cnn8 = Server.CreateObject("ADODB.Connection")
	cnn8.open (Session("ClientCnnString"))
	Set rs = Server.CreateObject("ADODB.Recordset")
	rs.CursorLocation = 3 
	Set rs = cnn8.Execute(SQL)
	
	If not rs.EOF Then
		CustAnalSum1OnOff_ORIG = rs("CustAnalSum1OnOff")
		CustAnalSum1EmailToUserNos_ORIG = rs("CustAnalSum1EmailToUserNos")
		CustAnalSum1UserNosToCC_ORIG = rs("CustAnalSum1UserNosToCC")
		MCSUserNosToCC_ORIG = rs("MCSUserNosToCC")
		MCSActivitySummaryOnOff_ORIG = rs("MCSActivitySummaryOnOff")
		MCSActivitySummaryEmailToUserNos_ORIG = rs("MCSActivitySummaryEmailToUserNos")
		MCSActivitySummaryUserNosToCC_ORIG = rs("MCSActivitySummaryUserNosToCC")
		MCSUseAlternateHeader_ORIG = rs("MCSUseAlternateHeader")
	End If
				
	set rs = Nothing
	cnn8.close
	set cnn8 = Nothing
	

	'***********************************************************
	'Update SQL with Request Form Field Data
	'***********************************************************

	SQL = "UPDATE  " & MUV_Read("SQL_Owner") & ".Settings_BizIntel SET  "
	SQL = SQL & "CustAnalSum1OnOff = " & CustAnalSum1OnOff & ","
	SQL = SQL & "CustAnalSum1EmailToUserNos = '" & CustAnalSum1EmailToUserNos & "',"
	SQL = SQL & "CustAnalSum1UserNosToCC = '" & CustAnalSum1UserNosToCC & "',"
	SQL = SQL & "MCSUserNosToCC = '" & MCSUserNosToCC & "',"
	SQL = SQL & "MCSActivitySummaryOnOff = " & MCSActivitySummaryOnOff & ","
	SQL = SQL & "MCSActivitySummaryEmailToUserNos = '" & MCSActivitySummaryEmailToUserNos & "',"
	SQL = SQL & "MCSActivitySummaryUserNosToCC = '" & MCSActivitySummaryUserNosToCC & "',"
	SQL = SQL & "MCSUseAlternateHeader = " & MCSUseAlternateHeader
	
	Response.write("<br><br><br>" & SQL)
	
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


	If Request.Form("chkCustAnalSum1OnOff")="on" then chkCustAnalSum1OnOff = "On" Else chkCustAnalSum1OnOff = "Off"
	If CustAnalSum1OnOff_ORIG ="on" then chkCustAnalSum1OnOff_ORIG = "On" Else chkCustAnalSum1OnOff_ORIG = "Off"
	
	If chkCustAnalSum1OnOff <> chkCustAnalSum1OnOff_ORIG Then
		CreateAuditLogEntry "Global Settings Change", "Global Settings Change", "Major", 1, "Biz Intel Automatic Customer Analysis Summary 1 changed from " & chkCustAnalSum1OnOff_ORIG & " to " & chkCustAnalSum1OnOff 
	End If
	
	If CustAnalSum1EmailToUserNos <> CustAnalSum1EmailToUserNos_ORIG Then

		userNames = ""
		userNamesOrig = ""
		
	
		IndividualCustAnalSum1EmailToUserNos = Split(CustAnalSum1EmailToUserNos,",")
		
		for i=0 to Ubound(IndividualCustAnalSum1EmailToUserNos)
		     userNames = userNames & " " & GetUserFirstAndLastNameByUserNo(IndividualCustAnalSum1EmailToUserNos (i))
		next
		
		
		IndividualCustAnalSum1EmailToUserNos_ORIG  = Split(CustAnalSum1EmailToUserNos_ORIG,",")
		
		for i=0 to Ubound(IndividualCustAnalSum1EmailToUserNos_ORIG)
		     userNamesOrig = userNamesOrig & " " & GetUserFirstAndLastNameByUserNo(IndividualCustAnalSum1EmailToUserNos_ORIG(i))
		next
		
		
		CreateAuditLogEntry "Global Settings Change", "Global Settings Change", "Major", 1, "Biz Intel Automatic Customer Analysis Summary 1 send email to changed from " & userNamesOrig & " to " & userNames
	End If


	If CustAnalSum1UserNosToCC <> CustAnalSum1UserNosToCC_ORIG Then

		userNames = ""
		userNamesOrig = ""
		
	
		IndividualCustAnalSum1UserNosToCC  = Split(CustAnalSum1UserNosToCC,",")
		
		for i=0 to Ubound(IndividualCustAnalSum1UserNosToCC )
		     userNames = userNames & " " & GetUserFirstAndLastNameByUserNo(IndividualCustAnalSum1UserNosToCC (i))
		next
		
		
		IndividualCustAnalSum1UserNosToCC_ORIG  = Split(IndividualCustAnalSum1UserNosToCC_ORIG,",")
		
		for i=0 to Ubound(IndividualCustAnalSum1UserNosToCC_ORIG)
		     userNamesOrig = userNamesOrig & " " & GetUserFirstAndLastNameByUserNo(IndividualCustAnalSum1UserNosToCC_ORIG(i))
		next
		
		
		CreateAuditLogEntry "Global Settings Change", "Global Settings Change", "Major", 1, "Biz Intel Automatic Customer Analysis Summary 1 users to Cc changed from " & userNamesOrig & " to " & userNames
	End If

		
		
	IF MCSUserNosToCC <> MCSUserNosToCC_ORIG Then
		CreateAuditLogEntry "Global Settings Change", "Global Settings Change", "Major", 1, "Biz Intel MCS Report user numbers to Cc changed from " & MCSUserNosToCC_ORIG & " to " & MCSUserNosToCC
	End If



'MCS Activity Summary

	If Request.Form("chkMCSActivitySummaryOnOff")="on" then chkMCSActivitySummaryOnOff = "On" Else chkMCSActivitySummaryOnOff = "Off"
	If MCSActivitySummaryOnOff_ORIG ="on" then chkMCSActivitySummaryOnOff_ORIG = "On" Else chkMCSActivitySummaryOnOff_ORIG = "Off"
	
	If chkMCSActivitySummaryOnOff <> chkMCSActivitySummaryOnOff_ORIG Then
		CreateAuditLogEntry "Global Settings Change", "Global Settings Change", "Major", 1, "Biz Intel MCS Activity report changed from " & chkMCSActivitySummaryOnOff_ORIG & " to " & chkMCSActivitySummaryOnOff
	End If
	
	If Request.Form("chkMCSUseAlternateHeader")="on" then chkMCSUseAlternateHeader = "On" Else chkMCSUseAlternateHeader = "Off"
	If MCSUseAlternateHeader_ORIG ="on" then chkMCSUseAlternateHeader_ORIG = "On" Else chkMCSUseAlternateHeader_ORIG = "Off"
	
	If chkMCSUseAlternateHeader <> chkMCSUseAlternateHeader_ORIG Then
		CreateAuditLogEntry "Global Settings Change", "Global Settings Change", "Major", 1, "Biz Intel MCS Activity report use alternate heading format changed from " & chkMCSUseAlternateHeader_ORIG & " to " & chkMCSUseAlternateHeader
	End If

	
	If MCSActivitySummaryEmailToUserNos <> MCSActivitySummaryEmailToUserNos_ORIG Then

		userNames = ""
		userNamesOrig = ""
		
		IndividualMCSActivitySummaryEmailToUserNos = Split(MCSActivitySummaryEmailToUserNos,",")
		
		for i=0 to Ubound(IndividualMCSActivitySummaryEmailToUserNos)
		     userNames = userNames & " " & GetUserFirstAndLastNameByUserNo(IndividualMCSActivitySummaryEmailToUserNos(i))
		next
		
		IndividualMCSActivitySummaryEmailToUserNos_ORIG  = Split(MCSActivitySummaryEmailToUserNos_ORIG,",")
		
		for i=0 to Ubound(IndividualMCSActivitySummaryEmailToUserNos_ORIG)
		     userNamesOrig = userNamesOrig & " " & GetUserFirstAndLastNameByUserNo(IndividualMCSActivitySummaryEmailToUserNos_ORIG(i))
		next
		
		CreateAuditLogEntry "Global Settings Change", "Global Settings Change", "Major", 1, "Biz Intel MCS Activity report send email to changed from " & userNamesOrig & " to " & userNames
		
	End If


	If MCSActivitySummaryUserNosToCC <> MCSActivitySummaryUserNosToCC_ORIG Then

		userNames = ""
		userNamesOrig = ""
		
	
		IndividualMCSActivitySummaryUserNosToCC  = Split(MCSActivitySummaryUserNosToCC,",")
		
		for i=0 to Ubound(IndividualMCSActivitySummaryUserNosToCC)
		     userNames = userNames & " " & GetUserFirstAndLastNameByUserNo(IndividualMCSActivitySummaryUserNosToCC(i))
		next
		
		
		IndividualMCSActivitySummaryUserNosToCC_ORIG  = Split(IndividualMCSActivitySummaryUserNosToCC_ORIG,",")
		
		for i=0 to Ubound(IndividualMCSActivitySummaryUserNosToCC_ORIG)
		     userNamesOrig = userNamesOrig & " " & GetUserFirstAndLastNameByUserNo(IndividualMCSActivitySummaryUserNosToCC_ORIG(i))
		next
		
		
		CreateAuditLogEntry "Global Settings Change", "Global Settings Change", "Major", 1, "Biz Intel MCS Activity report users to Cc changed from " & userNamesOrig & " to " & userNames
	End If


	Response.Redirect("bizintel.asp")
	
%><!--#include file="../../../inc/footer-main.asp"-->