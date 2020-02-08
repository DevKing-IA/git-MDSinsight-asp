<!--#include file="../../../../inc/header.asp"-->

<%

	'***********************************************************
	'Get Values Of All Form Fields Posted
	'***********************************************************
		
	OrderAPIRepostURL = Request.Form("txtOrderAPIRepostURL")
	If Request.Form("chkOrderAPIRepostONOFF") = "on" then OrderAPIRepostONOFF = 1 Else OrderAPIRepostONOFF = 0
	OrderAPIRepostMode = Request.Form("selOrderAPIRepostMode")
	OrderAPIOffsetDays = Request.Form("selOrderAPIOffsetDays")
	If Request.Form("chkOrderAPISwapAddressLines") = "on" then OrderAPISwapAddressLines = 1 Else OrderAPISwapAddressLines = 0
	
	InvoiceAPIRepostURL = Request.Form("txtInvoiceAPIRepostURL")
	If Request.Form("chkInvoiceAPIRepostONOFF") = "on" then InvoiceAPIRepostONOFF = 1 Else InvoiceAPIRepostONOFF = 0
	InvoiceAPIRepostMode = Request.Form("selInvoiceAPIRepostMode")
	InvoiceAPIOffsetDays = Request.Form("selInvoiceAPIOffsetDays")
	SendInvoiceType = Request.Form("selSendInvoiceType")
	
	CMAPIRepostURL = Request.Form("txtCMAPIRepostURL")
	If Request.Form("chkCMAPIRepostONOFF") = "on" then CMAPIRepostONOFF = 1 Else CMAPIRepostONOFF = 0
	CMAPIRepostMode = Request.Form("selCMAPIRepostMode")
	CMAPIOffsetDays = Request.Form("selCMAPIOffsetDays")
		
	RAAPIRepostURL = Request.Form("txtRAAPIRepostURL")
	If Request.Form("chkRAAPIRepostONOFF") = "on" then RAAPIRepostONOFF = 1 Else RAAPIRepostONOFF = 0
	RAAPIOffsetDays = Request.Form("selRAAPIOffsetDays")
	RAAPIRepostMode = Request.Form("selRAAPIRepostMode")
		
	SumInvAPIRepostURL = Request.Form("txtSumInvAPIRepostURL")
	If Request.Form("chkSumInvAPIRepostONOFF") = "on" then SumInvAPIRepostONOFF = 1 Else SumInvAPIRepostONOFF = 0
	SumInvAPIRepostMode = Request.Form("selSumInvAPIRepostMode")
	SumInvAPIOffsetDays = Request.Form("selSumInvAPIOffsetDays")
	
	If Request.Form("chkAPIDailyActivityReportOnOff") = "on" then APIDailyActivityReportOnOff = 1 Else APIDailyActivityReportOnOff = 0
	APIDailyActivityReportAdditionalEmails = Request.Form("txtAPIDailyActivityReportAdditionalEmails")
	APIDailyActivityReportEmailSubject = Request.Form("txtAPIDailyActivityReportEmailSubject")
	APIDailyActivityReportUserNos = Request.Form("lstSelectedAPIDailyActivityReportUserNos")
	
	OrderCutoffTime = Request.Form("selOrderCutoffTime")
	InvoiceCutoffTime = Request.Form("selInvoiceCutoffTime")
	RACutoffTime = Request.Form("selRACutoffTime")
	CMCutoffTime = Request.Form("selCMCutoffTime")
	
	
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
		OrderAPIRepostMode_ORIG = rs("OrderAPIRepostMode")	
		OrderAPIOffsetDays_ORIG = rs("OrderAPIOffsetDays")	
		InvoiceAPIRepostMode_ORIG = rs("InvoiceAPIRepostMode")	
		InvoiceAPIOffsetDays_ORIG = rs("InvoiceAPIOffsetDays")	
		SendInvoiceType_ORIG = rs("SendInvoiceType")	
		RAAPIRepostMode_ORIG = rs("RAAPIRepostMode")	
		RAAPIOffsetDays_ORIG = rs("RAAPIOffsetDays")	
		CMAPIRepostMode_ORIG = rs("CMAPIRepostMode")	
		CMAPIOffsetDays_ORIG = rs("CMAPIOffsetDays")	
		SumInvAPIRepostMode_ORIG = rs("SumInvAPIRepostMode")	
		SumInvAPIOffsetDays_ORIG = rs("SumInvAPIOffsetDays")	
		OrderAPIRepostURL_ORIG = rs("OrderAPIRepostURL")	
		OrderAPIRepostONOFF_ORIG = rs("OrderAPIRepostONOFF")	
		InvoiceAPIRepostURL_ORIG = rs("InvoiceAPIRepostURL")	
		InvoiceAPIRepostONOFF_ORIG = rs("InvoiceAPIRepostONOFF")	
		RAAPIRepostURL_ORIG = rs("RAAPIRepostURL")	
		RAAPIRepostONOFF_ORIG = rs("RAAPIRepostONOFF")	
		CMAPIRepostURL_ORIG = rs("CMAPIRepostURL")	
		CMAPIRepostONOFF_ORIG = rs("CMAPIRepostONOFF")	
		SumInvAPIRepostURL_ORIG = rs("SumInvAPIRepostURL")	
		SumInvAPIRepostONOFF_ORIG = rs("SumInvAPIRepostONOFF")	
		APIDailyActivityReportOnOff_ORIG	= rs("APIDailyActivityReportOnOff")			
		APIDailyActivityReportAdditionalEmails_ORIG = rs("APIDailyActivityReportAdditionalEmails")	
		APIDailyActivityReportEmailSubject_ORIG = rs("APIDailyActivityReportEmailSubject")	
		APIDailyActivityReportUserNos_ORIG = rs("APIDailyActivityReportUserNos")	
		OrderCutoffTime_ORIG = rs("OrderCutoffTime")
		InvoiceCutoffTime_ORIG = rs("InvoiceCutoffTime")
		RACutoffTime_ORIG = rs("RACutoffTime")
		CMCutoffTime_ORIG = rs("CMCutoffTime")		
		OrderAPISwapAddressLines_ORIG = rs("OrderAPISwapAddressLines")
					
	End If
	
	
	set rs = Nothing
	cnn8.close
	set cnn8 = Nothing
	

	'***********************************************************
	'Perform Audit Trail Entry Inserts
	'***********************************************************


	If Request.Form("chkOrderAPIRepostONOFF")="on" then OrderAPIRepostONOFFMsg = "On" Else OrderAPIRepostONOFFMsg = "Off"
	
	IF OrderAPIRepostONOFF <> OrderAPIRepostONOFF_ORIG Then
		CreateAuditLogEntry "Global Settings Change", "Global Settings Change", "Major", 1, "Order API repost orders changed from " & OrderAPIRepostONOFF_ORIG & " to " & OrderAPIRepostONOFFMsg 
	End If
	If OrderAPIRepostURL <> OrderAPIRepostURL_ORIG Then
		CreateAuditLogEntry "Global Settings Change", "Global Settings Change", "Major", 1, "Order API posting URL changed from " & OrderAPIRepostURL_ORIG & " to " & OrderAPIRepostURL 
	End If
	
	If Request.Form("chkInvoiceAPIRepostONOFF")="on" then InvoiceAPIRepostONOFFMsg = "On" Else InvoiceAPIRepostONOFFMsg = "Off"
	
	IF InvoiceAPIRepostONOFF <> InvoiceAPIRepostONOFF_ORIG Then
		CreateAuditLogEntry "Global Settings Change", "Global Settings Change", "Major", 1, "Invoice API repost invoices changed from " & InvoiceAPIRepostONOFF_ORIG & " to " & InvoiceAPIRepostONOFFMsg 
	End If
	If InvoiceAPIRepostURL <> InvoiceAPIRepostURL_ORIG Then
		CreateAuditLogEntry "Global Settings Change", "Global Settings Change", "Major", 1, "Invoice API posting URL changed from " & InvoiceAPIRepostURL_ORIG & " to " & InvoiceAPIRepostURL 
	End If			
	
	If Request.Form("chkRAAPIRepostONOFF")="on" then RAAPIRepostONOFFMsg = "On" Else RAAPIRepostONOFFMsg = "Off"
	
	IF RAAPIRepostONOFF <> RAAPIRepostONOFF_ORIG Then
		CreateAuditLogEntry "Global Settings Change", "Global Settings Change", "Major", 1, "RA API repost return authorizations changed from " & RAAPIRepostONOFF_ORIG & " to " & RAAPIRepostONOFFMsg 
	End If
	If RAAPIRepostURL <> RAAPIRepostURL_ORIG Then
		CreateAuditLogEntry "Global Settings Change", "Global Settings Change", "Major", 1, "RA API posting URL changed from " & RAAPIRepostURL_ORIG & " to " & RAAPIRepostURL 
	End If
	
	If Request.Form("chkCMAPIRepostONOFF")="on" then CMAPIRepostONOFFMsg = "On" Else CMAPIRepostONOFFMsg = "Off"
	
	IF CMAPIRepostONOFF <> CMAPIRepostONOFF_ORIG Then
		CreateAuditLogEntry "Global Settings Change", "Global Settings Change", "Major", 1, "CM API repost credit memos changed from " & CMAPIRepostONOFF_ORIG & " to " & CMAPIRepostONOFFMsg 
	End If
	If CMAPIRepostURL <> CMAPIRepostURL_ORIG Then
		CreateAuditLogEntry "Global Settings Change", "Global Settings Change", "Major", 1, "CM API posting URL changed from " & CMAPIRepostURL_ORIG & " to " & CMAPIRepostURL 
	End If
	
	If Request.Form("chkSumInvAPIRepostONOFF")="on" then SumInvAPIRepostONOFFMsg = "On" Else SumInvAPIRepostONOFFMsg = "Off"
	
	IF SumInvAPIRepostONOFF <> SumInvAPIRepostONOFF_ORIG Then
		CreateAuditLogEntry "Global Settings Change", "Global Settings Change", "Major", 1, "Summary Invoice API repost summary invoices changed from " & SumInvAPIRepostONOFF_ORIG & " to " & SumInvAPIRepostONOFFMsg 
	End If
	If SumInvAPIRepostURL <> SumInvAPIRepostURL_ORIG Then
		CreateAuditLogEntry "Global Settings Change", "Global Settings Change", "Major", 1, "Summary Invoice API posting URL changed from " & SumInvAPIRepostURL_ORIG & " to " & SumInvAPIRepostURL 
	End If
	
	
	If OrderAPIRepostMode <> OrderAPIRepostMode_ORIG Then
		CreateAuditLogEntry "Global Settings Change", "Global Settings Change", "Major", 1, "Order API Post Mode changed from " & OrderAPIRepostMode_ORIG & " to " & OrderAPIRepostMode 
	End If
	If cint(OrderAPIOffsetDays) <> cint(OrderAPIOffsetDays_ORIG) Then
		CreateAuditLogEntry "Global Settings Change", "Global Settings Change", "Major", 1, "Order API, Include orders offset by, changed from  " & OrderAPIOffsetDays_ORIG & " to " & OrderAPIOffsetDays & " days."
	End If
	
	If InvoiceAPIRepostMode <> InvoiceAPIRepostMode_ORIG Then
		CreateAuditLogEntry "Global Settings Change", "Global Settings Change", "Major", 1, "Invoice API Post Mode changed from " & InvoiceAPIRepostMode_ORIG & " to " & InvoiceAPIRepostMode
	End If
	If SendInvoiceType <> SendInvoiceType_ORIG Then
		CreateAuditLogEntry "Global Settings Change", "Global Settings Change", "Major", 1, "Invoice API Send Type changed from " & SendInvoiceType_ORIG & " to " & SendInvoiceType
	End If
	
	If cint(InvoiceAPIOffsetDays) <> cint(InvoiceAPIOffsetDays_ORIG) Then
		CreateAuditLogEntry "Global Settings Change", "Global Settings Change", "Major", 1, "Invoice API, Include invoices offset by, changed from  " & InvoiceAPIOffsetDays_ORIG & " to " & InvoiceAPIOffsetDays & " days."
	End If

	If RAAPIRepostMode <> RAAPIRepostMode_ORIG Then
		CreateAuditLogEntry "Global Settings Change", "Global Settings Change", "Major", 1, "Return Authorization API Post Mode changed from " & RAAPIRepostMode_ORIG & " to " & RAAPIRepostMode
	End If
	If cint(RAAPIOffsetDays) <> cint(RAAPIOffsetDays_ORIG) Then
		CreateAuditLogEntry "Global Settings Change", "Global Settings Change", "Major", 1, "Return Authorization API, Include RAs offset by, changed from  " & RAAPIOffsetDays_ORIG & " to " & RAAPIOffsetDays & " days."
	End If

	If CMAPIRepostMode <> CMAPIRepostMode_ORIG Then
		CreateAuditLogEntry "Global Settings Change", "Global Settings Change", "Major", 1, "Credit Memo API Post Mode changed from " & CMAPIRepostMode_ORIG & " to " & CMAPIRepostMode
	End If
	If cint(CMAPIOffsetDays) <> cint(CMAPIOffsetDays_ORIG) Then
		CreateAuditLogEntry "Global Settings Change", "Global Settings Change", "Major", 1, "Credit Memo API, Include CM's offset by, changed from  " & CMAPIOffsetDays_ORIG & " to " & CMAPIOffsetDays & " days."
	End If

	If SumInvAPIRepostMode <> SumInvAPIRepostMode_ORIG Then
		CreateAuditLogEntry "Global Settings Change", "Global Settings Change", "Major", 1, "Summary Invoices API Post Mode changed from " & SumInvAPIRepostMode_ORIG & " to " & SumInvAPIRepostMode 
	End If
	If cint(SumInvAPIOffsetDays) <> cint(SumInvAPIOffsetDays_ORIG) Then
		CreateAuditLogEntry "Global Settings Change", "Global Settings Change", "Major", 1, "Summary Invoices API, Include Sum. Inv. offset by, changed from  " & SumInvAPIOffsetDays_ORIG & " to " & SumInvAPIOffsetDays & " days."
	End If

	If OrderCutoffTime <> OrderCutoffTime_ORIG Then
		CreateAuditLogEntry "Global Settings Change", "Global Settings Change", "Major", 1, "Order API order cutoff time changed from " & OrderCutoffTime_ORIG & " to " & OrderCutoffTime
	End If
	If InvoiceCutoffTime <> InvoiceCutoffTime_ORIG Then
		CreateAuditLogEntry "Global Settings Change", "Global Settings Change", "Major", 1, "Order API invoice cutoff time changed from " & InvoiceCutoffTime_ORIG & " to " & InvoiceCutoffTime
	End If
	If RACutoffTime <> RACutoffTime_ORIG Then
		CreateAuditLogEntry "Global Settings Change", "Global Settings Change", "Major", 1, "Order API return authorization cutoff time changed from " & RACutoffTime_ORIG & " to " & RACutoffTime
	End If
	If CMCutoffTime <> CMCutoffTime_ORIG Then
		CreateAuditLogEntry "Global Settings Change", "Global Settings Change", "Major", 1, "Order API credit memo cutoff time changed from " & CMCutoffTime_ORIG & " to " & CMCutoffTime
	End If


	If Request.Form("chkAPIDailyActivityReportOnOff")="on" then APIDailyActivityReportOnOfftxt = "On" Else APIDailyActivityReportOnOfftxt = "Off"
	
	IF APIDailyActivityReportOnOff <> APIDailyActivityReportOnOff_ORIG Then
		CreateAuditLogEntry "Global Settings Change", "Global Settings Change", "Major", 1, "Email API Daily Activity Report changed from " & APIDailyActivityReportOnOff_ORIG & " to " & APIDailyActivityReportOnOfftxt 
	End If

	If APIDailyActivityReportEmailSubject <> APIDailyActivityReportEmailSubject_ORIG Then
		CreateAuditLogEntry "Global Settings Change", "Global Settings Change", "Major", 1, "API Daily Activity Report Email Subject changed from " & APIDailyActivityReportEmailSubject_ORIG & " to " & APIDailyActivityReportEmailSubject
	End If

	If APIDailyActivityReportAdditionalEmails <> APIDailyActivityReportAdditionalEmails_ORIG Then
		CreateAuditLogEntry "Global Settings Change", "Global Settings Change", "Major", 1, "API Daily Activity Report Additional Emails changed from " & APIDailyActivityReportAdditionalEmails_ORIG & " to " & APIDailyActivityReportAdditionalEmails
	End If

'''''''''''''''''''''''''''''''''''''''''''''''''
	If Request.Form("chkOrderAPISwapAddressLines")="on" then OrderAPISwapAddressLinesMsg = "On" Else OrderAPISwapAddressLinesMsg = "Off"
	
	IF OrderAPISwapAddressLines <> OrderAPISwapAddressLines_ORIG Then
		CreateAuditLogEntry "Global Settings Change", "Global Settings Change", "Major", 1, "Order API swap address lines changed from " & OrderAPISwapAddressLines_ORIG & " to " & OrderAPISwapAddressLinesMsg 
	End If



''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

	If APIDailyActivityReportUserNos <> APIDailyActivityReportUserNos_ORIG Then

		userNames = ""
		userNamesOrig = ""
	
		IndividualAPIDailyActivityReportUserNos = Split(APIDailyActivityReportUserNos,",")
		
		for i = 0 to Ubound(IndividualAPIDailyActivityReportUserNos)
		     userNames = userNames & " " & GetUserFirstAndLastNameByUserNo(IndividualAPIDailyActivityReportUserNos(i))
		next 
		
		IndividualAPIDailyActivityReportUserNos_ORIG  = Split(APIDailyActivityReportUserNos_ORIG,",")
		
		for i = 0 to Ubound(IndividualAPIDailyActivityReportUserNos_ORIG)
		     userNamesOrig = userNamesOrig & " " & GetUserFirstAndLastNameByUserNo(IndividualAPIDailyActivityReportUserNos_ORIG(i))
		next 
		
		
		CreateAuditLogEntry "Global Settings Change", "Global Settings Change", "Major", 1, "API Daily Activity Report Additional Send Report To changed from " & userNamesOrig & " to " & userNames
	End If

	'***********************************************************
	'Update SQL with Request Form Field Data
	'***********************************************************

	SQL = "UPDATE  " & MUV_Read("SQL_Owner") & ".Settings_Global SET  "
	SQL = SQL & "APIDailyActivityReportOnOff = " & APIDailyActivityReportOnOff & ","
	SQL = SQL & "APIDailyActivityReportAdditionalEmails	= '" & APIDailyActivityReportAdditionalEmails	& "',"
	SQL = SQL & "APIDailyActivityReportEmailSubject	= '" & APIDailyActivityReportEmailSubject	& "',"
	SQL = SQL & "APIDailyActivityReportUserNos = '" & APIDailyActivityReportUserNos	& "',"
	SQL = SQL & "OrderAPIRepostURL = '" & OrderAPIRepostURL & "',"
	SQL = SQL & "OrderAPIRepostONOFF = " & OrderAPIRepostONOFF & ","
	SQL = SQL & "InvoiceAPIRepostURL = '" & InvoiceAPIRepostURL & "',"
	SQL = SQL & "InvoiceAPIRepostONOFF = " & InvoiceAPIRepostONOFF & ","
	SQL = SQL & "RAAPIRepostURL = '" & RAAPIRepostURL & "',"
	SQL = SQL & "RAAPIRepostONOFF = " & RAAPIRepostONOFF & ","
	SQL = SQL & "CMAPIRepostURL = '" & CMAPIRepostURL & "',"
	SQL = SQL & "CMAPIRepostONOFF = " & CMAPIRepostONOFF & ","
	SQL = SQL & "SumInvAPIRepostURL = '" & SumInvAPIRepostURL & "',"
	SQL = SQL & "SumInvAPIRepostONOFF = " & SumInvAPIRepostONOFF & ","
	SQL = SQL & "OrderAPIRepostMode = '" & OrderAPIRepostMode & "',"
	SQL = SQL & "OrderAPIOffsetDays = " & OrderAPIOffsetDays & ","
	SQL = SQL & "InvoiceAPIRepostMode = '" & InvoiceAPIRepostMode & "',"
	SQL = SQL & "InvoiceAPIOffsetDays = " & InvoiceAPIOffsetDays & ","
	SQL = SQL & "SendInvoiceType = '" & SendInvoiceType & "',"
	SQL = SQL & "RAAPIRepostMode = '" & RAAPIRepostMode & "',"
	SQL = SQL & "RAAPIOffsetDays = " & RAAPIOffsetDays & ","
	SQL = SQL & "CMAPIRepostMode = '" & CMAPIRepostMode & "',"
	SQL = SQL & "CMAPIOffsetDays = " & CMAPIOffsetDays & ","
	SQL = SQL & "SumInvAPIRepostMode = '" & SumInvAPIRepostMode & "',"
	SQL = SQL & "SumInvAPIOffsetDays = " & SumInvAPIOffsetDays & ","
	SQL = SQL & "OrderCutoffTime = '" & OrderCutoffTime & "',"
	SQL = SQL & "InvoiceCutoffTime = '" & InvoiceCutoffTime & "',"
	SQL = SQL & "RACutoffTime = '" & RACutoffTime & "',"
	SQL = SQL & "CMCutoffTime = '" & CMCutoffTime & "'"
	SQL = SQL & ",OrderAPISwapAddressLines = " & OrderAPISwapAddressLines

						    
	
	Set cnn8 = Server.CreateObject("ADODB.Connection")
	cnn8.open (Session("ClientCnnString"))
	Set rs = Server.CreateObject("ADODB.Recordset")
	rs.CursorLocation = 3 
	
	Set rs = cnn8.Execute(SQL)

	set rs = Nothing
	cnn8.close
	set cnn8 = Nothing
	
	

	Response.Redirect("order-api.asp")
%><!--#include file="../../../../inc/footer-main.asp"-->