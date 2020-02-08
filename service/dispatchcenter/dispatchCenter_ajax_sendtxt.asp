<!--#include file="../../inc/InSightFuncs.asp"-->
<!--#include file="../../inc/InSightFuncs_Users.asp"-->
<!--#include file="../../inc/settings.asp"-->
<!--#include file="../../inc/mail.asp"-->
<!--#include file="../../inc/InsightFuncs_Service.asp"-->
<%
sURL = Request.ServerVariables("SERVER_NAME")
'baseURL should alwats have a trailing /slash, just in case, handle either way
If right(baseURL,1)="/" Then maildomain = Left(right(baseURL,len(baseURL)-7),len(right(baseURL,len(baseURL)-7))-1) Else maildomain = right(baseURL,len(baseURL)-7)

SendText = Request("chkSendText")
UserToDispatch = Request("selFieldTech")
ServiceTicketNumber = Request("txtServiceTicketNumber")
CustNum = GetServiceTicketCust(ServiceTicketNumber)

'CustNum = Request("txtAccountNumber") ' old code


'Send text if necessary
'**********************


	'See if ACK link should be included
	DLinkInText = False
	SQLtxt = "SELECT IncludeACKInDispatchText FROM Settings_EmailService"
	Set cnntxt = Server.CreateObject("ADODB.Connection")
	cnntxt.open (Session("ClientCnnString"))
	Set rstxt = Server.CreateObject("ADODB.Recordset")
	rstxt.CursorLocation = 3 
	Set rstxt = cnntxt.Execute(SQLtxt)
	If not rstxt.EOF Then DLinkInText = rstxt("IncludeACKInDispatchText")
	set rstxt = Nothing
	cnntxt.close
	set cnntxt = Nothing
	
	If UserToDispatch <> 0 Then
	
		If getUserCellNumber(UserToDispatch) <> ""  Then
		
			Send_To = getUserCellNumber(UserToDispatch)
	
			URL = BaseURL & "inc/sendtext.php"
	
			QString = "?n=" & Replace(getUserCellNumber(UserToDispatch),"-","")
			
			QString = QString & "&u1=" & EzTextingUserID()
			QString = QString & "&u2=" & EzTextingPassword()
		
			QString = QString & "&t=DISPATCH"
			QString = QString & "&R=Location: " & Server.URLEncode(BaseURL & "service/dispatchCenter/main.asp")
			
			If GetCustNameByCustNum(CustNum) <> "" Then
				QString = QString & "&m=" & GetTerm("Account") & ": " & EZTexting_Filter1(Replace(GetCustNameByCustNum(CustNum),"&"," "))
			Else
				QString = QString & "&m=" & GetTerm("Account") & ": " &  CustNum 
				QString = QString &  "   Ticket:" & ServiceTicketNumber
			End If
	
		
			If DLinkInText = 1 Then
				QString = QString & "    Tap the link for more info "
				QString = QString & Server.URLEncode(baseURL & "directlaunch/service/moreinfo_dispatch_from_email_or_text.asp?t=" & ServiceTicketNumber & "&u=" & UserToDispatch & "&c=" & CustNum & "&cl=" & MUV_READ("SERNO"))
			End If
	
			QString = QString &  "&cty=" & GetCompanyCountry()
			QString = Replace(Qstring," ", "%20")
			
		
			Response.Redirect (URL & Qstring)
	
			Description = "A dispatch text message was sent to " & GetUserDisplayNameByUserNo(UserToDispatch) & " (" & getUserCellNumber(UserToDispatch) & ") at " & NOW()
			CreateAuditLogEntry "Service Ticket System","Dispatch email sent","Minor",0,Description
		Else
			' Could not send dispatch test, no address on file
			emailBody = "Insight was unable to send a dispatch text message to " & GetUserDisplayNameByUserNo(UserToDispatch) & ". No cell number on file"
			If Instr(ucase(sURL),"DEV") <> 0 Then SEND_TO = "rich@ocsaccess.com" else SEND_TO = "rich@ocsaccess.com"
			SendMail "mailsender@" & maildomain ,SEND_TO,"Unable to send dispatch text message (" & MUV_READ("SERNO") & ")",emailBody,GetTerm("Service"),"Missing Email"
			Description = "Insight was unable to send a dispatch text message to " & GetUserDisplayNameByUserNo(UserToDispatch) & ". No cell number on file"
			CreateAuditLogEntry "Service Ticket System","Unable to send dispatch text message","Major",0,Description
		End If
	End If
	
	If UserToDispatch <> 0 Then
		
		If getUserCellNumber(HeldTech) <> "" Then
			
			Send_To = getUserCellNumber(UserToDispatch)
	
			URL = BaseURL & "inc/sendtext.php"
		
			QString = "?n=" & Replace(getUserCellNumber(HeldTech),"-","")
			
			QString = QString & "&u1=" & EzTextingUserID()
			QString = QString & "&u2=" & EzTextingPassword()
		
			QString = QString & "&t=CANCELLED"
			QString = QString & "&R=Location: " & Server.URLEncode(BaseURL & "service/main.asp")
			QString = QString & "&m=" & GetTerm("Account") & ":" &  Account
			QString = QString &  "   Ticket:" & ServiceTicketNumber
			QString = QString &  "&cty=" & GetCompanyCountry()					
			QString = Replace(Qstring," ", "%20")
	
			
	
			Description = "A cancel dispatch text message was sent to " & GetUserDisplayNameByUserNo(HeldTech) & " (" & getUserCellNumber(HeldTech) & ") at " & NOW()		
			CreateAuditLogEntry "Service Ticket System","Dispatch email sent","Minor",0,Description
		Else
			' Could not send dispatch test, no address on file
			emailBody = "Insight was unable to send a dispatch text message to " & GetUserDisplayNameByUserNo(HeldTech) & ". No cell number on file"
		
			If Instr(ucase(sURL),"DEV") <> 0 Then SEND_TO = "rich@ocsaccess.com" else SEND_TO = "rich@ocsaccess.com"
			SendMail "mailsender@" & maildomain ,SEND_TO,"Unable to send dispatch text message (" & MUV_READ("SERNO") & ")",emailBody,,GetTerm("Service"),"Missing Cell Number"
			Description = "Insight was unable to send a dispatch text message to " & GetUserDisplayNameByUserNo(HeldTech) & ". No cell number on file"		
			CreateAuditLogEntry "Service Ticket System","Unable to send dispatch text message","Major",0,Description
		
		End If
	End If


response.write("0")
%>

 
