<!--#include file="../inc/InSightFuncs.asp"-->
<!--#include file="../inc/InSightFuncs_Users.asp"-->
<!--#include file="../inc/settings.asp"-->
<!--#include file="../inc/mail.asp"-->
<!--#include file="../inc/InsightFuncs_Service.asp"-->
<%
sURL = Request.ServerVariables("SERVER_NAME")
'baseURL should alwats have a trailing /slash, just in case, handle either way
If right(baseURL,1)="/" Then maildomain = Left(right(baseURL,len(baseURL)-7),len(right(baseURL,len(baseURL)-7))-1) Else maildomain = right(baseURL,len(baseURL)-7)

SendEmail = Request.Form("chkSendEmail")
SendText = Request.Form("chkSendText")
UserToDispatch = Request.Form("selFieldTech")
ServiceTicketNumber = Request.Form("txtServiceTicketNumber")
CustNum = Request.Form("txtAccountNumber")

Response.Write("SendEmail:" & SendEmail & "<br>")
Response.Write("SendText:" & SendText& "<br>")
Response.Write("UserToDispatch :" & UserToDispatch & "<br>")
Response.Write("ServiceTicketNumber :" & ServiceTicketNumber & "<br>")
Response.Write("BaseURL :" & BaseURL & "<br>")
Response.Write(" GetUserDisplayNameByUserNo(UserToDispatch):" &  GetUserDisplayNameByUserNo(UserToDispatch) & "<br>")
Response.Write(" getUserEmailAddress(UserToDispatch):" &  getUserEmailAddress(UserToDispatch) & "<br>")
'Response.End

If UserToDispatch <> 0 then ' If 0, it ia an Undispatch
	SQLDispatch = "INSERT INTO FS_ServiceMemosDetail (MemoNumber, CustNum, MemoStage, "
	SQLDispatch = SQLDispatch & "UserNoOfServiceTech, SubmissionDateTime, USerNoSubmittingRecord,EmailAddressSentTo,TextNumberSentTo,OriginalDispatchDateTime,Remarks)"
	SQLDispatch = SQLDispatch &  " VALUES (" 
	SQLDispatch = SQLDispatch & "'"  & ServiceTicketNumber & "'"
	SQLDispatch = SQLDispatch & ",'"  & CustNum & "'"
	SQLDispatch = SQLDispatch & ",'Dispatched'"
	SQLDispatch = SQLDispatch & ","  & UserToDispatch 
	SQLDispatch = SQLDispatch & ",getdate() "
	SQLDispatch = SQLDispatch & ","  & Session("UserNo")
	SQLDispatch = SQLDispatch & ",'"  & getUserEmailAddress(UserToDispatch) & "'"
	SQLDispatch = SQLDispatch & ",'" & getUserCellNumber(UserToDispatch) & "' "
	SQLDispatch = SQLDispatch & ", getDate()"	
	SQLDispatch = SQLDispatch & ",'" &  GetUserDisplayNameByUserNo(UserToDispatch) & " has been dispatched. This ticket was reassigned via the dispatch center by " & GetUserDisplayNameByUserNo(Session("UserNo")) & "')"	
Else
	SQLDispatch = "INSERT INTO FS_ServiceMemosDetail (MemoNumber, CustNum, MemoStage, "
	SQLDispatch = SQLDispatch & "UserNoOfServiceTech, SubmissionDateTime, USerNoSubmittingRecord,EmailAddressSentTo,TextNumberSentTo,OriginalDispatchDateTime,Remarks)"
	SQLDispatch = SQLDispatch &  " VALUES (" 
	SQLDispatch = SQLDispatch & "'"  & ServiceTicketNumber & "'"
	SQLDispatch = SQLDispatch & ",'"  & CustNum & "'"
	SQLDispatch = SQLDispatch & ",'Received'"
	SQLDispatch = SQLDispatch & ","  & UserToDispatch 
	SQLDispatch = SQLDispatch & ",getdate() "
	SQLDispatch = SQLDispatch & ","  & Session("UserNo")
	SQLDispatch = SQLDispatch & ",'"  & getUserEmailAddress(UserToDispatch) & "'"
	SQLDispatch = SQLDispatch & ",'" & getUserCellNumber(UserToDispatch) & "' "
	SQLDispatch = SQLDispatch & ", getDate()"	
	SQLDispatch = SQLDispatch & ",'This ticket was changed from " &  GetServiceTicketCurrentStage(ServiceTicketNumber) & " to un-dispatched via the dispatch center by " & GetUserDisplayNameByUserNo(Session("UserNo")) & "')"	
End If

HeldStatus = GetServiceTicketCurrentStage(ServiceTicketNumber)
HeldTech = GetServiceTicketDispatchedTech(ServiceTicketNumber)
Response.Write("<br><br>HeldTech1: " & HeldTech & "<br><br><br>")

Set cnnDispatch = Server.CreateObject("ADODB.Connection")
cnnDispatch.open (Session("ClientCnnString"))
Set rsDispatch = Server.CreateObject("ADODB.Recordset")
Set rsDispatch = cnnDispatch.Execute(SQLDispatch)


'Write audit trail for dispatch
'*******************************
If UserToDispatch <> 0 Then
	Description = GetUserDisplayNameByUserNo(UserToDispatch) & " was dispatched via the dispatch center to service ticket number " & ServiceTicketNumber & " by " & GetUserDisplayNameByUserNo(Session("UserNo")) & " at " & NOW()
Else
	Description = "Service ticket " & ServiceTicketNumber  & " was changed from " &  HeldStatus  & " to un-dispatched via the dispatch center by " & GetUserDisplayNameByUserNo(Session("UserNo")) & " at " & NOW()
End If
CreateAuditLogEntry "Dispatch Center","Dispatch Center","Major",0,Description 

If UserToDispatch <> 0 Then
'Also set dispatched flag for simple dispatch model
	SQLDispatch = "Update FS_ServiceMemos Set Dispatched = CASE WHEN Dispatched = 0 THEN -1 ELSE 0 END Where MemoNumber = '"  & ServiceTicketNumber & "'"
	Set rsDispatch = cnnDispatch.Execute(SQLDispatch)
End If

Set rsDispatch = Nothing
cnnDispatch.Close
Set cnnDispatch = Nothing


'Send email is necessary
'***********************
		
If SendEmail="on" then
	If UserToDispatch <> 0 Then ' Normal xfer not UNDispatch
		If getUserEmailAddress(UserToDispatch) <> "" Then 
			Send_To = getUserEmailAddress(UserToDispatch)
			%>
			<!--#include file="../emails/ADVdispatch_dispatch.asp"-->
			<%
			'Failsafe for dev
			If Instr(ucase(sURL),"DEV") <> 0 Then Send_To = "rich@ocsaccess.com"
			SendMail "mailsender@" & maildomain ,Send_To,emailSubject,emailBody,GetTerm("Service"),"Service Dispatch"
			Description = "A dispatch email was sent to " & GetUserDisplayNameByUserNo(Session("UserNo")) & " (" & Send_To & ") at " & NOW()
			CreateAuditLogEntry "Service Ticket System","Dispatch email sent","Minor",0,Description
		Else
			' Could not send dispatch email, no address on file
			emailBody = "Insight was unable to send a dispatch email to " & GetUserDisplayNameByUserNo(UserToDispatch) & ". No email address on file"
			If Instr(ucase(sURL),"DEV") <> 0 Then SEND_TO = "rich@ocsaccess.com"
			SendMail "mailsender@" & maildomain ,SEND_TO,"Unable to send dispatch email",emailBody,GetTerm("Service"),"Missing Email"
			Description = "Insight was unable to send a dispatch email to " & GetUserDisplayNameByUserNo(UserToDispatch) & ". No email address on file"
			CreateAuditLogEntry "Service Ticket System","Unable to send dispatch email","Major",0,Description
		End If
	End If
	
	If UserToDispatch = 0 Then ' This is an undispatch
		If getUserEmailAddress(HeldTech) <> "" Then 
			Send_To = getUserEmailAddress(HeldTech)
			%>
			<!--#include file="../emails/ADVdispatch_cancel.asp"-->
			<%
			'Failsafe for dev
			If Instr(ucase(sURL),"DEV") <> 0 Then Send_To = "rich@ocsaccess.com"
			SendMail "mailsender@" & maildomain ,Send_To,emailSubject,emailBody,GetTerm("Service"),"Cancel Dispatch"
			Description = "A cancel dispatch email was sent to " & GetUserDisplayNameByUserNo(HeldTech) & " (" & Send_To & ") at " & NOW()
			CreateAuditLogEntry "Service Ticket System","Cancel dispatch email sent","Minor",0,Description
		Else
			' Could not send dispatch email, no address on file
			emailBody = "Insight was unable to send a cancel dispatch email to " & GetUserDisplayNameByUserNo(HeldTech) & ". No email address on file"
			If Instr(ucase(sURL),"DEV") <> 0 Then SEND_TO = "rich@ocsaccess.com"
			SendMail "mailsender@" & maildomain ,SEND_TO,"Unable to send dispatch email",emailBody,GetTerm("Service"),"Missing Email"
			Description = "Insight was unable to send a dispatch email to " & GetUserDisplayNameByUserNo(HeldTech) & ". No email address on file"
			CreateAuditLogEntry "Service Ticket System","Unable to send dispatch email","Major",0,Description
		End If
	End If
End If


'Send text if necessary
'**********************
If SendText="on" then

	'See if ACK link should be included
	DLinkInText = False
	SQLtxt = "SELECT  IncludeACKInDispatchText FROM Settings_EmailService"
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
			QString = QString & "&R=Location: " & Server.URLEncode(BaseURL & "service/serviceBoard.asp")
			
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
			If Instr(ucase(sURL),"DEV") <> 0 Then SEND_TO = "rich@ocsaccess.com"
			SendMail "mailsender@" & maildomain ,SEND_TO,"Unable to send dispatch text message (" & MUV_READ("SERNO") & ")",emailBody,GetTerm("Service"),"Missing Email"
			Description = "Insight was unable to send a dispatch text message to " & GetUserDisplayNameByUserNo(UserToDispatch) & ". No cell number on file"
			CreateAuditLogEntry "Service Ticket System","Unable to send dispatch text message","Major",0,Description
		End If
	End If
	
	If UserToDispatch <> 0 Then
	
		Response.Write("<br><br>HeldTech: " & HeldTech & "<br><br><br>")
	
		If getUserCellNumber(HeldTech) <> "" Then
			
			Send_To = getUserCellNumber(HeldTech)
	
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
	
			Response.Redirect (URL & Qstring)
	
			Description = "A cancel dispatch text message was sent to " & GetUserDisplayNameByUserNo(HeldTech) & " (" & getUserCellNumber(HeldTech) & ") at " & NOW()		
			CreateAuditLogEntry "Service Ticket System","Dispatch email sent","Minor",0,Description
		Else
			' Could not send dispatch test, no address on file
			emailBody = "Insight was unable to send a dispatch text message to " & GetUserDisplayNameByUserNo(HeldTech) & ". No cell number on file"
		
			If Instr(ucase(sURL),"DEV") <> 0 Then SEND_TO = "rich@ocsaccess.com"
			SendMail "mailsender@" & maildomain ,SEND_TO,"Unable to send dispatch text message (" & MUV_READ("SERNO") & ")",emailBody,,GetTerm("Service"),"Missing Cell Number"
			Description = "Insight was unable to send a dispatch text message to " & GetUserDisplayNameByUserNo(HeldTech) & ". No cell number on file"		
			CreateAuditLogEntry "Service Ticket System","Unable to send dispatch text message","Major",0,Description
		
		End If
	End If
End If

dummy=RemoveFromRedispatch(ServiceTicketNumber)

Response.Redirect(BaseURL & "service/serviceBoard.asp")
%>

 
