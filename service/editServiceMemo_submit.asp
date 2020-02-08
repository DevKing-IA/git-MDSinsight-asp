<!--#include file="../inc/header-field-service.asp"-->
<!--#include file="../inc/mail.asp"-->
<%
sURL = Request.ServerVariables("SERVER_NAME")
'baseURL should alwats have a trailing /slash, just in case, handle either way
If right(baseURL,1)="/" Then maildomain = Left(right(baseURL,len(baseURL)-7),len(right(baseURL,len(baseURL)-7))-1) Else maildomain = right(baseURL,len(baseURL)-7)

ReturnPath = Request.Form("txtReturnPathCloseCancel")

Account = Request.Form("txtCustIDCloseCancel")

CompanyName = GetCustNameByCustNum(Account)

ProblemCode = Request.Form("txtProblemCode")

If ProblemCode <> "" Then
	EquipmentProblemArray = Split(ProblemCode,"*")
	EquipmentProblemCodeIntRecID = EquipmentProblemArray(0)
	EquipmentProblemCodeDesc = EquipmentProblemArray(1)
End If

ResolutionCode = Request.Form("txtResolutionCode")

If ResolutionCode <> "" Then
	EquipmentResolutionArray = Split(ResolutionCode,"*")
	EquipmentResolutionCodeIntRecID = EquipmentResolutionArray(0)
	EquipmentResolutionCodeDesc = EquipmentResolutionArray(1)
End If

CloseOrCancelNotes = Request.Form("ServiceNotes")

'************************************************************************
'In subsequent versions, these fields are no longer retrieved.
'************************************************************************
'AssetTagNumber = Request.Form("txtAssetTagNumber")
'AssetLocation = Request.Form("txtAssetLocation")

AssetTagNumber = "" 
AssetLocation = "" 
'************************************************************************

sURL = Request.ServerVariables("SERVER_NAME")
PrintedName =  Request.Form("txtPrintedName")
CloseorCancel = Request.Form("optCloseOrCancel")
MemoNumber = Request.Form("txtMemoNumberCloseCancel")
FieldTech = Request.Form("selFieldTech")
CloseorCancelDate = Request.Form("txtCloseCancelDate")
CloseorCancelTime = Request.Form("txtCloseCancelTime")
DoNotEmail = Request.Form("chkDoNotEmail")

Response.Write("Account Start:<b>" & Account &"</b>:End<br>")
Response.Write("CloseorCancel Start:<b>" & CloseorCancel &"</b>:End<br>")
Response.Write("MemoNumber Start:<b>" & MemoNumber &"</b>:End<br>")
Response.Write("FieldTech Start:<b>" & FieldTech &"</b>:End<br>")
Response.Write("Stage Start:<b>" & GetServiceTicketCurrentStage(MemoNumber) &"</b>:End<br>")
'response.end




'Immediately write to Do Not Email If needed
If DoNotEmail = "on" Then	
	Set cnnDoNotEmail = Server.CreateObject("ADODB.Connection")
	cnnDoNotEmail.open Session("ClientCnnString")

	SQLDoNotEmail = "Insert into tblDoNotEmail (MemoNumber) VALUES ('" & MemoNumber & "')"

	Set rsDoNotEmail = Server.CreateObject("ADODB.Recordset")
	rsDoNotEmail.CursorLocation = 3 
	Set rsDoNotEmail = cnnDoNotEmail.Execute(SQLDoNotEmail)

	set rsDoNotEmail = Nothing
	set cnnDoNotEmail= Nothing

End If

'Response.Write("DoNotEmail:" &  DoNotEmail & ":XX")
'Response.End


DO_Post = 0
		
Do_Post = GetPOSTParams("AssetLocationURL1ONOFF") 
			
If IsNull(Do_Post) or Do_Post = "" Then Do_Post = 0

If cint(Do_Post) = 1 Then
		
	CreateINSIGHTAuditLogEntry sURL,"Asset Location Post Loop "& x,GetPOSTParams("Mode")

	'*********************************************
	'If we have asset information, post that first
	'*********************************************
	If AssetTagNumber <> "" And AssetLocation <> "" Then
	
		Description = Description & "     Account: "  & Account 
		Description = Description & ",    Asset Tag#: "  & AssetTagNumber
		Description = Description & ",    New Asset Location: "  & AssetLocation 
		Description = Description & ",   Submitted via MDS Insight"
		CreateAuditLogEntry "Asset Location Updated","Asset Location Updated","Major",0,Description
	
		data = "asset_id=" & AssetTagNumber 
		data = data & "&asset_loc=" & AssetLocation
		data = data & "&md=" & GetPOSTParams("Mode")
		data = data & "&serno=" & GetPOSTParams("Serno")
		data = data & "&src=MDS Insight"
		
		'Set httpRequest = Server.CreateObject("MSXML2.ServerXMLHTTP")
		
		'httpRequest.Open "POST", GetPOSTParams("AssetLocationURL1"), False

		'httpRequest.SetRequestHeader "Content-Type", "application/x-www-form-urlencoded"
		'httpRequest.Send data
	
		'Set httpRequest = Nothing
	
	End If
	'*************************************
	' End of posting asset location update 
	'*************************************
End If
		


If GetPOSTParams("SERVICEMEMOURL1MPLEXFORMAT") = 1 then ' only for metroplex for now

	CloseOrCancelNotes = Replace(CloseOrCancelNotes,"&","%26") 
	CloseOrCancelNotes = Replace(CloseOrCancelNotes," ","%20") & "%0A%0D"

	If AssetTagNumber <> "" And AssetLocation <> "" Then
		CloseOrCancelNotes = CloseOrCancelNotes & "Asset Location updated. Asset Tag#: " & AssetTagNumber & "  New Location:  " & AssetLocation 
		CloseOrCancelNotes = Replace(CloseOrCancelNotes,"&","%26") 
		CloseOrCancelNotes = Replace(CloseOrCancelNotes," ","%20") & "%0A%0D"
	End If

	If PrintedName <> "" AND CloseorCancel <> "Cancel" Then
		CloseOrCancelNotes = CloseOrCancelNotes & "Printed Name: " & PrintedName 
		CloseOrCancelNotes = Replace(CloseOrCancelNotes,"&","%26") 
		CloseOrCancelNotes = Replace(CloseOrCancelNotes," ","%20") & "%0A%0D"
	End If
End If




Description = "Service ticket #: " & MemoNumber & "-  "
If CloseorCancel = "Close" Then
	Description = Description & "Service call closed.    "
Else
	Select Case CloseorCancel
			Case "Cancel"
				Description = Description & "Service call cancelled.    "
	End Select
End If
Description = Description & "User: "  & 	MUV_Read("DisplayName") & " - " & Session("UserEmail")
Description = Description & ",     Account: "  & Account & " - " & Company 
CloseOrCancelNotesForAudit  = Replace(CloseOrCancelNotes,"%26","&")
CloseOrCancelNotesForAudit  = Replace(CloseOrCancelNotesForAudit  ,"%20"," ")
CloseOrCancelNotesForAudit  = Replace(CloseOrCancelNotesForAudit  ,"%0A%0D"," ")
Description = Description & ",     Service Notes: "  & CloseOrCancelNotesForAudit 
Description = Description & ",    Submitted via MDS Insight" & "%0A%0D" 
Description = Replace(Description ,"%20"," ")
If CloseorCancel = "Close" Then
	CreateAuditLogEntry "Service Call Closed","Service Call Closed","Major",0,Description
Else
	Select Case CloseorCancel
			Case "Cancel"
				CreateAuditLogEntry "Service Call Cancelled","Service Call Cancelled","Major",0,Description
	End Select
End If
Description = "Post to " & GetPOSTParams("ServiceMemoURL1")
CreateINSIGHTAuditLogEntry sURL,Description,GetPOSTParams("Mode")


'response.Write("<br><br>")
'response.write("data: " & data & "<br>")
'response.Write("<br><br>")
'response.end

'**************************************
'This is the post of the service ticket
'**************************************
Do_Post = 0
		
Do_Post = GetPOSTParams("ServiceMemoURL1ONOFF") 
		
If IsNull(Do_Post) or Do_Post = "" Then Do_Post = 0
	
If cint(Do_Post) = 1 Then
		
	CreateINSIGHTAuditLogEntry sURL,"Service Ticket Post Loop "& x,GetPOSTParams("Mode")

	If GetPOSTParams("SERVICEMEMOURL1MPLEXFORMAT") = 1 then 'On first post, see if using metroplex format
		
		CreateINSIGHTAuditLogEntry sURL,"Metroplex post format",GetPOSTParams("Mode")
						
		data = "create_service_request&account="
		data = data & Account
		If CloseorCancel = "Close" Then
			data = data & "&st=CLOSE"
		Elseif CloseorCancel = "Cancel" Then
			data = data & "&st=CANCEL"
		End If
		data = data & "&usr=" & FieldTech
		data = data & "&tnum=" & MemoNumber
		data = data & "&prob=" 
		data = data & CloseOrCancelNotes & "%0A%0D"
		If CloseorCancel = "Close" Then data = data & "Closed by: "  & 	GetUserDisplayNameByUserNo(FieldTech) & " - " & GetUserEmailByUserNo(FieldTech) & "%0A%0D"
		data = data & "Submitted via MDS Insight" & "%0A%0D"
		data = data & "&md=" &  GetPOSTParams("Mode")
		data = data & "&serno="  & GetPOSTParams("Serno")
		data = data & "&src=MDS Insight"

	
	Else ' Regular XML format

		CreateINSIGHTAuditLogEntry sURL,"Regular XML post format",GetPOSTParams("Mode")
			
		'Post to APIs goes here
		
		data = ""
		
		If CloseorCancel = "Cancel" Then
			data = data & "<POST_DATA>"
			data = data & "<IDENTITY>Pm8316wyc011</IDENTITY>"
			data = data & "<ACCOUNT_NUM>" & Account & "</ACCOUNT_NUM>"
			data = data & "<COMPANY_NAME>" & CompanyName & "</COMPANY_NAME>"
			data = data & "<PROBLEM_DESCRIPTION>" & CloseOrCancelNotes & "</PROBLEM_DESCRIPTION>"
			data = data & "<RECORD_TYPE>SENDSERVICEMSG</RECORD_TYPE>"
			data = data & "<RECORD_SUBTYPE>CANCEL</RECORD_SUBTYPE>"
			data = data & "<SERVICE_TICKET_NUMBER>" & MemoNumber & "</SERVICE_TICKET_NUMBER>"
			data = data & "<SUBMISSION_SOURCE>MDS Insight</SUBMISSION_SOURCE>"
			data = data & "<SERNO>" & GetPOSTParams("Serno") & "</SERNO>"
			data = data & "<CLIENT_ID>" & MUV_READ("ClientID") & "</CLIENT_ID>"
			data = data & "<USER_NO>" & FieldTech & "</USER_NO>"
			data = data & "<MODE>" & GetPOSTParams("Mode") & "</MODE>"
			data = data & "<PROBLEM_CODE>0</PROBLEM_CODE>"
			data = data & "<PROBLEM_CODE_DESC></PROBLEM_CODE_DESC>"
			data = data & "<RESOLUTION_CODE>0</RESOLUTION_CODE>"
			data = data & "<RESOLUTION_CODE_DESC></RESOLUTION_CODE_DESC>"
			data = data & "</POST_DATA>"
		Else
			data = data & "<POST_DATA>"
			data = data & "<IDENTITY>Pm8316wyc011</IDENTITY>"
			data = data & "<ACCOUNT_NUM>" & Account & "</ACCOUNT_NUM>"
			data = data & "<COMPANY_NAME>" & CompanyName & "</COMPANY_NAME>"
			data = data & "<PROBLEM_DESCRIPTION>" & CloseOrCancelNotes & "</PROBLEM_DESCRIPTION>"
			data = data & "<RECORD_TYPE>SENDSERVICEMSG</RECORD_TYPE>"
			data = data & "<RECORD_SUBTYPE>CLOSE</RECORD_SUBTYPE>"
			data = data & "<SERVICE_TICKET_NUMBER>" & MemoNumber & "</SERVICE_TICKET_NUMBER>"
			data = data & "<SUBMISSION_SOURCE>MDS Insight</SUBMISSION_SOURCE>"
			data = data & "<SERNO>" & GetPOSTParams("Serno") & "</SERNO>"
			data = data & "<CLIENT_ID>" & MUV_READ("ClientID") & "</CLIENT_ID>"
			data = data & "<USER_NO>" & FieldTech & "</USER_NO>"
			data = data & "<MODE>" & GetPOSTParams("Mode") & "</MODE>"
			data = data & "<PROBLEM_CODE>" & EquipmentProblemCodeIntRecID & "</PROBLEM_CODE>"
			data = data & "<PROBLEM_CODE_DESC>" & EquipmentProblemCodeDesc & "</PROBLEM_CODE_DESC>"
			data = data & "<RESOLUTION_CODE>" & EquipmentResolutionCodeIntRecID & "</RESOLUTION_CODE>"
			data = data & "<RESOLUTION_CODE_DESC>" & EquipmentResolutionCodeDesc & "</RESOLUTION_CODE_DESC>"
			data = data & "</POST_DATA>"
		End If
	
	End IF
	
	Description = "Post to " & GetPOSTParams("ServiceMemoURL1")
				
	CreateINSIGHTAuditLogEntry sURL,Description,GetPOSTParams("Mode")
	Description = "data:" & data 
	CreateINSIGHTAuditLogEntry sURL,Description,GetPOSTParams("Mode")
					
	Set httpRequest = Server.CreateObject("MSXML2.ServerXMLHTTP")
					
	httpRequest.Open "POST", GetPOSTParams("ServiceMemoURL1"), False
					
	'httpRequest.SetRequestHeader "Content-Type", "text/xml"
	httpRequest.SetRequestHeader "Content-Type", "application/x-www-form-urlencoded"
	
	httpRequest.Send data
						
	If httpRequest.status = 200 THEN 
					
		If Instr(httpRequest.responseText,"success") <> 0 Then ' Success
		
			Description ="success! httpRequest.status returned " & httpRequest.status & " when posting <RECORD_TYPE>SENDSERVICEMSG and <RECORD_SUBTYPE>CANCEL"& "<br>"
			Description = "httpRequest.responseText:" & httpRequest.responseText & "<br>"
			Description = Description & "PAGE: " & Request.ServerVariables("SERVER_NAME") & Request.ServerVariables("URL") & "<br>"
			If x = 1 Then
				Description = Description & "Posted to " & GetPOSTParams("ServiceMemoURL1") & "<br>"
			Else
				Description = Description & "Posted to " & GetPOSTParams("ServiceMemoURL2") & "<br>"
			End IF
			Description = Description & "POSTED DATA:" & data & "<br>"
			Description = Description & "SERNO:" & GetPOSTParams("Serno") & "<br>"
			Description = Description & "MODE:" & GetPOSTParams("Mode") & "<br>"
	
			CreateINSIGHTAuditLogEntry Request.ServerVariables("SERVER_NAME"),Description,GetPOSTParams("Mode")
	
		Else
			'FAILURE
			emailbody="httpRequest.status returned " & httpRequest.status & " when posting <RECORD_TYPE>SENDSERVICEMSG and <RECORD_SUBTYPE>CANCEL"& "<br>"
			emailBody = "httpRequest.responseText:" & httpRequest.responseText & "<br>"
			emailBody = emailBody & "PAGE: " & Request.ServerVariables("SERVER_NAME") & Request.ServerVariables("URL") & "<br>"
			If x = 1 Then
				emailBody = emailBody & "Posted to " & GetPOSTParams("ServiceMemoURL1") & "<br>"
			Else
				emailBody = emailBody  & "Posted to " & GetPOSTParams("ServiceMemoURL2") & "<br>"
			End IF
			emailBody = emailBody & "POSTED DATA:" & data & "<br>"
			emailBody = emailBody & "SERNO:" & GetPOSTParams("Serno") & "<br>"
			emailBody = emailBody & "MODE:" & GetPOSTParams("Mode") & "<br>"
							
			If Len(GetPOSTParams("EMAILFORNON200RESPONSES")) > 1 Then
				SendErrorToArray = Split(GetPOSTParams("EMAILFORNON200RESPONSES"),";")
				For z = 0 to Ubound(SendErrorToArray)
					If isEmailValid(SendErrorToArray(z)) = 0 Then 
						SendMail "mailsender@" & maildomain ,SendErrorToArray(z),MUV_READ("ClientID") & " ADD SERVICE MEMO POST ERROR",emailBody,GetTerm("Service"),"Post Failure"
					Else
						SendMail "mailsender@" & maildomain ,"rich@ocsaccess.com",MUV_READ("ClientID") & " ADD SERVICE MEMO POST ERROR",emailBody,GetTerm("Service"),"Post Failure"
					End If
				Next 
			Else
				SendMail "mailsender@" & maildomain ,"rich@ocsaccess.com",MUV_READ("ClientID") & " ADD SERVICE MEMO POST ERROR",emailBody,GetTerm("Service"),"Post Failure"
			End If
		
			Description = emailBody 
			CreateINSIGHTAuditLogEntry Request.ServerVariables("SERVER_NAME"),Description,GetPOSTParams("Mode")
		
		End If
		
	Else
						
		emailbody="httpRequest.status returned " & httpRequest.status & " when posting <RECORD_TYPE>SENDSERVICEMSG and <RECORD_SUBTYPE>CANCEL"& "<br>"
		emailBody = "httpRequest.responseText:" & httpRequest.responseText & "<br>"
		emailBody = emailBody & "PAGE: " & Request.ServerVariables("SERVER_NAME") & Request.ServerVariables("URL") & "<br>"
		If x = 1 Then
			emailBody = emailBody & "Posted to " & GetPOSTParams("ServiceMemoURL1") & "<br>"
		Else
			emailBody = emailBody  & "Posted to " & GetPOSTParams("ServiceMemoURL2") & "<br>"
		End IF
		emailBody = emailBody & "POSTED DATA:" & data & "<br>"
		emailBody = emailBody & "SERNO:" & GetPOSTParams("Serno") & "<br>"
		emailBody = emailBody & "MODE:" & GetPOSTParams("Mode") & "<br>"
	
		If Len(GetPOSTParams("EMAILFORNON200RESPONSES")) > 1 Then
			SendErrorToArray = Split(GetPOSTParams("EMAILFORNON200RESPONSES"),";")
			For z = 0 to Ubound(SendErrorToArray)
				If isEmailValid(SendErrorToArray(z)) = 0 Then 
					SendMail "mailsender@" & maildomain ,SendErrorToArray(z),MUV_READ("ClientID") & " ADD SERVICE MEMO POST ERROR",emailBody,GetTerm("Service"),"Post Failure"
				Else
					SendMail "mailsender@" & maildomain ,"rich@ocsaccess.com",MUV_READ("ClientID") & " ADD SERVICE MEMO POST ERROR",emailBody,GetTerm("Service"),"Post Failure"
				End If
			Next 
		Else
			SendMail "mailsender@" & maildomain ,"rich@ocsaccess.com",MUV_READ("ClientID") & " ADD SERVICE MEMO POST ERROR",emailBody,GetTerm("Service"),"Post Failure"
		End If
	
		Description = emailBody 
		CreateINSIGHTAuditLogEntry Request.ServerVariables("SERVER_NAME"),Description,GetPOSTParams("Mode")
						
	End If
	
	Set httpRequest = Nothing
	
End IF


'********************************************
'A d v a n c e d  D i s p a t c h For Cancels
'********************************************
If CloseorCancel = "Cancel" Then 
	'If advanced dispatch is on, determine how far we are in the stages
	If advancedDispatchIsOn() Then
		Select Case GetServiceTicketCurrentStage(MemoNumber)
			Case "Dispatched","Dispatch Acknowledged"
				' Force tech notification via text & email	
				Call SendCancelEmail
				Call SendCancelText	
			Case "En Route"
				' Force tech notification via text & email
				' and tell Joe to call
				If getUserCellNumber(GetServiceTicketDispatchedTech(MemoNumber)) <> "" Then 
					Session("MultiUseVar") = "You cancelled ticket # " & MemoNumber & ". " 
					Session("MultiUseVar") = Session("MultiUseVar") &  GetUserDisplayNameByUserNo(GetServiceTicketDispatchedTech(MemoNumber)) & " is En Route to this service call. "
					Session("MultiUseVar") = Session("MultiUseVar")  & "You should notify " &GetUserDisplayNameByUserNo(GetServiceTicketDispatchedTech(MemoNumber)) & " via phone to ensure the cancellation is received. "
					Session("MultiUseVar") = Session("MultiUseVar") & "Cell: " & getUserCellNumber(GetServiceTicketDispatchedTech(MemoNumber))
				End If
				Call SendCancelEmail
				Call SendCancelText	
		End Select
	End If
End If
'***************************************************
'E n d  A d v a n c e d  D i s p a t c h For Cancels
'***************************************************
	
'	Response.Write("<br>XX" &  data & "XX<br>")
'	Response.Write("<br><b>X" &  postResponse & "X</b><br>")
'	Response.end

If ReturnPath = "" Then
	Response.Redirect("main.asp")
ElseIf ReturnPath = "ServiceMain" Then
	Response.Redirect("main.asp")
ElseIf ReturnPath = "ServiceBoard" Then
	Response.Redirect("serviceBoard.asp")
ElseIf ReturnPath = "DispatchCenter" Then
	Response.Redirect("dispatchcenter/main.asp")
Else
	Response.Redirect("../customerservice/main.asp")
End IF





'***************************
'Functions & Subs begin here    
'***************************
Sub SendCancelEmail

	If getUserEmailAddress(GetServiceTicketDispatchedTech(MemoNumber)) <> "" Then
		Send_To = getUserEmailAddress(GetServiceTicketDispatchedTech(MemoNumber))
		CustNum = Account ' needed to keep emails consistent
		ServiceTicketNumber = MemoNumber ' needed to keep emails consistent%><!--#include file="../emails/ADVdispatch_cancel.asp"--><%	
		'Failsafe for dev
		If Instr(ucase(sURL),"DEV") <> 0 Then Send_To = "rich@ocsaccess.com"
		SendMail "mailsender@" & maildomain ,Send_To,emailSubject,emailBody,GetTerm("Service"),"Ticket Cancelled"
		Description = "A cancel dispatch email was sent to " & GetUserDisplayNameByUserNo(GetServiceTicketDispatchedTech(MemoNumber)) & " (" & Send_To & ") at " & NOW()
		CreateAuditLogEntry "Service Ticket System","Cancel dispatch email sent","Minor",0,Description
	Else
		' Could not send dispatch email, no address on file
		emailBody = "Insight was unable to send a cancel dispatch email to " & GetUserDisplayNameByUserNo(GetServiceTicketDispatchedTech(MemoNumber)) & ". No email address on file"
		If Instr(ucase(sURL),"DEV") <> 0 Then SEND_TO = "rich@ocsaccess.com" else SEND_TO = "rich@ocsaccess.com"
		SendMail "mailsender@" & maildomain ,SEND_TO,"Unable to send dispatch email",emailBody,GetTerm("Service"),"No email address on file"
		Description = "Insight was unable to send a cencel dispatch email to " & GetUserDisplayNameByUserNo(GetServiceTicketDispatchedTech(MemoNumber)) & ". No email address on file"
		CreateAuditLogEntry "Service Ticket System","Unable to send cancel dispatch email","Major",0,Description,GetTerm("Service"),"No email address on file"
	End If
	
	'If this was cancelled by someone other than a service manager, the service managers need to get emails as well
	'Get all the service manager email addresses
		'Fixit
	' cheap fix to let adam henchel see service stuff wihtout being a service manager


	If UserIsServiceManager(Session("UserNo")) <> True or UserNo=56 Then
		Set cnn_SvcMan = Server.CreateObject("ADODB.Connection")
		cnn_SvcMan.open (Session("ClientCnnString"))
		Set rs_SvcMan = Server.CreateObject("ADODB.Recordset")
		rs_SvcMan.CursorLocation = 3 
		SQL_SvcMan = "SELECT userEmail FROM tblUsers WHERE (userType = 'Service Manager' or userno=56) and userArchived <> 1" 
		Set rs_SvcMan = cnn_SvcMan.Execute(SQL_SvcMan)
		If not rs_SvcMan.EOF Then
			Do
				If rs_SvcMan("userEmail") <> "" AND Not IsNull(rs_SvcMan("userEmail")) Then Send_To = Send_To & rs_SvcMan("userEmail") & ";"
				rs_SvcMan.MoveNext
			Loop Until rs_SvcMan.Eof
		End If
		Set rs_SvcMan = Nothing
		cnn_SvcMan.Close
		Set cnn_SvcMan = Nothing
	End IF
	
	'Got all the addresses so now break them up
	Send_To_Array = Split(Send_To,";")

	For x = 0 to Ubound(Send_To_Array) -1
		Send_To = Send_To_Array(x)
		CustNum = Account ' needed to keep emails consistent
		ServiceTicketNumber = MemoNumber ' needed to keep emails consistent%><!--#include file="../emails/ADVdispatch_cancel.asp"--><%	
		'Failsafe for dev
		If Instr(ucase(sURL),"DEV") <> 0 Then Send_To = "rich@ocsaccess.com"
		SendMail "mailsender@" & maildomain ,Send_To,emailSubject,emailBody,GetTerm("Service"),"Ticket Cancelled"
		Description = "A cancel dispatch email was sent to " & GetUserDisplayNameByUserNo(GetServiceTicketDispatchedTech(MemoNumber)) & " (" & Send_To & ") at " & NOW()
		CreateAuditLogEntry "Service Ticket System","Cancel dispatch email sent","Minor",0,Description
	Next 

End Sub



Sub SendCancelText
	
		If getUserCellNumber(GetServiceTicketDispatchedTech(MemoNumber)) <> "" Then
			Send_To = getUserCellNumber(GetServiceTicketDispatchedTech(MemoNumber))
	
			URL = BaseURL & "inc/sendtext.php"
			QString = "?n=" & Replace(getUserCellNumber(GetServiceTicketDispatchedTech(MemoNumber)),"-","")
			
			QString = QString & "&u1=" & EzTextingUserID()
			QString = QString & "&u2=" & EzTextingPassword()

			QString = QString & "&t=CANCELLED"
			QString = QString & "&R=Location: " & Server.URLEncode(BaseURL & "/service/main.asp")
			QString = QString & "&m=" & GetTerm("Account") & ":" &  Account
			QString = QString &  "   Ticket:" & MemoNumber

			QString = QString &  "&cty=" & GetCompanyCountry()	
			QString = Replace(Qstring," ", "%20")
	
			Response.Redirect (URL & Qstring)
	
			Description = "A cancel dispatch text message was sent to " & GetUserDisplayNameByUserNo(GetServiceTicketDispatchedTech(MemoNumber)) & " (" & getUserCellNumber(GetServiceTicketDispatchedTech(MemoNumber)) & ") at " & NOW()
			CreateAuditLogEntry "Service Ticket System","Cancel dispatch email sent","Minor",0,Description
		Else
			' Could not send dispatch test, no address on file
			emailBody = "Insight was unable to send a cancel dispatch text message to " & GetUserDisplayNameByUserNo(GetServiceTicketDispatchedTech(MemoNumber)) & ". No cell number on file"
			If Instr(ucase(sURL),"DEV") <> 0 Then SEND_TO = "rich@ocsaccess.com" else SEND_TO = "rich@ocsaccess.com"
			SendMail "mailsender@" & maildomain ,SEND_TO,"Unable to send cancel dispatch text message",emailBody,GetTerm("Service"),"No email address on file"
			Description = "Insight was unable to send a cancel dispatch text message to " & GetUserDisplayNameByUserNo(GetServiceTicketDispatchedTech(MemoNumber)) & ". No cell number on file"
			CreateAuditLogEntry "Service Ticket System","Unable to send cancel dispatch text message","Major",0,Description
	
		End If
End Sub
 
%><!--#include file="../inc/footer-field-service-noTimeout.asp"-->