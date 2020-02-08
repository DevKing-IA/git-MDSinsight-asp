<!--#include file="../inc/subsandfuncs.asp"-->
<!--#include file="../inc/InSightFuncs.asp"-->
<!--#include file="../inc/mail.asp"-->
<%

EquipmentType = Request.Form("txtEquipmentType")
EquipmentSymptom = Request.Form("txtEquipmentSymptom")

EquipmentSymptomArray = Split(EquipmentSymptom,"*")
EquipmentSymptomCodeIntRecID = EquipmentSymptomArray(0)
EquipmentSymptomCodeDesc = EquipmentSymptomArray(1)

txtDateProblemStarted = Request.Form("txtDateProblemStarted")
AccountNumber = Request.Form("txtAccount")
Company = Request.Form("txtCompany")
ContactName = Request.Form("txtWhoToContactUponArrival")
ContactPhone = Request.Form("txtContactPhone")
ContactEmail = Request.Form("txtContactEmail")
ProblemLocation = Request.Form("txtFloorSuite")
ProblemDescription = Request.Form("txtDescription")
SumissionSource = "MDS Insight"
sURL = Request.ServerVariables("SERVER_NAME")
'baseURL should always have a trailing /slash, just in case, handle either way
If right(baseURL,1)="/" Then maildomain = Left(right(baseURL,len(baseURL)-7),len(right(baseURL,len(baseURL)-7))-1) Else maildomain = right(baseURL,len(baseURL)-7)



' Write Audit trail first, then post
Set cnn8 = Server.CreateObject("ADODB.Connection")
cnn8.open (Session("ClientCnnString"))


Description = "OPEN,    "
Description = Description & "Account: "  & AccountNumber & " - " & Company 
Description = Description & ",     Location: "  & ProblemLocation 
Description = Description & ",    Description: "  & ProblemDescription 
CreateAuditLogEntry "Service Memo Added","Service Memo Added","Major",0,Description




DO_Post = 0
		
Do_Post = GetPOSTParams("ServiceMemoURL1ONOFF") 
		
If IsNull(Do_Post) or Do_Post = "" Then Do_Post = 0
	
If cint(Do_Post) = 1 Then
		
	CreateINSIGHTAuditLogEntry sURL,"Post Loop "& x,GetPOSTParams("Mode")
		
	If GetPOSTParams("SERVICEMEMOURL1MPLEXFORMAT") = 1 then 'see if using metroplex format
		
		CreateINSIGHTAuditLogEntry sURL,"Metroplex post format",GetPOSTParams("Mode")
		
		'Post to APIs goes here
		postmessage=Replace(ProblemDescription," ","%20")
		postmessage= Replace(postmessage,"&","%26") 
		data = "create_service_request&account="
		data = data & AccountNumber
		data = data & "&serno=" & GetPOSTParams("Serno")
		data = data & "&prob=" &  postmessage  & "%0A%0D"
		data = data & "Opened by: " & 	MUV_Read("DisplayName") & " - " & Session("UserEmail")
		data = data & "&probloc=" & ProblemLocation 
		data = data & "&sbn=" & ContactName 
		data = data & "&sbp=" & ContactPhone 
		data = data & "&sbe=" & "rsmith@ocsaccess.com"
		data = data & "&contnm=" & ContactName
		data = data & "&st=" & "OPEN"
		data = data & "&md=" & GetPOSTParams("Mode")
				
		data = data & "&src=" & SumissionSource
		data = data & "&usr=" & Session("userNo")
		If GetPOSTParams("NeverPutOnHold") = 0  Then data = data & "&hld=1" Else data = data & "&hld=0" 'I know it's wierd. It's the opposite of how it is stored in the table
				
	Else ' Regular XML format

		'Post to APIs goes here
				
		data = ""
				
		data = data & "<POST_DATA>"
		data = data & "<IDENTITY>Pm8316wyc011</IDENTITY>"
		data = data & "<ACCOUNT_NUM>" & AccountNumber & "</ACCOUNT_NUM>"
		data = data & "<PROBLEM_LOCATION>" & ProblemLocation & "</PROBLEM_LOCATION>"
		data = data & "<PROBLEM_DESCRIPTION>" & ProblemDescription & "</PROBLEM_DESCRIPTION>"
		data = data & "<RECORD_TYPE>SENDSERVICEMSG</RECORD_TYPE>"
		data = data & "<RECORD_SUBTYPE>OPEN</RECORD_SUBTYPE>"
		data = data & "<SERVICE_TICKET_NUMBER>AUTO</SERVICE_TICKET_NUMBER>"
		data = data & "<COMPANY_NAME>" & Replace(GetCustNameByCustNum(AccountNumber),"&","&amp;") & "</COMPANY_NAME>"
		data = data & "<SUBMISSION_SOURCE>MDS Insight</SUBMISSION_SOURCE>"
		data = data & "<SERNO>" & GetPOSTParams("Serno") & "</SERNO>"
		data = data & "<CLIENT_ID>" & MUV_READ("ClientID") & "</CLIENT_ID>"
		data = data & "<USER_NO>" & Session("UserNo") & "</USER_NO>"
		data = data & "<MODE>" & GetPOSTParams("Mode") & "</MODE>"
		data = data & "<SUBMITTED_BY_PHONE>" & ContactPhone & "</SUBMITTED_BY_PHONE>"
		data = data & "<SUBMITTED_BY_EMAIL>" & ContactEmail & "</SUBMITTED_BY_EMAIL>"
		data = data & "<SUBMITTED_BY_NAME>" & ContactEmail & "</SUBMITTED_BY_NAME>"
		data = data & "<SYMPTOM_CODE>" & EquipmentSymptomCodeIntRecID & "</SYMPTOM_CODE>"
		data = data & "<SYMPTOM_CODE_DESC>" & EquipmentSymptomCodeDesc & "</SYMPTOM_CODE_DESC>"	
		data = data & "</POST_DATA>"
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
					
			Description ="success! httpRequest.status returned " & httpRequest.status & " when posting <RECORD_TYPE>SENDSERVICEMSG and <RECORD_SUBTYPE>OPEN"& "<br>"
			Description = "httpRequest.responseText:" & httpRequest.responseText & "<br>"
			Description = Description & "PAGE: " & Request.ServerVariables("SERVER_NAME") & Request.ServerVariables("URL") & "<br>"
			Description = Description & "Posted to " & GetPOSTParams("ServiceMemoURL1") & "<br>"
			Description = Description & "POSTED DATA:" & data & "<br>"
			Description = Description & "SERNO:" & GetPOSTParams("Serno") & "<br>"
			Description = Description & "MODE:" & GetPOSTParams("Mode") & "<br>"
	
			CreateINSIGHTAuditLogEntry Request.ServerVariables("SERVER_NAME"),Description,GetPOSTParams("Mode")

		Else
			'FAILURE
			emailbody="httpRequest.status returned " & httpRequest.status & " when posting <RECORD_TYPE>SENDSERVICEMSG and <RECORD_SUBTYPE>OPEN"& "<br>"
			emailBody = "httpRequest.responseText:" & httpRequest.responseText & "<br>"
			emailBody = emailBody & "PAGE: " & Request.ServerVariables("SERVER_NAME") & Request.ServerVariables("URL") & "<br>"
			emailBody = emailBody & "Posted to " & GetPOSTParams("ServiceMemoURL1") & "<br>"
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
					
		emailbody="httpRequest.status returned " & httpRequest.status & " when posting <RECORD_TYPE>SENDSERVICEMSG and <RECORD_SUBTYPE>OPEN"& "<br>"
		emailBody = "httpRequest.responseText:" & httpRequest.responseText & "<br>"
		emailBody = emailBody & "PAGE: " & Request.ServerVariables("SERVER_NAME") & Request.ServerVariables("URL") & "<br>"
		emailBody = emailBody & "Posted to " & GetPOSTParams("ServiceMemoURL1") & "<br>"
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


Response.Redirect("main.asp")	

'Response.Write("<br>Posted To: " &  GetPOSTParams("ServiceMemoURL1") & "<br>")	
'Response.Write("<br>XX" &  data & "XX<br>")
'Response.Write("<br>XX" &  postResponse & "XX<br>")
'Response.end
	


%>















