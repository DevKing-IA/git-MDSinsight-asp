<!--#include file="inc/header-tech-and-driver.asp"-->
<!--#include file="../../inc/mail.asp"-->

<%


Set Upload = Server.CreateObject("Persits.Upload.1")
Upload.OverwriteFiles = False
Upload.Save
SelectedMemoNumber = Upload.Form("txtTicketNumber")


'Rename the files
' Construct the save path
Pth ="../../clientfiles/" & trim(GetPOSTParams("Serno")) & "/SvcMemoPics/"

x =1
For Each File in Upload.Files
   File.SaveAsVirtual  Pth & SelectedMemoNumber & "-" & x & File.Ext
   x=x+1
Next

Account = GetServiceTicketCust(SelectedMemoNumber)
ServiceNotes = Upload.Form("ServiceNotes")


'Might come from a dropdown or typed in
If Upload.Form("txtAssetTagNumber")<> "" Then AssetTagNumber = Upload.Form("txtAssetTagNumber")
If Upload.Form("selAssetID")<> "" Then AssetTagNumber = Upload.Form("selAssetID")

AssetLocation = Upload.Form("txtAssetLocation")
sURL = Request.ServerVariables("SERVER_NAME")
PrintedName =  Upload.Form("txtPrintedName")


'**************************************
'This is the post of the service ticket
'**************************************
		
	
If GetPOSTParams("ServiceMemoURL1ONOFF")  = 1 Then

		CreateSystemAuditLogEntry sURL,"Post Loop "& x,GetPOSTParams("Mode")

		If GetPOSTParams("SERVICEMEMOURL1MPLEXFORMAT") = 1 then ' see if using metroplex format

			CreateSystemAuditLogEntry sURL,"Metroplex post format",GetPOSTParams("Mode")
			
			data = "create_service_request&account="
			data = data & Account
			data = data & "&st=CLOSE"
			data = data & "&usr=" & Session("UserNo")
			data = data & "&tnum=" & SelectedMemoNumber 
			data = data & "&prob=" 
			data = data & CloseOrCancelNotes & "%0A%0D"
			data = data & "Closed by: "  & 	GetUserDisplayNameByUserNo(Session("UserNo")) & " - " & GetUserEmailByUserNo(Session("UserNo")) & "%0A%0D"
			data = data & "Submitted via MDS Insight" & "%0A%0D"
			data = data & "&md=" &  GetPOSTParams("Mode")
			data = data & "&serno="  & GetPOSTParams("Serno")
			data = data & "&src=MDS Insight"


		Else ' Regular XML format

			'Post to APIs goes here
			
			data = ""
			
			data = data & "<POST_DATA>"
			data = data & "<IDENTITY>Pm8316wyc011</IDENTITY>"
			data = data & "<ACCOUNT_NUM>" & Account & "</ACCOUNT_NUM>"
			data = data & "<PROBLEM_DESCRIPTION>" & CloseOrCancelNotes & "</PROBLEM_DESCRIPTION>"
			data = data & "<RECORD_TYPE>SENDSERVICEMSG</RECORD_TYPE>"
			data = data & "<RECORD_SUBTYPE>CLOSE</RECORD_SUBTYPE>"
			data = data & "<SERVICE_TICKET_NUMBER>" & SelectedMemoNumber & "</SERVICE_TICKET_NUMBER>"
			data = data & "<SUBMISSION_SOURCE>MDS Insight</SUBMISSION_SOURCE>"
			data = data & "<SERNO>" & GetPOSTParams("Serno") & "</SERNO>"
			data = data & "<CLIENT_ID>" & MUV_READ("ClientID") & "</CLIENT_ID>"
			data = data & "<USER_NO>" & Session("UserNo") & "</USER_NO>"
			data = data & "<MODE>" & GetPOSTParams("Mode") & "</MODE>"
			data = data & "</POST_DATA>"
				
		End IF

		Description = "Post to " & GetPOSTParams("ServiceMemoURL1")
		
		CreateSystemAuditLogEntry sURL,Description,GetPOSTParams("Mode")
		Description = "data:" & data 
		CreateSystemAuditLogEntry sURL,Description,GetPOSTParams("Mode")
		
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
		
				CreateSystemAuditLogEntry Request.ServerVariables("SERVER_NAME"),Description,GetPOSTParams("Mode")

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
				CreateSystemAuditLogEntry Request.ServerVariables("SERVER_NAME"),Description,GetPOSTParams("Mode")
			
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
			CreateSystemAuditLogEntry Request.ServerVariables("SERVER_NAME"),Description,GetPOSTParams("Mode")
							
		End If
		
		Set httpRequest = Nothing
End IF


Response.Redirect("main.asp")
	

%>