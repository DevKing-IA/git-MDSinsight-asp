<!--#include file="../inc/header-field-service.asp"-->
<!--#include file="../inc/mail.asp"-->
<!--#include file="../inc/InSightFuncs_Service.asp"-->
<%
sURL = Request.ServerVariables("SERVER_NAME")
'baseURL should alwats have a trailing /slash, just in case, handle either way
If right(baseURL,1)="/" Then maildomain = Left(right(baseURL,len(baseURL)-7),len(right(baseURL,len(baseURL)-7))-1) Else maildomain = right(baseURL,len(baseURL)-7)


PartNumber = Request.Form("txtPartNumber[]")
'Response.Write(PartNumber)
allPartNumber = split(PartNumber,",")

PartDescription = Request.Form("txtPartDescription[]")
'Response.Write(PartDescription)
allPartDescription = split(PartDescription,",")

PartQty = Request.Form("txtPartQty[]")
'Response.Write(PartQty)
allPartQty = split(PartQty,",")

PartNotes = Request.Form("txtPartNotes[]")
'Response.Write(PartNotes)
allPartNotes = split(PartNotes,",")

PartDesc = Request.Form("txtPartDesc")

requestData = "<table width='100%' border='1' cellspacing='5' cellpadding='5'>"
requestData = requestData & "<th>Part</th><th>Description</th><th>Quantity</th><th>Notes</th>"

For i = 0 to ubound(allPartQty)
	if trim(allPartQty(i)) <> "" then
		if allPartQty(i) > 0 then
			Response.Write(allPartNumber(i) & "<br>")
			Response.Write(allPartQty(i) & "<br>")
			Response.Write(allPartNotes(i) & "<br>")
			requestData = requestData & "<tr><td>" & allPartNumber(i) & "</td><td>" & allPartDescription(i) & "</td><td>" & allPartQty(i) & "</td><td>" & allPartNotes(i) & "</td></tr>"
		end if
	end if
Next

requestData = requestData & "<tr><td colspan=4>&nbsp;</td></tr>"
requestData = requestData & "<tr><td colspan=4><strong>Other Remark:</strong></td></tr>"
requestData = requestData & "<tr><td colspan=4>" & PartDesc & "</td></tr></table>"

'Response.Write(requestData)






emailBody = ""


emailBody =  emailBody & "<table width='650' border='0' cellspacing='0' align='center' style='padding:10px; border:1px solid #000000;'>"

emailBody =  emailBody & "<tr><td>"

emailBody =  emailBody & "<table width='650' border='0' cellspacing='0' cellpadding='15'  >"

emailBody =  emailBody & "<tr><td width='650' style='font-family:Arial, Helvetica, sans-serif; font-size:21px; font-weight:normal; padding-top:15px; padding-bottom:15px; margin-left:3px margin-right:3px;' align='center'>"
emailBody =  emailBody & "Parts request from " & GetUserDisplayNameByUserNo(Session("userNo")) 

emailBody =  emailBody & "</td></tr>"

emailBody =  emailBody & "</table>"

emailBody =  emailBody & "</tr></td>"

emailBody =  emailBody & "<tr><td>"

emailBody =  emailBody & requestData

emailBody =  emailBody & "</td></tr>"
emailBody =  emailBody & "</table>"

Response.Write(emailBody)

'emailBody = requestData

Call Send_Parts_Request_Email

Response.Redirect("main_menu.asp")



Sub Send_Parts_Request_Email

		Send_To="rsmith@ocsaccess.com"		
		'emailSubject="Part Request Detail"
		emailSubject = "Parts Request from " & GetUserDisplayNameByUserNo(Session("userNo"))
			'Response.Write("mailsender@" & maildomain & "','" & Send_To & "','" & emailSubject & "','" & emailBody &"','" & GetTerm("Field Service") & "','" & "'Part Request<br>'")

			SendMail "mailsender@" & maildomain,Send_To,emailSubject,emailBody,GetTerm("Field Service"),"Parts Request"
			
			Description = "A parts request was sent to " & Send_To & " at " & Now() 
			'Description = Description & " The text of the email was: " & ServiceNotes 
			CreateAuditLogEntry "Part Request","Part Request","Minor",0,Description
			
End Sub

%>















