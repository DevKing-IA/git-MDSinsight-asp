<%
'****************************************************
'Create the email that goes to the service tech
'****************************************************

If GetCustNameByCustNum(CustNum) <> "" Then
	emailSubject = "DISPATCH - " & GetCustNameByCustNum(CustNum)
Else
	emailSubject = "DISPATCH - Ticket # " & ServiceTicketNumber & " - " & GetTerm("Account") & " # " & CustNum
End If

 
emailBody = ""


emailBody =  emailBody & "<table width='650' border='0' cellspacing='0' align='center' style='padding:10px; border:1px solid #000000;'>"

emailBody =  emailBody & "<tr><td>"

emailBody =  emailBody & "<table width='100%' border='0' cellspacing='0' cellpadding='15'  >"

emailBody =  emailBody & "<tr><td width='650' style='font-family:Arial, Helvetica, sans-serif; font-size:21px; font-weight:normal; padding-top:15px; padding-bottom:15px; margin-left:3px margin-right:3px;' align='center'>*DISPATCH*<br>Ticket # " & ServiceTicketNumber & " - " & GetTerm("Account") & " # " & CustNum

emailBody =  emailBody & "</tr></td>"

emailBody =  emailBody & "</table>"

emailBody =  emailBody & "</tr></td>"

emailBody =  emailBody & "<tr><td>"

'Lookup the details of the ticket
SQLeml = "SELECT * FROM FS_ServiceMemos where MemoNumber = '" & ServiceTicketNumber & "'"
Set cnneml = Server.CreateObject("ADODB.Connection")
cnneml.open (Session("ClientCnnString"))
Set rseml = Server.CreateObject("ADODB.Recordset")
rseml.CursorLocation = 3 
Set rseml = cnneml.Execute(SQLeml)
If not rseml.EOF Then
	CurrentStatus = rseml("CurrentStatus")
	RecordSubType = rseml("RecordSubType")
	SubmittedByName = rseml("SubmittedByName")
	Company = rseml("Company")
	ProblemLocation = rseml("ProblemLocation")
	SubmittedByPhone = rseml("SubmittedByPhone")
	SubmittedByEmail = rseml("SubmittedByEmail")
	SubmissionDateTime = rseml("SubmissionDateTime")
	ProblemDescription = rseml("ProblemDescription")
	Mode = rseml("Mode")
	SubmissionSource = rseml("SubmissionSource")
	UserNoOfServiceTech = rseml("UserNoOfServiceTech")
	ReleasedDateTime = rseml("ReleasedDateTime")
	ReleasedByUserNo = rseml("ReleasedByUserNo")
	ReleasedNotes = rseml("ReleasedNotes")
End If
set rseml = Nothing
cnneml.close
set cnneml = Nothing

'See if ACK link should be included
DLinkInEmail = False
SQLeml = "SELECT  IncludeACKInDispatchEmail FROM Settings_EmailService"
Set cnneml = Server.CreateObject("ADODB.Connection")
cnneml.open (Session("ClientCnnString"))
Set rseml = Server.CreateObject("ADODB.Recordset")
rseml.CursorLocation = 3 
Set rseml = cnneml.Execute(SQLeml)
If not rseml.EOF Then DLinkInEmail = rseml("IncludeACKInDispatchEmail")
set rseml = Nothing
cnneml.close
set cnneml = Nothing


emailBody =  emailBody & "<table width='100%' border='0' cellspacing='0' cellpadding='15'  >"

emailBody =  emailBody & " <tr style='border-bottom:1px solid #666;' ><td width='40%' style='font-family:Arial, Helvetica, sans-serif; font-size:16px; font-weight:normal;'><strong>" & GetTerm("Account") & " #</strong></td><td width='60%' style='font-weight:normal; font-size:16px; font-family:Arial, Helvetica, sans-serif;' >"
emailBody =  emailBody & CustNum

emailBody =  emailBody & " <tr style='border-bottom:1px solid #666;' ><td width='40%' style='font-family:Arial, Helvetica, sans-serif; font-size:16px; font-weight:normal;'><strong>Company</strong></td><td width='60%' style='font-weight:normal; font-size:16px; font-family:Arial, Helvetica, sans-serif;' >"
emailBody =  emailBody & FormattedCustInfoByCustNum(CustNum)


emailBody =  emailBody & "</td></tr>"

If DLinkInEmail = 1 Then
	If FS_TechCanDecline() Then
		emailBody =  emailBody & " <tr ><td width='40%' style='font-family:Arial, Helvetica, sans-serif; font-size:16px; font-weight:normal;' valign='top'><strong>Acknowldge</strong></td><td width='60%' style='font-weight:normal; font-size:16px; font-family:Arial, Helvetica, sans-serif;' >"
		emailBody =  emailBody & "To viem more information, decline or acknowledge this ticket <a href='" & baseURL & "directlaunch/service/moreinfo_dispatch_from_email_or_text.asp?t=" & ServiceTicketNumber & "&u=" & UserToDispatch & "&c=" & CustNum & "&cl=" & MUV_READ("SERNO") & "'>click this link</a>"
		emailBody =  emailBody & "</td></tr>"
	Else
		emailBody =  emailBody & " <tr ><td width='40%' style='font-family:Arial, Helvetica, sans-serif; font-size:16px; font-weight:normal;' valign='top'><strong>Acknowldge</strong></td><td width='60%' style='font-weight:normal; font-size:16px; font-family:Arial, Helvetica, sans-serif;' >"
		emailBody =  emailBody & "To acknowledge this dispatch notification <a href='" & baseURL & "directlaunch/service/Ack_dispatch_from_email_or_text.asp?t=" & ServiceTicketNumber & "&u=" & UserToDispatch & "&c=" & CustNum & "&cl=" & MUV_READ("SERNO") & "'>click this link</a>"
		emailBody =  emailBody & "</td></tr>"
	End If
End If

emailBody =  emailBody & "</table>"

emailBody =  emailBody & "</td></tr>"


emailBody =  emailBody & "<tr><td>"

emailBody =  emailBody & "</td></tr>"
emailBody =  emailBody & "</table>"
%>