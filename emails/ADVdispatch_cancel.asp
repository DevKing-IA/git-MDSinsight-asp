<%
'****************************************************
'Create the email that goes to the service tech
'****************************************************

emailSubject = "CANCELLED - Ticket # " & ServiceTicketNumber & " - " & GetTerm("Account") & " # " & CustNum

emailBody = ""

emailBody =  emailBody & "<table width='650' border='0' cellspacing='0' align='center' style='padding:10px; border:1px solid #000000;'>"

emailBody =  emailBody & "<tr><td>"

emailBody =  emailBody & "<table width='100%' border='0' cellspacing='0' cellpadding='15'  >"

emailBody =  emailBody & "<tr><td width='650' style='font-family:Arial, Helvetica, sans-serif; font-size:21px; font-weight:normal; padding-top:15px; padding-bottom:15px; margin-left:3px margin-right:3px;' align='center'>*CANCELLATION*<br>Ticket # " & ServiceTicketNumber 

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


emailBody =  emailBody & "<table width='100%' border='0' cellspacing='0' cellpadding='15'  >"

emailBody =  emailBody & " <tr style='border-bottom:1px solid #666;' ><td width='40%' style='font-family:Arial, Helvetica, sans-serif; font-size:16px; font-weight:normal;'><strong>" & GetTerm("Account") & " #</strong></td><td width='60%' style='font-weight:normal; font-size:16px; font-family:Arial, Helvetica, sans-serif;' >"
emailBody =  emailBody & CustNum

emailBody =  emailBody & " <tr style='border-bottom:1px solid #666;' ><td width='40%' style='font-family:Arial, Helvetica, sans-serif; font-size:16px; font-weight:normal;'><strong>Company</strong></td><td width='60%' style='font-weight:normal; font-size:16px; font-family:Arial, Helvetica, sans-serif;' >"
emailBody =  emailBody & FormattedCustInfoByCustNum(CustNum)


emailBody =  emailBody & "</td></tr>"

emailBody =  emailBody & "</table>"

emailBody =  emailBody & "</td></tr>"


emailBody =  emailBody & "<tr><td>"

emailBody =  emailBody & "</td></tr>"
emailBody =  emailBody & "</table>"
%>