<%
'****************************************************
'Create the email that goes to the service tech
'****************************************************

emailSubject = "Parts Request from " & GetUserDisplayNameByUserNo(GetServiceTicketDispatchedTech(SelectedMemoNumber))
emailSubject = emailSubject & " - Service ticket: " & SelectedMemoNumber & " - " & GetTerm("Account") & ": " & Account

 
emailBody = ""


emailBody =  emailBody & "<table width='650' border='0' cellspacing='0' align='center' style='padding:10px; border:1px solid #000000;'>"

emailBody =  emailBody & "<tr><td>"

emailBody =  emailBody & "<table width='650' border='0' cellspacing='0' cellpadding='15'  >"

emailBody =  emailBody & "<tr><td width='650' style='font-family:Arial, Helvetica, sans-serif; font-size:21px; font-weight:normal; padding-top:15px; padding-bottom:15px; margin-left:3px margin-right:3px;' align='center'>"
emailBody =  emailBody & "Parts request from " & GetUserDisplayNameByUserNo(GetServiceTicketDispatchedTech(SelectedMemoNumber)) & "<br>Service ticket: " & SelectedMemoNumber & " - " & GetTerm("Account") & ": " & Account

emailBody =  emailBody & "</td></tr>"

emailBody =  emailBody & "</table>"

emailBody =  emailBody & "</tr></td>"

emailBody =  emailBody & "<tr><td>"

emailBody =  emailBody & "<table width='100%' border='0' cellspacing='0' cellpadding='15'  >"

emailBody =  emailBody & "  <tr><th scope='col' width='40%' style='font-family:Arial, Helvetica, sans-serif; font-size:16px; font-weight:normal;' align='left'><strong>Parts Request: </strong> </th> <th scope='col' width='60%' style='font-weight:normal; font-size:16px;' align='left'>"
emailBody =  emailBody & ServiceNotes & "</th></tr>"

emailBody =  emailBody & "  <tr><th scope='col' width='40%' style='font-family:Arial, Helvetica, sans-serif; font-size:16px; font-weight:normal;' align='left'><strong>Service Ticket: </strong> </th> <th scope='col' width='60%' style='font-weight:normal; font-size:16px;' align='left'>"
emailBody =  emailBody & SelectedMemoNumber & "</th></tr>"

If GetServiceTicketProblemLocationByTicketNumber(SelectedMemoNumber) <> "" Then 
	emailBody =  emailBody & " <tr ><td width='40%' style='font-family:Arial, Helvetica, sans-serif; font-size:16px; font-weight:normal;' valign='top'><strong>Location</strong></td><td width='60%' style='font-weight:normal; font-size:16px; font-family:Arial, Helvetica, sans-serif;' >"
	emailBody =  emailBody & GetServiceTicketProblemLocationByTicketNumber(PassedMemoNumber) & "</td></tr>"
End If

emailBody =  emailBody & " <tr ><td width='40%' style='font-family:Arial, Helvetica, sans-serif; font-size:16px; font-weight:normal;' valign='top'><strong>Original Problem: </strong></td><td width='60%' style='font-weight:normal; font-size:16px; font-family:Arial, Helvetica, sans-serif;' >"
emailBody =  emailBody & GetServiceTicketProblemByTicketNumber(SelectedMemoNumber) & "</td></tr>"

emailBody =  emailBody & " <tr style='border-bottom:1px solid #666;' ><td width='40%' style='font-family:Arial, Helvetica, sans-serif; font-size:16px; font-weight:normal;'><strong>Account: </strong></td><td width='60%' style='font-weight:normal; font-size:16px; font-family:Arial, Helvetica, sans-serif;' >"
emailBody =  emailBody & GetServiceTicketCust(SelectedMemoNumber)

emailBody =  emailBody & " <tr style='border-bottom:1px solid #666;' ><td width='40%' style='font-family:Arial, Helvetica, sans-serif; font-size:16px; font-weight:normal;'><strong>Company :</strong></td><td width='60%' style='font-weight:normal; font-size:16px; font-family:Arial, Helvetica, sans-serif;' >"
emailBody =  emailBody & FormattedCustInfoByCustNum(GetServiceTicketCust(SelectedMemoNumber))

emailBody =  emailBody & "</td></tr>"

emailBody =  emailBody & "</table>"

emailBody =  emailBody & "</td></tr>"

emailBody =  emailBody & "<tr><td>"
emailBody =  emailBody & "(" & MUV_READ("SERNO") & ")"
emailBody =  emailBody & "</td></tr>"
emailBody =  emailBody & "</table>"
%>