<%
emailBody = ""

emailBody =  emailBody & "<table width='650' border='0' cellspacing='0' align='center' style='padding:10px; border:1px solid #000000;'>"

emailBody =  emailBody & "<tr><td>"

emailBody =  emailBody & "<table width='100%' border='0' cellspacing='0' cellpadding='15'  >"

emailBody =  emailBody & "<tr><td width='650' style='font-family:Arial, Helvetica, sans-serif; font-size:21px; font-weight:normal; padding-top:15px; padding-bottom:15px; margin-left:3px margin-right:3px;' align='center'>"

If passedNotificationType = "Alert" Then
	emailBody =  emailBody & "<font color='red'>*ALERT*</font><br>"
Else
	emailBody =  emailBody & "*Notification*<br>"
End If
	
emailBody =  emailBody & "Field Tech Service Notes from " & GetUserDisplayNameByUserNo(GetServiceTicketDispatchedTech(PassedMemoNumber))

emailBody =  emailBody & "</tr></td>"

emailBody =  emailBody & "</table>"


emailBody =  emailBody & "</tr></td>"

emailBody =  emailBody & "<tr><td>"

emailBody =  emailBody & "<table width='100%' border='0' cellspacing='0' cellpadding='15'  >"

emailBody =  emailBody & "  <tr><th scope='col' width='30%' style='font-family:Arial, Helvetica, sans-serif; font-size:14px; font-weight:normal;' align='left'>"
emailBody =  emailBody & "<font color='blue'><strong>Field Tech Notes:</strong></font></th>"
emailBody =  emailBody & " <th scope='col' width='70%' style='font-weight:normal; font-size:14px;' align='left'>"

'ServiceNotesForEmail comes from a global var dimmed in the top oft he directlaunch check page

emailBody =  emailBody & "<font color='blue'><strong>" & ServiceNotesForEmail & "</strong></font></th></tr>"


emailBody =  emailBody & "  <tr><th scope='col' width='30%' style='font-family:Arial, Helvetica, sans-serif; font-size:14px; font-weight:normal;' align='left'><strong>Status / Stage:</strong></th>"
emailBody =  emailBody & " <th scope='col' width='70%' style='font-weight:normal; font-size:14px;' align='left'>"


emailBody =  emailBody & GetServiceTicketCurrentStage(PassedMemoNumber) & " at " & GetServiceTicketSTAGEDateTime(PassedMemoNumber,GetServiceTicketCurrentStage(PassedMemoNumber)) & "</th></tr>"

emailBody =  emailBody & "  <tr><th scope='col' width='30%' style='font-family:Arial, Helvetica, sans-serif; font-size:14px; font-weight:normal;' align='left'><strong>Technician:</strong></th>"
emailBody =  emailBody & " <th scope='col' width='70%' style='font-weight:normal; font-size:14px;' align='left'>"
emailBody =  emailBody &  GetUserDisplayNameByUserNo(GetServiceTicketDispatchedTech(PassedMemoNumber)) & "</th></tr>"

emailBody =  emailBody & "  <tr><th scope='col' width='30%' style='font-family:Arial, Helvetica, sans-serif; font-size:14px; font-weight:normal;' align='left'><strong>Submitted Date/Time:</strong> </th> <th scope='col' width='70%' style='font-weight:normal; font-size:14px;' align='left'>"
emailBody =  emailBody & GetServiceTicketSubmissionDateTimeByTicketNumber(PassedMemoNumber) & " via " & GetServiceTicketSubmissionSourceByTicketNumber(PassedMemoNumber) & "</th></tr>"

emailBody =  emailBody & "  <tr><th scope='col' width='30%' style='font-family:Arial, Helvetica, sans-serif; font-size:14px; font-weight:normal;' align='left'><strong>Service Ticket:</strong> </th> <th scope='col' width='70%' style='font-weight:normal; font-size:14px;' align='left'>"
emailBody =  emailBody & PassedMemoNumber & "</th></tr>"

If GetServiceTicketProblemLocationByTicketNumber(PassedMemoNumber) <> "" Then 
	emailBody =  emailBody & " <tr ><td width='30%' style='font-family:Arial, Helvetica, sans-serif; font-size:14px; font-weight:normal;' valign='top'><strong>Location:</strong></td><td width='70%' style='font-weight:normal; font-size:14px; font-family:Arial, Helvetica, sans-serif;' >"
	emailBody =  emailBody & GetServiceTicketProblemLocationByTicketNumber(PassedMemoNumber) & "</td></tr>"
End If

emailBody =  emailBody & " <tr ><td width='30%' style='font-family:Arial, Helvetica, sans-serif; font-size:14px; font-weight:normal;' valign='top'><strong>Original Problem:</strong></td><td width='70%' style='font-weight:normal; font-size:14px; font-family:Arial, Helvetica, sans-serif;' >"
emailBody =  emailBody & GetServiceTicketProblemByTicketNumber(PassedMemoNumber) & "</td></tr>"

emailBody =  emailBody & " <tr style='border-bottom:1px solid #666;' ><td width='30%' style='font-family:Arial, Helvetica, sans-serif; font-size:14px; font-weight:normal;'><strong>Account:</strong></td><td width='70%' style='font-weight:normal; font-size:14px; font-family:Arial, Helvetica, sans-serif;' >"
emailBody =  emailBody & GetServiceTicketCust(PassedMemoNumber)

emailBody =  emailBody & " <tr style='border-bottom:1px solid #666;' ><td width='30%' style='font-family:Arial, Helvetica, sans-serif; font-size:14px; font-weight:normal;'><strong>Company:</strong></td><td width='70%' style='font-weight:normal; font-size:14px; font-family:Arial, Helvetica, sans-serif;' >"
emailBody =  emailBody & FormattedCustInfoByCustNum(GetServiceTicketCust(PassedMemoNumber))

emailBody =  emailBody & "</td></tr>"

emailBody =  emailBody & "</table>"

emailBody =  emailBody & "</td></tr>"

emailBody =  emailBody & "<tr><td>"

emailBody =  emailBody & "(" & rsAlert("AlertName") & ":" & rsAlert("InternalAlertRecNumber") & ":" & ClientKey & ")"

emailBody =  emailBody & "</td></tr>"
emailBody =  emailBody & "</table>"
 %>