<%
'****************************************************
'Create the email that goes to the customer
'****************************************************

If ElapsedMinutes > 59 Then
	emailSubject = "HOLD ALERT Service Ticket " & rs100("MemoNumber") & " on hold over " & cInt(Round((ElapsedMinutes/60),2)) & " hours (" & MUV_READ("SERNO") & ")"
Else
	emailSubject = "HOLD ALERT Service Ticket " & rs100("MemoNumber") & " on hold over " & Round(ElapsedMinutes,2)-1 & " minutes (" & MUV_READ("SERNO") & ")"
End IF

 
emailBody = ""


emailBody =  emailBody & "<table width='650' border='0' cellspacing='0' align='center' style='padding:10px; border:1px solid #000000;'>"

emailBody =  emailBody & "<tr><td>"

emailBody =  emailBody & "<table width='100%' border='0' cellspacing='0' cellpadding='15'  >"

If ElapsedMinutes > 59 Then
	emailBody =  emailBody & "<tr><td width='650' style='font-family:Arial, Helvetica, sans-serif; font-size:21px; font-weight:normal; padding-top:15px; padding-bottom:15px; margin-left:3px margin-right:3px;' align='center'><font color='red'>*HOLD ALERT*</font><br>Service Ticket On Hold Over " & cInt(Round((ElapsedMinutes/60),2)) & " Hours <br>"
Else
	emailBody =  emailBody & "<tr><td width='650' style='font-family:Arial, Helvetica, sans-serif; font-size:21px; font-weight:normal; padding-top:15px; padding-bottom:15px; margin-left:3px margin-right:3px;' align='center'><font color='red'>*HOLD ALERT*</font><br>Service Ticket On Hold Over " & Round(ElapsedMinutes,2)-1 & " Minutes <br>"
End If

emailBody =  emailBody & "<font color='red'>On Hold</font>"

emailBody =  emailBody & "</tr></td>"

emailBody =  emailBody & "</table>"

emailBody =  emailBody & "</tr></td>"

emailBody =  emailBody & "<tr><td>"

emailBody =  emailBody & "<table width='100%' border='0' cellspacing='0' cellpadding='15'  >"

emailBody =  emailBody & "  <tr><th scope='col' width='40%' style='font-family:Arial, Helvetica, sans-serif; font-size:16px; font-weight:normal;' align='left'><strong>Submitted Date/Time</strong> </th> <th scope='col' width='60%' style='font-weight:normal; font-size:16px;' align='left'>"
emailBody =  emailBody & rs100.Fields("SubmissionDateTime") & " via " & rs100.Fields("SUBMISSIONSOURCE") & "</th></tr>"

emailBody =  emailBody & "  <tr><th scope='col' width='40%' style='font-family:Arial, Helvetica, sans-serif; font-size:16px; font-weight:normal;' align='left'><strong>Service Ticket #</strong> </th> <th scope='col' width='60%' style='font-weight:normal; font-size:16px;' align='left'>"
emailBody =  emailBody & rs100.Fields("MemoNumber") & "</th></tr>"


If rs100.Fields("ProblemLocation") <> "" Then 
	emailBody =  emailBody & " <tr ><td width='40%' style='font-family:Arial, Helvetica, sans-serif; font-size:16px; font-weight:normal;' valign='top'><strong>Location</strong></td><td width='60%' style='font-weight:normal; font-size:16px; font-family:Arial, Helvetica, sans-serif;' >"
	emailBody =  emailBody & rs100.Fields("ProblemLocation") & "</td></tr>"
End If

emailBody =  emailBody & " <tr ><td width='40%' style='font-family:Arial, Helvetica, sans-serif; font-size:16px; font-weight:normal;' valign='top'><strong>Description</strong></td><td width='60%' style='font-weight:normal; font-size:16px; font-family:Arial, Helvetica, sans-serif;' >"
emailBody =  emailBody & rs100.Fields("ProblemDescription") & "</td></tr>"

emailBody =  emailBody & " <tr style='border-bottom:1px solid #666;' ><td width='40%' style='font-family:Arial, Helvetica, sans-serif; font-size:16px; font-weight:normal;'><strong>Account#</strong></td><td width='60%' style='font-weight:normal; font-size:16px; font-family:Arial, Helvetica, sans-serif;' >"
emailBody =  emailBody & rs100.Fields("AccountNumber")

emailBody =  emailBody & " <tr style='border-bottom:1px solid #666;' ><td width='40%' style='font-family:Arial, Helvetica, sans-serif; font-size:16px; font-weight:normal;'><strong>Company</strong></td><td width='60%' style='font-weight:normal; font-size:16px; font-family:Arial, Helvetica, sans-serif;' >"
emailBody =  emailBody & FormattedCustInfoByCustNum(rs100.Fields("AccountNumber"))


emailBody =  emailBody & "</td></tr>"


emailBody =  emailBody & "</table>"

emailBody =  emailBody & "</td></tr>"


emailBody =  emailBody & "<tr><td>"

emailBody =  emailBody & "</td></tr>"
emailBody =  emailBody & "</table>"
 %>